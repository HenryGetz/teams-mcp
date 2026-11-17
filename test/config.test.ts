import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { config, ConfigError, parseConfigYaml, resetConfigCache } from "../src/config.js";

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-config-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  delete process.env.TEAMS_MCP_CONFIG_FILE;
  delete process.env.TEAMS_MCP_TIMEZONE;
  delete process.env.TEAMS_MCP_CACHE_DIR;
  delete process.env.TEAMS_MCP_SNAPSHOT_DIR;
  delete process.env.TEAMS_MCP_DEFAULT_DURATION_MIN;
  resetConfigCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("parseConfigYaml", () => {
  it("fills defaults for an empty document", () => {
    const parsed = parseConfigYaml("");
    expect(parsed.timezone).toBe("America/Chicago");
    expect(parsed.work_hours).toEqual({ days: ["mon", "tue", "wed", "thu", "fri"], start: "09:00", end: "17:00" });
    expect(parsed.default_duration_min).toBe(30);
    expect(parsed.cache.busy_ttl_minutes).toBe(30);
    expect(parsed.aliases).toEqual({});
  });

  it("accepts a full document", () => {
    const parsed = parseConfigYaml(`
timezone: America/New_York
work_hours: {days: [mon, wed], start: "08:00", end: "16:30"}
default_duration_min: 45
aliases: {joel: joel@example.com}
invite_allowlist: ["@example.com"]
blocked_times: [{start: "12:00", end: "13:00", reason: lunch}]
test_event_prefix: "[TEST]"
max_results: {mail: 10, slots: 3}
`);
    expect(parsed.work_hours.days).toEqual(["mon", "wed"]);
    expect(parsed.aliases.joel).toBe("joel@example.com");
    expect(parsed.blocked_times[0]?.reason).toBe("lunch");
    expect(parsed.max_results.mail).toBe(10);
    expect(parsed.max_results.calendar).toBe(50);
  });

  it("rejects an invalid clock time with a named issue", () => {
    expect(() => parseConfigYaml("work_hours: {start: 9am}")).toThrow(ConfigError);
    try {
      parseConfigYaml("work_hours: {start: 9am}");
    } catch (error) {
      expect((error as ConfigError).issues.join(" ")).toContain("work_hours.start");
    }
  });

  it("rejects out-of-range caps", () => {
    expect(() => parseConfigYaml("max_results: {mail: 500}")).toThrow(ConfigError);
  });

  it("reports YAML syntax errors", () => {
    expect(() => parseConfigYaml("timezone: [unclosed")).toThrow(/YAML parse error|invalid config/);
  });
});

describe("config()", () => {
  it("applies defaults without a file", () => {
    const cfg = config();
    expect(cfg.timezone).toBe("America/Chicago");
    expect(cfg.snapshotPath.endsWith("calendar.yaml")).toBe(true);
    expect(cfg.tokenPath.endsWith("token.json")).toBe(true);
  });

  it("reads config.yaml and converts TTL minutes to seconds", async () => {
    await writeFile(
      join(dir, "config.yaml"),
      "timezone: America/Denver\ncache: {own_calendar_ttl_minutes: 2, busy_ttl_minutes: 45}\n",
    );
    const cfg = config();
    expect(cfg.timezone).toBe("America/Denver");
    expect(cfg.cache.ownCalendarTtlSeconds).toBe(120);
    expect(cfg.cache.busyTtlSeconds).toBe(2700);
  });

  it("lets environment variables win over the file", async () => {
    await writeFile(join(dir, "config.yaml"), "timezone: America/Denver\ndefault_duration_min: 45\n");
    process.env.TEAMS_MCP_TIMEZONE = "Europe/Berlin";
    process.env.TEAMS_MCP_DEFAULT_DURATION_MIN = "15";
    const cfg = config();
    expect(cfg.timezone).toBe("Europe/Berlin");
    expect(cfg.defaultDurationMin).toBe(15);
  });

  it("hot-reloads when the file changes", async () => {
    await writeFile(join(dir, "config.yaml"), "default_duration_min: 20\n");
    expect(config().defaultDurationMin).toBe(20);
    await writeFile(join(dir, "config.yaml"), "default_duration_min: 60\n");
    expect(config().defaultDurationMin).toBe(60);
  });

  it("does not reuse the memo across config dirs when neither has a file", async () => {
    // Both dirs are in the same "no config.yaml" state; the resolved paths must
    // still follow the environment, or a second process/in-process switch would
    // read and write the first dir's token.json.
    const first = config();
    const other = await mkdtemp(join(tmpdir(), "teams-mcp-config-other-"));
    process.env.TEAMS_MCP_CONFIG_DIR = other;
    try {
      const second = config();
      expect(second.tokenPath).toBe(join(other, "token.json"));
      expect(second.tokenPath).not.toBe(first.tokenPath);
      expect(second.profileDir).not.toBe(first.profileDir);
    } finally {
      await rm(other, { recursive: true, force: true });
    }
  });

  it("reports the file path and honours TEAMS_MCP_CACHE_DIR", () => {
    process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
    const cfg = config();
    expect(cfg.configFile).toBe(join(dir, "config.yaml"));
    expect(cfg.snapshotDir).toBe(join(dir, "cache", "snapshots"));
  });

  it("uses browser-profile and Outlook REST v2 defaults", () => {
    const cfg = config();
    expect(cfg.profileDir).toBe(join(dir, "browser-profile"));
    expect(cfg.outlookBaseUrl).toBe("https://outlook.office.com/api/v2.0");
    expect(cfg.graphBaseUrl).toBe("https://graph.microsoft.com/v1.0");
    expect(cfg.tokenPath).toBe(join(dir, "token.json"));
  });

  it("defaults the Teams send allowlist to empty and reads it from YAML", () => {
    expect(config().teams.allowedTargets).toEqual([]);
    expect(
      parseConfigYaml(`
teams:
  allowed_targets: ["chat:19:abc@thread.v2", "channel:t1/19:ch"]
`).teams.allowed_targets,
    ).toEqual(["chat:19:abc@thread.v2", "channel:t1/19:ch"]);
    expect(() => parseConfigYaml("teams: {allowed_targets: [nonsense]}")).toThrow(ConfigError);
  });

  it("normalises and env-overrides TEAMS_MCP_TEAMS_ALLOWED_TARGETS", () => {
    process.env.TEAMS_MCP_TEAMS_ALLOWED_TARGETS = " CHAT:19:abc , channel:t1/19:ch ";
    resetConfigCache();
    expect(config().teams.allowedTargets).toEqual(["chat:19:abc", "channel:t1/19:ch"]);

    process.env.TEAMS_MCP_TEAMS_ALLOWED_TARGETS = "19:abc,channel:t1/19:ch";
    resetConfigCache();
    expect(() => config()).toThrow(ConfigError);
  });

  it("allows the profile dir and Outlook base to be overridden", () => {
    process.env.TEAMS_MCP_PROFILE_DIR = join(dir, "custom-profile");
    process.env.TEAMS_MCP_OUTLOOK_BASE_URL = "https://outlook.example.com/api/v2.0/";
    resetConfigCache();
    const cfg = config();
    expect(cfg.profileDir).toBe(join(dir, "custom-profile"));
    expect(cfg.outlookBaseUrl).toBe("https://outlook.example.com/api/v2.0");
  });
});
