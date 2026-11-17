import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { resetConfigCache } from "../src/config.js";
import {
  effectiveSettings,
  getSetting,
  reloadSettings,
  SettingsError,
  setSetting,
  validateSettings,
} from "../src/settings.js";

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-settings-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  delete process.env.TEAMS_MCP_CONFIG_FILE;
  delete process.env.TEAMS_MCP_TIMEZONE;
  delete process.env.TEAMS_MCP_DEFAULT_DURATION_MIN;
  resetConfigCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("settings", () => {
  it("sets a scalar, persists it 0600 and returns the effective value", async () => {
    const result = await setSetting("default_duration_min", "45");
    expect(result.value).toBe(45);

    const path = join(dir, "config.yaml");
    const text = await readFile(path, "utf8");
    expect(text).toContain("default_duration_min: 45");
    expect((await stat(path)).mode & 0o777).toBe(0o600);

    await expect(getSetting("default_duration_min")).resolves.toMatchObject({ value: 45 });
  });

  it("writes nested keys and alias entries", async () => {
    await setSetting("work_hours.start", "08:30");
    await setSetting("aliases.joel", "joel@example.com");
    const settings = effectiveSettings();
    expect(settings.work_hours.start).toBe("08:30");
    expect(settings.aliases.joel).toBe("joel@example.com");

    await expect(getSetting("aliases.joel")).resolves.toMatchObject({ value: "joel@example.com" });
  });

  it("refuses an invalid value without touching the file", async () => {
    await setSetting("default_duration_min", 30);
    await expect(setSetting("work_hours.start", "9am")).rejects.toBeInstanceOf(SettingsError);

    const text = await readFile(join(dir, "config.yaml"), "utf8");
    expect(text).toContain("default_duration_min: 30");
    expect(text).not.toContain("9am");
  });

  it("refuses unknown keys", async () => {
    await expect(setSetting("nope.value", 1)).rejects.toBeInstanceOf(SettingsError);
    await expect(setSetting("__proto__.polluted", 1)).rejects.toBeInstanceOf(SettingsError);
  });

  it("coerces CLI strings", async () => {
    await setSetting("max_results.slots", "8");
    await setSetting("work_hours.days", "mon,tue,wed");
    const settings = effectiveSettings();
    expect(settings.max_results.slots).toBe(8);
    expect(settings.work_hours.days).toEqual(["mon", "tue", "wed"]);
  });

  it("reloads from disk and reports validation results", async () => {
    await writeFile(join(dir, "config.yaml"), "default_duration_min: 25\n");
    resetConfigCache();
    const reloaded = await reloadSettings();
    expect(reloaded.reloaded).toBe(true);
    expect(reloaded.settings?.default_duration_min).toBe(25);

    await writeFile(join(dir, "config.yaml"), "work_hours: {start: 9am}\n");
    resetConfigCache();
    const invalid = await validateSettings();
    expect(invalid.valid).toBe(false);
    expect(invalid.issues?.join(" ")).toContain("work_hours.start");
  });

  it("reports an env override for the key", async () => {
    process.env.TEAMS_MCP_TIMEZONE = "Europe/Berlin";
    resetConfigCache();
    await expect(getSetting("timezone")).resolves.toMatchObject({
      value: "Europe/Berlin",
      env_override: "TEAMS_MCP_TIMEZONE",
    });
  });

  it("returns the whole effective config for get without a key", async () => {
    const result = await getSetting();
    expect(result.settings?.timezone).toBe("America/Chicago");
    expect(result.settings?.max_results).toBeDefined();
  });
});
