/**
 * Capability scoping: read-only mode, unsubscribe opt-in, write rate limit and
 * the append-only audit journal.
 *
 * The live-account failure these lock down: an autonomous agent performed bulk
 * mailbox actions — mass moves and irreversible unsubscribes. Every test
 * here uses an injected fetch and a temp config/cache dir; the strongest claim is
 * `calls.length === 0` for each blocked path, i.e. the refusal precedes the
 * request rather than merely preceding the mutation.
 */
import { mkdir, mkdtemp, open, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { ConfigError, config, configFile, parseConfigYaml, resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { setOneClickPoster, setOneClickResolver, setSentLookupDelays } from "../src/outlook/mail.js";
import { auditWrite, isMutatingOp, resetWriteLimiter } from "../src/safety.js";
import { authTool } from "../src/tools/auth.js";
import { calendarTool } from "../src/tools/calendar.js";
import { mailTool } from "../src/tools/mail.js";
import { settingsTool } from "../src/tools/settings.js";
import { teamsTool } from "../src/tools/teams.js";
import { TOOL_MODULES } from "../src/tools/registry.js";
import type { FetchLike } from "../src/http/client.js";

const BASE = "https://outlook.office.com/api/v2.0";

interface Call {
  method: string;
  url: string;
  body: unknown;
}

function harness(responses: Array<() => Response>) {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({
      method: init?.method ?? "GET",
      url,
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    });
    const factory = responses[Math.min(index, responses.length - 1)];
    index += 1;
    if (!factory) throw new Error("no response queued");
    return factory();
  };
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
  return calls;
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status });
const empty = (status: number): (() => Response) => () => new Response(null, { status });

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

function payload(result: CallToolResult): Record<string, unknown> {
  return JSON.parse(textOf(result)) as Record<string, unknown>;
}

async function auditLines(path: string): Promise<Array<Record<string, unknown>>> {
  const text = await readFile(path, "utf8");
  return text
    .split("\n")
    .filter((line) => line.trim() !== "")
    .map((line) => JSON.parse(line) as Record<string, unknown>);
}

/** `target` is free-form JSON in the journal; read one field with a guard. */
function targetField(line: Record<string, unknown>, key: string): unknown {
  const target = line.target;
  if (target === null || typeof target !== "object") return undefined;
  return Reflect.get(target, key);
}

/** Enum members of a tool's `op` schema, narrowed rather than cast inline. */
function opOptions(schema: unknown): string[] {
  if (schema !== null && typeof schema === "object" && "options" in schema && Array.isArray(schema.options)) {
    return schema.options.filter((value): value is string => typeof value === "string");
  }
  throw new Error("tool op schema has no enum options");
}

let dir: string;
let auditPath: string;
const saved = { ...process.env };

/** Change env + forget the memoised config in one step. */
function setEnv(name: string, value: string | undefined): void {
  if (value === undefined) delete process.env[name];
  else process.env[name] = value;
  resetConfigCache();
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-safety-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  delete process.env.TEAMS_MCP_READ_ONLY;
  delete process.env.TEAMS_MCP_UNSUBSCRIBE_ENABLED;
  delete process.env.TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP;
  delete process.env.TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE;
  delete process.env.TEAMS_MCP_USERNAME;
  delete process.env.TEAMS_MCP_PASSWORD;
  delete process.env.TEAMS_MCP_TOTP_SECRET;
  resetConfigCache();
  resetWriteLimiter();
  // Deterministic: no real Sent Items backoff on the write paths.
  setSentLookupDelays([0, 0]);
  auditPath = join(dir, "cache", "write-audit.jsonl");
});

afterEach(async () => {
  setOutlookClient(undefined);
  setOneClickPoster(undefined);
  setOneClickResolver(undefined);
  setSentLookupDelays(undefined);
  process.env = { ...saved };
  resetConfigCache();
  resetWriteLimiter();
  await rm(dir, { recursive: true, force: true });
});

const READ_ONLY_CALLS: Array<{ name: string; run: () => Promise<CallToolResult> }> = [
  {
    name: "mail.send",
    run: () =>
      mailTool.handle({ op: "send", to: ["a@example.com"], subject: "s", body: "b" }),
  },
  { name: "mail.draft", run: () => mailTool.handle({ op: "draft", to: ["a@example.com"], subject: "s", body: "b" }) },
  {
    name: "mail.send_draft",
    run: () => mailTool.handle({ op: "send_draft", draft_id: "d1", confirm: true }),
  },
  { name: "mail.reply", run: () => mailTool.handle({ op: "reply", message_id: "m1", body: "b" }) },
  {
    name: "mail.forward",
    run: () => mailTool.handle({ op: "forward", message_id: "m1", to: ["a@example.com"] }),
  },
  { name: "mail.mark", run: () => mailTool.handle({ op: "mark", message_id: "m1", mark: "read" }) },
  {
    name: "mail.move",
    run: () => mailTool.handle({ op: "move", message_id: "m1", destination: "archive" }),
  },
  {
    name: "mail.unsubscribe_apply",
    run: () => mailTool.handle({ op: "unsubscribe_apply", message_id: "m1", confirm: true }),
  },
  {
    name: "calendar.create",
    run: () => calendarTool.handle({ op: "create", subject: "s", start: "2026-11-18T16:00:00Z" }),
  },
  {
    name: "calendar.update",
    run: () =>
      calendarTool.handle({ op: "update", event_id: "e1", expected_subject: "s", subject: "t" }),
  },
  {
    name: "calendar.delete",
    run: () => calendarTool.handle({ op: "delete", event_id: "e1", expected_subject: "s" }),
  },
  {
    name: "calendar.respond",
    run: () =>
      calendarTool.handle({
        op: "respond",
        event_id: "e1",
        expected_subject: "s",
        response: "accept",
      }),
  },
  { name: "settings.set", run: () => settingsTool.handle({ op: "set", key: "timezone", value: "UTC" }) },
  { name: "auth.logout", run: () => authTool.handle({ op: "logout" }) },
  {
    name: "teams.send",
    run: () =>
      teamsTool.handle({ op: "send", chat_id: "19:chat", text: "hi", confirm: true }),
  },
  {
    name: "teams.create",
    run: () => teamsTool.handle({ op: "create", kind: "chat", topic: "t", members: ["a@b.com"], confirm: true }),
  },
  {
    name: "teams.delete",
    run: () =>
      teamsTool.handle({
        op: "delete",
        team_id: "t1",
        channel_id: "c1",
        expected_name: "x",
        confirm: true,
      }),
  },
];

describe("read-only mode", () => {
  it("is off by default and enabled by env or config, with env taking precedence", async () => {
    expect(config().readOnly).toBe(false);
    for (const value of ["1", "true", "TRUE", "yes", "on", "enabled"]) {
      setEnv("TEAMS_MCP_READ_ONLY", value);
      expect(config().readOnly, value).toBe(true);
    }
    for (const value of ["0", "false", "no", "off", "disabled"]) {
      setEnv("TEAMS_MCP_READ_ONLY", value);
      expect(config().readOnly, value).toBe(false);
    }
    setEnv("TEAMS_MCP_READ_ONLY", undefined);
    const file = parseConfigYaml("safety:\n  read_only: true\n");
    expect(file.safety.read_only).toBe(true);

    await writeFile(configFile(), "safety:\n  read_only: true\n");
    resetConfigCache();
    expect(config().readOnly).toBe(true);
    setEnv("TEAMS_MCP_READ_ONLY", "0");
    expect(config().readOnly).toBe(false);
  });

  it("classifies every advertised op, and every write is in readOnlyGuard's table", () => {
    const expected: Record<string, string[]> = {
      auth: ["logout"],
      mail: ["send", "draft", "send_draft", "reply", "forward", "mark", "move", "unsubscribe_apply"],
      calendar: ["create", "update", "delete", "respond"],
      availability: [],
      people: [],
      settings: ["set"],
      teams: ["send", "create", "delete"],
    };
    const advertised: Record<string, string[]> = {};
    for (const tool of TOOL_MODULES) {
      const options = opOptions(tool.shape.op);
      advertised[tool.name] = options;
      for (const op of options) {
        expect(isMutatingOp(tool.name, op), `${tool.name}.${op}`).toBe(
          (expected[tool.name] ?? []).includes(op),
        );
      }
    }
    // No advertised write is missing from the classification table (a missing
    // entry would silently skip the read-only guard).
    for (const [tool, writes] of Object.entries(expected)) {
      for (const op of writes) expect(advertised[tool], `${tool}.${op}`).toContain(op);
    }
    expect(Object.values(advertised).flat()).toHaveLength(46);
  });

  it("refuses to let the model loosen its own safety policy through settings.set", async () => {
    const attempts: Array<[string, unknown]> = [
      ["safety.read_only", false],
      ["safety.unsubscribe.enabled", true],
      ["safety.unsubscribe.daily_cap", 1000],
      ["safety.write_rate_limit_per_minute", 600],
      ["invite_allowlist", ["attacker@evil.com"]],
      ["teams.allowed_targets", ["chat:19:attacker"]],
    ];
    for (const [key, value] of attempts) {
      const result = await settingsTool.handle({ op: "set", key, value });
      expect(result.isError, key).toBe(true);
      expect(textOf(result), key).toContain("operator-owned");
    }
    const written = await readFile(configFile(), "utf8").catch(() => "");
    expect(written).not.toContain("safety");
    expect(written).not.toContain("teams");
    const lines = await auditLines(auditPath);
    expect(lines).toHaveLength(attempts.length);
    expect(lines.every((line) => line.outcome === "safety_key_protected")).toBe(true);
    expect(lines.map((line) => targetField(line, "key"))).toEqual(attempts.map(([key]) => key));
  });

  it("cannot disable read-only through direct handler calls while it is active", async () => {
    await writeFile(configFile(), "safety:\n  read_only: true\n");
    resetConfigCache();
    const before = await readFile(configFile(), "utf8");
    const calls = harness([]);
    for (const key of [
      "safety.read_only",
      "safety.unsubscribe.enabled",
      "safety.unsubscribe.daily_cap",
      "safety.write_rate_limit_per_minute",
    ]) {
      const result = await settingsTool.handle({ op: "set", key, value: key === "safety.read_only" ? false : 0 });
      expect(result.isError, key).toBe(true);
      expect(textOf(result), key).toContain("read_only");
    }
    expect(await readFile(configFile(), "utf8")).toBe(before);
    expect(calls).toHaveLength(0);
  });

  it("keeps auth.login available in read-only mode but never touches the account", async () => {
    setEnv("TEAMS_MCP_READ_ONLY", "1");
    const calls = harness([]);
    // No credentials on disk: the pre-flight refuses before any browser/network.
    const login = await authTool.handle({ op: "login" });
    expect(login.isError).toBe(true);
    expect(textOf(login)).toContain("credentials_missing");
    expect(calls).toHaveLength(0);
  });

  it("enables read-only through the operator config file, after which writes are blocked", async () => {
    await writeFile(configFile(), "safety:\n  read_only: true\n");
    resetConfigCache();
    const calls = harness([json({ Id: "m1-moved" })]);
    const blocked = await mailTool.handle({ op: "move", message_id: "m1", destination: "archive" });
    expect(textOf(blocked)).toContain("read_only");
    expect(calls).toHaveLength(0);
  });

  it("refuses a typo'd env flag instead of silently leaving writes enabled", () => {
    setEnv("TEAMS_MCP_READ_ONLY", "yse");
    expect(() => config()).toThrow(ConfigError);
    expect(() => config()).toThrow(/TEAMS_MCP_READ_ONLY/);
  });

  it("blocks every mutating op before any request and journals read_only", async () => {
    await writeFile(configFile(), "safety:\n  read_only: true\n");
    resetConfigCache();
    const calls = harness([json({ value: [] })]);

    for (const entry of READ_ONLY_CALLS) {
      const before = calls.length;
      const result = await entry.run();
      expect(result.isError, entry.name).toBe(true);
      expect(textOf(result), entry.name).toContain("read_only");
      expect(calls.length, `${entry.name} issued no request`).toBe(before);
    }

    const lines = await auditLines(auditPath);
    expect(lines).toHaveLength(READ_ONLY_CALLS.length);
    expect(lines.every((line) => line.outcome === "read_only")).toBe(true);
    expect(lines.map((line) => `${line.tool}.${line.op}`).sort()).toEqual(
      READ_ONLY_CALLS.map((entry) => entry.name).sort(),
    );
    expect((await stat(auditPath)).mode & 0o777).toBe(0o600);
  });

  it("keeps the token cache on a blocked auth.logout", async () => {
    const tokenPath = join(dir, "token.json");
    await writeFile(tokenPath, '{"version":1}', { mode: 0o600 });
    setEnv("TEAMS_MCP_READ_ONLY", "1");
    const result = await authTool.handle({ op: "logout" });
    expect(result.isError).toBe(true);
    expect(await readFile(tokenPath, "utf8")).toBe('{"version":1}');
  });

  it("allows reads, auth.status and availability while read-only", async () => {
    setEnv("TEAMS_MCP_READ_ONLY", "1");
    const calls = harness([
      json({
        value: [
          {
            Id: "m1",
            ConversationId: "c1",
            Subject: "Readable",
            ReceivedDateTime: "2026-11-16T12:00:00Z",
            BodyPreview: "p",
          },
        ],
      }),
    ]);
    const list = await mailTool.handle({ op: "list", top: 1 });
    expect(list.isError).toBeUndefined();
    expect(payload(list).count).toBe(1);
    expect(calls).toHaveLength(1);

    expect((await settingsTool.handle({ op: "get" })).isError).toBeUndefined();
    expect((await settingsTool.handle({ op: "path" })).isError).toBeUndefined();
    // auth.status is offline; no token cache here means authenticated:false, not
    // a read_only refusal.
    const status = await authTool.handle({ op: "status" });
    expect(status.isError).toBeUndefined();
    expect(payload(status).authenticated).toBe(false);
  });
});

describe("unsubscribe opt-in", () => {
  const bulkMessage = () =>
    json({
      Id: "bulk-1",
      Subject: "Newsletter",
      From: { EmailAddress: { Address: "news@example.com" } },
      InternetMessageHeaders: [{ Name: "List-Unsubscribe", Value: "<mailto:leave@example.com>" }],
    });

  it("is disabled by default and refuses with zero requests", async () => {
    const calls = harness([bulkMessage()]);
    const result = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
      expected_sender: "news@example.com",
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("unsubscribe_disabled");
    expect(calls).toHaveLength(0);
    const lines = await auditLines(auditPath);
    expect(lines.at(-1)).toMatchObject({ tool: "mail", op: "unsubscribe_apply", outcome: "unsubscribe_disabled" });
  });

  it("requires confirm, expected_sender and a matching From address", async () => {
    setEnv("TEAMS_MCP_UNSUBSCRIBE_ENABLED", "1");
    const calls = harness([bulkMessage(), bulkMessage(), bulkMessage(), bulkMessage()]);

    const noConfirm = await mailTool.handle({ op: "unsubscribe_apply", message_id: "bulk-1" });
    expect(payload(noConfirm).refused).toBe(true);
    expect(String(payload(noConfirm).reason)).toContain("confirm=true");

    const noSender = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
    });
    expect(payload(noSender).refused).toBe(true);
    expect(String(payload(noSender).reason)).toContain("expected_sender");

    const wrongSender = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
      expected_sender: "someone-else@example.com",
    });
    expect(payload(wrongSender).refused).toBe(true);
    expect(String(payload(wrongSender).reason)).toContain("does not match");

    // Three header reads (one per attempt), never a POST/sendmail.
    expect(calls).toHaveLength(3);
    expect(calls.every((call) => call.method === "GET")).toBe(true);
    const outcomes = (await auditLines(auditPath)).map((line) => line.outcome);
    expect(outcomes).toEqual([
      "confirm_required",
      "expected_sender_required",
      "sender_mismatch",
    ]);
  });

  it("applies once when enabled, confirmed and sender-matched", async () => {
    setEnv("TEAMS_MCP_UNSUBSCRIBE_ENABLED", "1");
    const calls = harness([bulkMessage(), empty(202)]);
    const result = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
      expected_sender: "News@Example.com",
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.ok).toBe(true);
    expect(out.results).toEqual([{ target: "leave@example.com", type: "mailto", status: "sent" }]);
    expect(calls.map((call) => `${call.method} ${call.url}`)).toEqual([
      `GET ${BASE}/me/messages/bulk-1?$select=` +
        encodeURIComponent("Id,Subject,From,ReceivedDateTime,InternetMessageHeaders,WebLink"),
      `POST ${BASE}/me/sendmail`,
    ]);
    const statePath = join(dir, "cache", "unsubscribe-state.json");
    expect(JSON.parse(await readFile(statePath, "utf8"))).toMatchObject({ count: 1 });
    expect((await stat(statePath)).mode & 0o777).toBe(0o600);
    expect((await auditLines(auditPath)).at(-1)).toMatchObject({ outcome: "ok" });
  });

  it("enforces the daily cap before any request, with a reset time", async () => {
    setEnv("TEAMS_MCP_UNSUBSCRIBE_ENABLED", "1");
    setEnv("TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP", "1");
    const calls = harness([bulkMessage(), empty(202), bulkMessage()]);
    const args = {
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
      expected_sender: "news@example.com",
    };
    expect((await mailTool.handle(args)).isError).toBeUndefined();
    const before = calls.length;

    const capped = await mailTool.handle(args);
    expect(capped.isError).toBe(true);
    expect(textOf(capped)).toContain("cap_reached");
    expect(calls.length).toBe(before);
    const state = JSON.parse(await readFile(join(dir, "cache", "unsubscribe-state.json"), "utf8"));
    expect(state.count).toBe(1);
    // The hint carries the reset instant; it must be in the future.
    const resetText = textOf(capped).match(/reset(?:s)? at ([^)]+)/)?.[1];
    expect(resetText).toBeDefined();
    expect(Date.parse(resetText as string)).toBeGreaterThan(Date.now());
    expect((await auditLines(auditPath)).at(-1)).toMatchObject({ outcome: "cap_reached" });
  });

  it("treats a daily cap of 0 as fully disabled after the opt-in", async () => {
    setEnv("TEAMS_MCP_UNSUBSCRIBE_ENABLED", "1");
    setEnv("TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP", "0");
    const calls = harness([bulkMessage()]);
    const result = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
      expected_sender: "news@example.com",
    });
    expect(textOf(result)).toContain("cap_reached");
    expect(calls).toHaveLength(0);
  });

  const twoTargetMessage = () =>
    json({
      Id: "bulk-2",
      Subject: "Newsletter",
      From: { EmailAddress: { Address: "news@example.com" } },
      InternetMessageHeaders: [
        {
          Name: "List-Unsubscribe",
          Value: "<https://lists.unsubscribe-host.com/u/1>, <mailto:leave@example.com>",
        },
      ],
    });

  it("charges one cap slot per external action, not per apply", async () => {
    setEnv("TEAMS_MCP_UNSUBSCRIBE_ENABLED", "1");
    setEnv("TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP", "3");
    const posts: string[] = [];
    setOneClickPoster(async (target) => {
      posts.push(target);
      return { status: 200 };
    });
    setOneClickResolver(async () => ["93.184.216.34"]);
    const calls = harness([twoTargetMessage(), empty(202), twoTargetMessage()]);
    const args = {
      op: "unsubscribe_apply",
      message_id: "bulk-2",
      confirm: true,
      expected_sender: "news@example.com",
    };

    const first = await mailTool.handle(args);
    expect(first.isError, textOf(first)).toBeUndefined();
    const results = payload(first).results;
    expect(Array.isArray(results) ? results.length : -1).toBe(2);
    expect(posts).toHaveLength(1);
    expect(calls.map((call) => call.method)).toEqual(["GET", "POST"]);
    const statePath = join(dir, "cache", "unsubscribe-state.json");
    expect(JSON.parse(await readFile(statePath, "utf8"))).toMatchObject({ count: 2 });

    // 2 used + 2 requested > 3: refused before any second external action.
    const second = await mailTool.handle(args);
    expect(second.isError).toBe(true);
    expect(textOf(second)).toContain("cap_reached");
    expect(posts).toHaveLength(1);
    expect(calls.map((call) => call.method)).toEqual(["GET", "POST", "GET"]);
    expect(JSON.parse(await readFile(statePath, "utf8"))).toMatchObject({ count: 2 });
  });

  it("refuses a two-target apply that cannot fit the remaining cap", async () => {
    setEnv("TEAMS_MCP_UNSUBSCRIBE_ENABLED", "1");
    setEnv("TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP", "1");
    const posts: string[] = [];
    setOneClickPoster(async (target) => {
      posts.push(target);
      return { status: 200 };
    });
    setOneClickResolver(async () => ["93.184.216.34"]);
    const calls = harness([twoTargetMessage()]);
    const result = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-2",
      confirm: true,
      expected_sender: "news@example.com",
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("cap_reached");
    // Only the header read: no POST and no mailto send.
    expect(calls.map((call) => call.method)).toEqual(["GET"]);
    expect(posts).toHaveLength(0);
    const statePath = join(dir, "cache", "unsubscribe-state.json");
    expect(await readFile(statePath, "utf8").catch(() => null)).toBeNull();
  });
});

describe("write rate limit", () => {
  it("limits mail.move before the request and reports retry_after", async () => {
    setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "2");
    resetWriteLimiter();
    const calls = harness([json({ Id: "m1-moved" })]);
    const args = { op: "move", message_id: "m1", destination: "archive" };

    expect((await mailTool.handle(args)).isError).toBeUndefined();
    expect((await mailTool.handle(args)).isError).toBeUndefined();
    expect(calls).toHaveLength(2);

    const limited = await mailTool.handle(args);
    expect(limited.isError).toBe(true);
    expect(textOf(limited)).toContain("rate_limited");
    expect(textOf(limited)).toMatch(/retry in [1-9]\d*s/);
    expect(calls).toHaveLength(2);
    expect((await auditLines(auditPath)).map((line) => line.outcome)).toEqual([
      "ok",
      "ok",
      "rate_limited",
    ]);
  });

  it("shares one bucket across mark and move but not reads", async () => {
    setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "2");
    resetWriteLimiter();
    const calls = harness([
      json({ Id: "m1-moved" }),
      json({ Id: "m2" }),
      json({ value: [] }),
    ]);
    expect((await mailTool.handle({ op: "move", message_id: "m1", destination: "archive" })).isError).toBeUndefined();
    expect((await mailTool.handle({ op: "mark", message_id: "m2", mark: "read" })).isError).toBeUndefined();
    const limited = await mailTool.handle({ op: "mark", message_id: "m3", mark: "read" });
    expect(textOf(limited)).toContain("rate_limited");
    expect(calls).toHaveLength(2);

    // A read is never rate-limited and still reaches the provider.
    const list = await mailTool.handle({ op: "list", top: 1 });
    expect(list.isError).toBeUndefined();
    expect(calls).toHaveLength(3);
  });

  it("does not consume a token for an invalid call", async () => {
    setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "1");
    resetWriteLimiter();
    const calls = harness([json({ Id: "m1-moved" })]);
    for (let attempt = 0; attempt < 3; attempt += 1) {
      expect((await mailTool.handle({ op: "move", message_id: "m1" })).isError).toBe(true);
    }
    expect(calls).toHaveLength(0);
    expect((await mailTool.handle({ op: "move", message_id: "m1", destination: "archive" })).isError).toBeUndefined();
    expect(calls).toHaveLength(1);
  });

  it("leaves a normal single write unaffected under the default limit", async () => {
    resetWriteLimiter();
    const calls = harness([json({ Id: "m1-moved" })]);
    const result = await mailTool.handle({ op: "move", message_id: "m1", destination: "archive" });
    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
  });

  const eventBody = () => ({
    Id: "e1",
    Subject: "Real subject",
    Start: { DateTime: "2026-11-18T16:00:00.0000000", TimeZone: "UTC" },
    End: { DateTime: "2026-11-18T16:30:00.0000000", TimeZone: "UTC" },
    IsAllDay: false,
    ShowAs: "Busy",
    Attendees: [],
  });
  const messageBody = () => ({
    Id: "m1",
    Subject: "Hi",
    From: { EmailAddress: { Address: "a@example.com" } },
  });

  it("limits every outbound/account mutation before its first request", async () => {
    const cases: Array<{ name: string; responses: Array<() => Response>; run: () => Promise<CallToolResult> }> = [
      {
        name: "mail.send",
        responses: [empty(202)],
        run: () => mailTool.handle({ op: "send", to: ["a@example.com"], subject: "s", body: "b" }),
      },
      {
        name: "mail.reply",
        responses: [json(messageBody()), empty(202)],
        run: () => mailTool.handle({ op: "reply", message_id: "m1", body: "b" }),
      },
      {
        name: "mail.forward",
        responses: [json(messageBody()), empty(202)],
        run: () => mailTool.handle({ op: "forward", message_id: "m1", to: ["a@example.com"] }),
      },
      {
        name: "mail.mark",
        responses: [json({ Id: "m1" })],
        run: () => mailTool.handle({ op: "mark", message_id: "m1", mark: "read" }),
      },
      {
        name: "mail.move",
        responses: [json({ Id: "m1-moved" })],
        run: () => mailTool.handle({ op: "move", message_id: "m1", destination: "archive" }),
      },
      {
        name: "calendar.create",
        responses: [json(eventBody(), 201)],
        run: () => calendarTool.handle({ op: "create", subject: "s" }),
      },
      {
        name: "calendar.update",
        responses: [json(eventBody()), json(eventBody())],
        run: () =>
          calendarTool.handle({ op: "update", event_id: "e1", expected_subject: "Real subject", subject: "t" }),
      },
      {
        name: "calendar.delete",
        responses: [json(eventBody()), empty(204)],
        run: () => calendarTool.handle({ op: "delete", event_id: "e1", expected_subject: "Real subject" }),
      },
      {
        name: "calendar.respond",
        responses: [json(eventBody()), empty(202)],
        run: () =>
          calendarTool.handle({
            op: "respond",
            event_id: "e1",
            expected_subject: "Real subject",
            response: "accept",
          }),
      },
      {
        name: "settings.set",
        responses: [],
        run: () => settingsTool.handle({ op: "set", key: "timezone", value: "UTC" }),
      },
      { name: "auth.logout", responses: [], run: () => authTool.handle({ op: "logout" }) },
    ];

    for (const entry of cases) {
      setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "1");
      resetWriteLimiter();
      const calls = harness(entry.responses);
      const first = await entry.run();
      expect(first.isError, `${entry.name} first call`).toBeUndefined();
      const afterFirst = calls.length;
      const second = await entry.run();
      expect(second.isError, `${entry.name} second call`).toBe(true);
      expect(textOf(second), entry.name).toContain("rate_limited");
      expect(textOf(second), entry.name).toMatch(/retry in [1-9]\d*s/);
      expect(calls.length, `${entry.name} issued no request while limited`).toBe(afterFirst);
      expect((await auditLines(auditPath)).at(-1), entry.name).toMatchObject({
        tool: entry.name.split(".")[0],
        op: entry.name.split(".")[1],
        outcome: "rate_limited",
      });
    }
  });

  it("does not consume a token for invalid arguments on any write op", async () => {
    setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "1");
    resetWriteLimiter();
    const calls = harness([empty(202)]);
    const invalid: Array<Promise<CallToolResult>> = [
      mailTool.handle({ op: "send", to: [], subject: "s", body: "b" }),
      mailTool.handle({ op: "draft", to: [], subject: "s", body: "b" }),
      mailTool.handle({ op: "send_draft", confirm: true }),
      mailTool.handle({ op: "reply", message_id: "m1" }),
      mailTool.handle({ op: "forward", message_id: "m1", to: [] }),
      mailTool.handle({ op: "mark", message_id: "m1" }),
      mailTool.handle({ op: "move", message_id: "m1" }),
      calendarTool.handle({ op: "create" }),
      calendarTool.handle({ op: "update", event_id: "e1", expected_subject: "s" }),
      calendarTool.handle({ op: "delete", event_id: "e1" }),
      calendarTool.handle({ op: "respond", event_id: "e1", expected_subject: "s", response: "maybe" }),
      settingsTool.handle({ op: "set", key: "timezone" }),
      settingsTool.handle({ op: "set", key: "no.such.key", value: 1 }),
    ];
    for (const result of await Promise.all(invalid)) expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);

    const sent = await mailTool.handle({ op: "send", to: ["a@example.com"], subject: "S", body: "B" });
    expect(sent.isError).toBeUndefined();
    // One 202 plus the Sent Items lookups for the id (nothing is sent under the
    // temp config; the id lookup is best-effort and harmless).
    expect(calls[0]?.url).toBe(`${BASE}/me/sendmail`);
    expect(calls.length).toBe(4);
  });

  it("does not spend a token when auth.login cannot even start", async () => {
    setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "1");
    resetWriteLimiter();
    const calls = harness([empty(202)]);
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const login = await authTool.handle({ op: "login" });
      expect(login.isError).toBe(true);
      expect(textOf(login)).toContain("credentials_missing");
    }
    const sent = await mailTool.handle({ op: "send", to: ["a@example.com"], subject: "s", body: "b" });
    expect(sent.isError).toBeUndefined();
    expect(calls[0]?.url).toBe(`${BASE}/me/sendmail`);
    expect(calls.length).toBe(4);
  });

  it("refunds the token when a write is refused before it mutates", async () => {
    setEnv("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", "1");
    resetWriteLimiter();
    const calls = harness([empty(503), json(messageBody()), empty(202)]);
    const failed = await mailTool.handle({ op: "reply", message_id: "m1", body: "b" });
    expect(failed.isError).toBe(true);
    expect(calls).toHaveLength(1);

    const retried = await mailTool.handle({ op: "reply", message_id: "m1", body: "b" });
    expect(retried.isError, textOf(retried)).toBeUndefined();
    // GET original + POST reply + the zero-delay Sent Items lookups.
    expect(calls.length).toBe(6);
  });
});

describe("write audit journal", () => {
  it("records one line per attempt without bodies or tokens", async () => {
    const calls = harness([json({ Id: "m1-moved" })]);
    await mailTool.handle({ op: "move", message_id: "m1", destination: "archive" });
    expect(calls).toHaveLength(1);

    const lines = await auditLines(auditPath);
    expect(lines).toHaveLength(1);
    const line = lines[0] as Record<string, unknown>;
    expect(line).toMatchObject({
      tool: "mail",
      op: "move",
      target: { message_id: "m1", destination: "archive" },
      outcome: "ok",
    });
    expect(typeof line.ts).toBe("string");
    expect(Number.isNaN(Date.parse(String(line.ts)))).toBe(false);
    expect(JSON.stringify(line)).not.toMatch(/Bearer|body|token/i);
  });

  it("rotates at 5 MB and keeps the new file owner-only", async () => {
    const path = auditPath;
    await mkdir(join(dir, "cache"), { recursive: true });
    await writeFile(path, "", { mode: 0o600 });
    const handle = await open(path, "r+");
    await handle.truncate(5 * 1024 * 1024);
    await handle.close();

    setEnv("TEAMS_MCP_READ_ONLY", "1");
    const result = await authTool.handle({ op: "logout" });
    expect(result.isError).toBe(true);

    expect((await stat(`${path}.1`)).size).toBeGreaterThanOrEqual(5 * 1024 * 1024);
    const lines = await auditLines(path);
    expect(lines).toHaveLength(1);
    expect(lines[0]).toMatchObject({ tool: "auth", op: "logout", outcome: "read_only" });
    expect((await stat(path)).mode & 0o777).toBe(0o600);
    expect((await stat(`${path}.1`)).mode & 0o777).toBe(0o600);
  });

  it("never records subjects, bodies or tokens", async () => {
    const calls = harness([empty(202)]);
    await mailTool.handle({
      op: "send",
      to: ["a@example.com"],
      subject: "SECRET SUBJECT",
      body: "SECRET BODY Bearer SECRET-TOKEN",
    });
    expect(calls[0]?.url).toBe(`${BASE}/me/sendmail`);
    const raw = await readFile(auditPath, "utf8");
    expect(raw).not.toContain("SECRET");
    const line = (await auditLines(auditPath))[0];
    expect(line?.target).toEqual({ recipients: 1 });
    expect(line).toMatchObject({ tool: "mail", op: "send", outcome: "ok" });
  });

  it("records guard and allowlist refusals without touching the account", async () => {
    await writeFile(configFile(), 'invite_allowlist: ["@allowed.com"]\n');
    resetConfigCache();
    const calls = harness([
      json({
        Id: "e1",
        Subject: "Real subject",
        Start: { DateTime: "2026-11-18T16:00:00.0000000", TimeZone: "UTC" },
        End: { DateTime: "2026-11-18T16:30:00.0000000", TimeZone: "UTC" },
      }),
    ]);
    const refusedDelete = await calendarTool.handle({
      op: "delete",
      event_id: "e1",
      expected_subject: "Nope",
    });
    expect(refusedDelete.isError).toBe(true);
    expect(calls.map((call) => call.method)).toEqual(["GET"]);

    const refusedSend = await mailTool.handle({
      op: "send",
      to: ["attacker@evil.com"],
      subject: "s",
      body: "b",
    });
    expect(refusedSend.isError).toBe(true);
    expect(textOf(refusedSend)).toContain("recipient_not_allowed");
    expect(calls).toHaveLength(1);

    const outcomes = (await auditLines(auditPath)).map((line) => line.outcome);
    expect(outcomes).toEqual(["subject_mismatch", "recipient_not_allowed"]);
  });

  it("keeps writes working when the journal cannot be written", async () => {
    // A regular file where the cache directory should be makes mkdir/append fail.
    await writeFile(join(dir, "blocker"), "not a directory", { mode: 0o600 });
    setEnv("TEAMS_MCP_CACHE_DIR", join(dir, "blocker", "cache"));
    expect(() =>
      auditWrite({ tool: "mail", op: "move", target: { message_id: "m1" }, outcome: "ok" }),
    ).not.toThrow();

    const calls = harness([json({ Id: "m1-moved" })]);
    const result = await mailTool.handle({ op: "move", message_id: "m1", destination: "archive" });
    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
  });
});
