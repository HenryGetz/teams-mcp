/**
 * Deterministic adversarial fuzzing for the six op-dispatched tools.
 *
 * Every call goes through `handle()` directly (the MCP schema is fuzzed
 * separately in `protocol-fuzz.test.ts`) with an injected Outlook/Graph transport:
 * no test can reach the live account, and every mutating request is classified
 * against the op that was asked for. Seeds are fixed, so a failure reproduces.
 */
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { z } from "zod";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { config, ConfigError, isValidTimeZone, resetConfigCache } from "../src/config.js";
import { GraphClient, setGraphClient } from "../src/graph/client.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import {
  setOneClickPoster,
  setOneClickResolver,
  setSentLookupDelays,
} from "../src/outlook/mail.js";
import {
  clearSnapshotCache,
  readSnapshot,
  snapshotPath,
} from "../src/snapshot/calendar.js";
import { clearPeopleCache } from "../src/tools/people.js";
import { setSetting } from "../src/settings.js";
import { resetWriteLimiter } from "../src/safety.js";
import { authTool } from "../src/tools/auth.js";
import { availabilityTool } from "../src/tools/availability.js";
import { calendarTool } from "../src/tools/calendar.js";
import { mailTool } from "../src/tools/mail.js";
import { peopleTool } from "../src/tools/people.js";
import { settingsTool } from "../src/tools/settings.js";
import { encodeCursor, type ToolModule } from "../src/tools/shared.js";

/* ------------------------------------------------------------------ */
/* Deterministic randomness                                            */
/* ------------------------------------------------------------------ */

/** mulberry32: tiny, seedable, stable across Node versions. */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffled<T>(values: readonly T[], next: () => number): T[] {
  const copy = [...values];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swap = Math.floor(next() * (index + 1));
    const current = copy[index] as T;
    copy[index] = copy[swap] as T;
    copy[swap] = current;
  }
  return copy;
}

/* ------------------------------------------------------------------ */
/* Mock Outlook REST + Graph transport                                 */
/* ------------------------------------------------------------------ */

interface LoggedCall {
  api: "outlook" | "graph";
  method: string;
  url: string;
  body: unknown;
  headers: unknown;
}

type Fetch = (url: string, init?: RequestInit) => Promise<Response>;
type Responder = (url: string, init: RequestInit | undefined, body: unknown) => Promise<Response>;

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
const empty = (status: number): Response => new Response(null, { status });
const notFound = (): Response => json({ error: { code: "NotFound", message: "not found" } }, 404);

function messageRow(id = "msg-1"): Record<string, unknown> {
  return {
    Id: id,
    ConversationId: "conv-1",
    Subject: "Fuzz message",
    From: { EmailAddress: { Name: "Sender", Address: "sender@example.com" } },
    ToRecipients: [{ EmailAddress: { Address: "me@example.com" } }],
    CcRecipients: [],
    ReceivedDateTime: "2026-11-16T12:00:00Z",
    IsRead: true,
    BodyPreview: "preview text",
    HasAttachments: false,
    Body: { ContentType: "HTML", Content: "<p>body text</p>" },
    InternetMessageHeaders: [
      { Name: "List-Unsubscribe", Value: "<https://lists.unsubscribe-host.com/u/1>, <mailto:leave@allowed.com>" },
      { Name: "List-Unsubscribe-Post", Value: "List-Unsubscribe=One-Click" },
    ],
    WebLink: "https://outlook.office365.com/owa/?ItemID=msg-1",
  };
}

function sentRow(subject: string): Record<string, unknown> {
  return {
    Id: "sent-1",
    ConversationId: "conv-1",
    Subject: subject,
    ReceivedDateTime: new Date().toISOString(),
  };
}

function eventRow(id = "evt-1", subject = "Fuzz Event"): Record<string, unknown> {
  return {
    Id: id,
    Subject: subject,
    Start: { DateTime: "2026-11-18T16:00:00.0000000", TimeZone: "UTC" },
    End: { DateTime: "2026-11-18T17:00:00.0000000", TimeZone: "UTC" },
    IsAllDay: false,
    ShowAs: "Busy",
    Location: { DisplayName: "Room 1" },
    Organizer: { EmailAddress: { Name: "Me", Address: "me@example.com" } },
    Attendees: [
      {
        EmailAddress: { Address: "a@example.com", Name: "A" },
        Type: "Required",
        Status: { Response: "Accepted" },
      },
    ],
    IsOnlineMeeting: false,
    WebLink: "https://outlook.office365.com/owa/?itemid=evt-1",
    BodyPreview: "event preview",
  };
}

function directoryUser(id = "user-1"): Record<string, unknown> {
  return {
    id,
    displayName: "Fuzz User",
    mail: "fuzz@example.com",
    userPrincipalName: "fuzz@example.com",
    jobTitle: "Tester",
    department: "QA",
  };
}

function scheduleRows(body: unknown): Record<string, unknown> {
  const schedules =
    body !== null && typeof body === "object" && "Schedules" in body && Array.isArray(body.Schedules)
      ? body.Schedules.filter((value): value is string => typeof value === "string")
      : [];
  return {
    value: schedules.map((email) => ({
      ScheduleId: email,
      AvailabilityView: "0",
      WorkingHours: { StartTime: "09:00", EndTime: "17:00" },
      ScheduleItems: [],
    })),
  };
}

function outlookResponse(method: string, url: string, body: unknown): Response {
  const path = new URL(url).pathname;
  if (method === "POST" && path.endsWith("/me/sendmail")) return empty(202);
  if (method === "POST" && /\/me\/messages\/[^/]+\/(reply|forward)$/.test(path)) return empty(202);
  if (method === "POST" && /\/me\/messages\/[^/]+\/move$/.test(path)) return json({ Id: "moved-1" });
  if (method === "PATCH" && /\/me\/messages\/[^/]+$/.test(path)) return json(messageRow("patched-1"));
  if (method === "GET" && /\/me\/mailfolders\/[^/]+\/messages$/.test(path)) {
    if (path.includes("sentitems")) return json({ value: [sentRow("RE: Fuzz message")] });
    return json({ value: [messageRow()] });
  }
  if (method === "GET" && path.endsWith("/me/messages")) return json({ value: [messageRow()] });
  if (method === "GET" && /\/me\/messages\/[^/]+$/.test(path)) return json(messageRow());
  if (method === "GET" && path.endsWith("/me/calendarview")) return json({ value: [eventRow()] });
  if (method === "POST" && path.endsWith("/me/events")) return json(eventRow("created-1"), 201);
  if (method === "GET" && /\/me\/events\/[^/]+\/instances$/.test(path)) return json({ value: [eventRow()] });
  if (method === "POST" && /\/me\/events\/[^/]+\/(accept|decline|tentativelyaccept)$/.test(path)) {
    return empty(202);
  }
  if (method === "PATCH" && /\/me\/events\/[^/]+$/.test(path)) return json(eventRow());
  if (method === "DELETE" && /\/me\/events\/[^/]+$/.test(path)) return empty(204);
  if (method === "GET" && /\/me\/events\/[^/]+$/.test(path)) return json(eventRow());
  if (method === "POST" && path.endsWith("/me/calendar/getschedule")) return json(scheduleRows(body));
  if (method === "POST" && path.endsWith("/me/findmeetingtimes")) {
    return json({ MeetingTimeSuggestions: [] });
  }
  return notFound();
}

function graphResponse(method: string, url: string): Response {
  const parsed = new URL(url);
  if (method !== "GET") return notFound();
  if (/\/users\/[^/]+$/.test(parsed.pathname)) {
    if (decodeURIComponent(parsed.pathname).includes("missing")) return notFound();
    return json(directoryUser());
  }
  if (parsed.pathname.endsWith("/users")) {
    const search = parsed.searchParams.get("$search") ?? "";
    if (search.includes("ambiguous")) {
      return json({ value: [directoryUser("user-1"), directoryUser("user-2")] });
    }
    if (search.includes("nobody")) return json({ value: [] });
    return json({ value: [directoryUser()] });
  }
  return notFound();
}

interface Harness {
  calls: LoggedCall[];
  posters: string[];
}

function installMock(overrides?: { outlook?: Responder; graph?: Responder }): Harness {
  const calls: LoggedCall[] = [];
  const posters: string[] = [];
  const capture =
    (api: "outlook" | "graph"): Fetch =>
    async (url, init) => {
      const body: unknown = typeof init?.body === "string" ? JSON.parse(init.body) : undefined;
      calls.push({
        api,
        method: init?.method ?? "GET",
        url,
        body,
        headers: init?.headers,
      });
      if (api === "outlook") {
        return overrides?.outlook
          ? overrides.outlook(url, init, body)
          : Promise.resolve(outlookResponse(init?.method ?? "GET", url, body));
      }
      return overrides?.graph
        ? overrides.graph(url, init, body)
        : Promise.resolve(graphResponse(init?.method ?? "GET", url));
    };
  setOutlookClient(
    new OutlookClient({
      tokenProvider: async () => ({ token: "mock-outlook-token" }),
      fetchImpl: capture("outlook"),
      maxAttempts: 1,
    }),
  );
  setGraphClient(
    new GraphClient({
      tokenProvider: async () => ({ token: "mock-graph-token" }),
      fetchImpl: capture("graph"),
      maxAttempts: 1,
    }),
  );
  setOneClickPoster(async (target) => {
    posters.push(target);
    return { status: 200 };
  });
  setOneClickResolver(async () => ["93.184.216.34"]);
  setSentLookupDelays([0]);
  return { calls, posters };
}

/* ------------------------------------------------------------------ */
/* Fixtures and helpers                                                */
/* ------------------------------------------------------------------ */

const TOOLS: readonly ToolModule[] = [
  authTool,
  mailTool,
  calendarTool,
  availabilityTool,
  peopleTool,
  settingsTool,
];

const ALL_OPS: ReadonlyArray<{ tool: ToolModule; op: string }> = [
  ...["status"].map((op) => ({ tool: authTool, op })),
  ...["list", "search", "get", "thread", "send", "reply", "forward", "mark", "move", "unsubscribe_preview", "unsubscribe_apply"].map(
    (op) => ({ tool: mailTool, op }),
  ),
  ...["list", "get", "create", "update", "delete", "respond", "instances"].map((op) => ({
    tool: calendarTool,
    op,
  })),
  ...["freebusy", "suggest", "snapshot"].map((op) => ({ tool: availabilityTool, op })),
  ...["search", "resolve"].map((op) => ({ tool: peopleTool, op })),
  ...["get", "validate", "reload", "path"].map((op) => ({ tool: settingsTool, op })),
];

/** Args that make the op reach its happy path; random keys then corrupt them. */
function validArgs(tool: string, op: string): Record<string, unknown> {
  const start = "2026-11-17";
  const end = "2026-11-18";
  const map: Record<string, Record<string, unknown>> = {
    "auth.status": {},
    "mail.list": { top: 5 },
    "mail.search": { query: "fuzz" },
    "mail.get": { message_id: "msg-1" },
    "mail.thread": { thread_id: "conv-1", top: 5 },
    "mail.send": { to: ["a@example.com"], subject: "Fuzz send", body: "body" },
    "mail.reply": { message_id: "msg-1", body: "reply body" },
    "mail.forward": { message_id: "msg-1", to: ["a@example.com"] },
    "mail.mark": { message_id: "msg-1", mark: "read" },
    "mail.move": { message_id: "msg-1", destination: "archive" },
    "mail.unsubscribe_preview": { message_id: "msg-1" },
    "mail.unsubscribe_apply": { message_id: "msg-1" },
    "calendar.list": { top: 5 },
    "calendar.get": { event_id: "evt-1" },
    "calendar.create": {
      subject: "Fuzz Event",
      start: "2026-11-18T16:00:00Z",
      end: "2026-11-18T17:00:00Z",
    },
    "calendar.update": { event_id: "evt-1", expected_subject: "Fuzz Event", subject: "Renamed" },
    "calendar.delete": { event_id: "evt-1", expected_subject: "Fuzz Event" },
    "calendar.respond": { event_id: "evt-1", expected_subject: "Fuzz Event", response: "accept" },
    "calendar.instances": { event_id: "evt-1", start, end },
    "availability.freebusy": { attendees: ["a@example.com"], start, end },
    "availability.suggest": { attendees: ["a@example.com"], start, end },
    "availability.snapshot": { attendees: ["a@example.com"], start, end },
    "people.search": { query: "fuzz", top: 5 },
    "people.resolve": { emails: ["a@example.com"] },
    "settings.get": {},
    "settings.validate": {},
    "settings.reload": {},
    "settings.path": {},
  };
  return { ...(map[`${tool}.${op}`] ?? {}) };
}

const VALID_CURSOR = encodeCursor("https://outlook.office.com/api/v2.0/me/messages?$skip=1");
const FOREIGN_CURSOR = encodeCursor("https://graph.microsoft.com/v1.0/users?$top=1");
const GARBAGE_CURSOR = encodeCursor("not-a-url");

/** Values that have broken a validator in some system, or a sibling. */
const HOSTILE: ReadonlyArray<() => unknown> = [
  () => undefined,
  () => null,
  () => 0,
  () => -1,
  () => 1.5,
  () => Number.NaN,
  () => Number.POSITIVE_INFINITY,
  () => "",
  () => "   ",
  () => "x".repeat(10_000),
  () => "\u0000\r\n💥\u007f",
  () => "../../etc/passwd",
  () => "attacker@evil.com, victim@allowed.com",
  () => "a@b@c",
  () => "Name <a@example.com>",
  () => "2026-13-01",
  () => "2026-02-30",
  () => "Mars/Base",
  () => "0001-01-01",
  () => "9999-12-31",
  () => VALID_CURSOR,
  () => FOREIGN_CURSOR,
  () => GARBAGE_CURSOR,
  () => [],
  () => [1, 2, 3],
  () => [null],
  () => [""],
  () => ["a@example.com"],
  () => ["a@b@c"],
  () => ["a\r\nBcc: b@example.com"],
  () => Array.from({ length: 1000 }, (_, index) => `u${index}@example.com`),
  () => ({}),
  () => ({ nested: { deeper: { deepest: true } } }),
  () => true,
  () => false,
  () => "__proto__",
];

function describeArgs(args: Record<string, unknown>): string {
  return Object.entries(args)
    .map(([key, value]) =>
      Array.isArray(value)
        ? `${key}=array(${value.length})`
        : `${key}=${typeof value}${typeof value === "string" ? `(${value.length})` : ""}`,
    )
    .join(",");
}

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

function payloadOf(result: CallToolResult): Record<string, unknown> {
  return z.record(z.string(), z.unknown()).parse(JSON.parse(textOf(result)));
}

function mutationOf(call: LoggedCall): string | undefined {
  if (call.method === "GET") return undefined;
  const path = new URL(call.url).pathname;
  // Scheduling reads are POSTs (`getschedule`, `findmeetingtimes`): no state changes.
  if (path.endsWith("/me/calendar/getschedule") || path.endsWith("/me/findmeetingtimes")) {
    return undefined;
  }
  if (call.api === "outlook") {
    if (call.method === "POST" && path.endsWith("/me/sendmail")) return "mail.send";
    if (call.method === "POST" && /\/me\/messages\/[^/]+\/reply$/.test(path)) return "mail.reply";
    if (call.method === "POST" && /\/me\/messages\/[^/]+\/forward$/.test(path)) return "mail.forward";
    if (call.method === "POST" && /\/me\/messages\/[^/]+\/move$/.test(path)) return "mail.move";
    if (call.method === "PATCH" && /\/me\/messages\/[^/]+$/.test(path)) return "mail.mark";
    if (call.method === "POST" && path.endsWith("/me/events")) return "calendar.create";
    if (call.method === "PATCH" && /\/me\/events\/[^/]+$/.test(path)) return "calendar.update";
    if (call.method === "DELETE" && /\/me\/events\/[^/]+$/.test(path)) return "calendar.delete";
    if (call.method === "POST" && /\/me\/events\/[^/]+\/(accept|decline|tentativelyaccept)$/.test(path)) {
      return "calendar.respond";
    }
  }
  return `unexpected ${call.method} ${path}`;
}

const ALLOWED_MUTATIONS: Record<string, readonly string[]> = {
  "auth.status": [],
  "mail.list": [],
  "mail.search": [],
  "mail.get": [],
  "mail.thread": [],
  "mail.send": ["mail.send"],
  "mail.reply": ["mail.reply"],
  "mail.forward": ["mail.forward"],
  "mail.mark": ["mail.mark"],
  "mail.move": ["mail.move"],
  "mail.unsubscribe_preview": [],
  "mail.unsubscribe_apply": ["mail.send"],
  "calendar.list": [],
  "calendar.get": [],
  "calendar.create": ["calendar.create"],
  "calendar.update": ["calendar.update"],
  "calendar.delete": ["calendar.delete"],
  "calendar.respond": ["calendar.respond"],
  "calendar.instances": [],
  "availability.freebusy": [],
  "availability.suggest": [],
  "availability.snapshot": [],
  "people.search": [],
  "people.resolve": [],
  "settings.get": [],
  "settings.validate": [],
  "settings.reload": [],
  "settings.path": [],
};

const ONE_LINE_ERROR_STACK = /\n\s+at\s+[^\n]+:\d+:\d+/;
const BEARER = /Bearer\s+[A-Za-z0-9._~+/=-]{10,}/;

/* ------------------------------------------------------------------ */
/* Test lifecycle                                                      */
/* ------------------------------------------------------------------ */

let dir: string;
let harness: Harness;
const savedEnv = { ...process.env };

async function seedToken(target: string): Promise<void> {
  await writeFile(
    join(target, "token.json"),
    `${JSON.stringify(
      {
        version: 1,
        account: { upn: "me@example.com", id: "account-1" },
        tokens: {
          outlook: { access_token: "mock-outlook-token", expires_at: "2099-01-01T00:00:00.000Z" },
          graph: { access_token: "mock-graph-token", expires_at: "2099-01-01T00:00:00.000Z" },
        },
        source: "headless",
        updated_at: new Date().toISOString(),
      },
      null,
      2,
    )}\n`,
    { mode: 0o600 },
  );
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-fuzz-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  process.env.TEAMS_MCP_SNAPSHOT_DIR = join(dir, "snapshots");
  delete process.env.TEAMS_MCP_CONFIG_FILE;
  delete process.env.TEAMS_MCP_TIMEZONE;
  delete process.env.TEAMS_MCP_USERNAME;
  delete process.env.TEAMS_MCP_PASSWORD;
  delete process.env.TEAMS_MCP_TOTP_SECRET;
  // The matrix intentionally drives hundreds of writes; keep the shared bucket
  // from masking behaviour under test (rate limiting has its own suite).
  process.env.TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE = "600";
  resetConfigCache();
  resetWriteLimiter();
  clearSnapshotCache();
  clearPeopleCache();
  await seedToken(dir);
  harness = installMock();
});

afterEach(async () => {
  setOutlookClient(undefined);
  setGraphClient(undefined);
  setOneClickPoster(undefined);
  setOneClickResolver(undefined);
  setSentLookupDelays(undefined);
  clearSnapshotCache();
  clearPeopleCache();
  process.env = { ...savedEnv };
  resetConfigCache();
  resetWriteLimiter();
  await rm(dir, { recursive: true, force: true });
});

/* ------------------------------------------------------------------ */
/* 1. The fuzz matrix: every op, mixed valid/hostile args              */
/* ------------------------------------------------------------------ */

describe("fuzz matrix (seeded)", () => {
  const SEEDS = [1, 7, 42, 1337, 2026];
  const CALLS_PER_SEED = 200;

  for (const seed of SEEDS) {
    it(`survives ${CALLS_PER_SEED} mixed calls with seed ${seed}`, async () => {
      const next = random(seed);
      const opPool = shuffled(
        Array.from({ length: Math.ceil(CALLS_PER_SEED / ALL_OPS.length) }, () => ALL_OPS).flat(),
        next,
      ).slice(0, CALLS_PER_SEED);

      for (const [index, entry] of opPool.entries()) {
        const args = validArgs(entry.tool.name, entry.op);
        for (const key of Object.keys(entry.tool.shape)) {
          if (key === "op") continue;
          const roll = next();
          if (roll < 0.22) {
            const make = HOSTILE[Math.floor(next() * HOSTILE.length)] ?? (() => undefined);
            args[key] = make();
          } else if (roll < 0.4) {
            delete args[key];
          }
        }
        // Redirect `timezone` to valid zones sometimes; a fully hostile pool
        // would starve the deeper date logic.
        if ("timezone" in entry.tool.shape && next() < 0.5) args.timezone = "UTC";

        const before = harness.calls.length;
        const postersBefore = harness.posters.length;
        const call = { ...args, op: entry.op };
        let result: CallToolResult;
        try {
          result = await entry.tool.handle(call);
        } catch (error) {
          throw new Error(
            `handler threw for ${entry.tool.name}.${entry.op} (seed ${seed}, call ${index}, ${describeArgs(call)}): ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        }

        const text = textOf(result);
        expect(text, `${entry.tool.name}.${entry.op} returned no text`).not.toBe("");
        expect(text.length).toBeLessThan(300_000);
        expect(text).not.toMatch(ONE_LINE_ERROR_STACK);
        expect(text).not.toMatch(BEARER);
        if (result.isError) {
          // Every tool error is a single line, never a serialized exception.
          expect(text, `multi-line error for ${entry.tool.name}.${entry.op}`).not.toContain("\n");
        }

        const allowed = ALLOWED_MUTATIONS[`${entry.tool.name}.${entry.op}`] ?? [];
        let unexpected: string | undefined;
        for (const request of harness.calls.slice(before)) {
          const mutation = mutationOf(request);
          if (mutation && !allowed.includes(mutation)) unexpected = mutation;
        }
        expect(
          unexpected,
          `${entry.tool.name}.${entry.op} issued ${unexpected} (seed ${seed}, call ${index}, ${describeArgs(call)})`,
        ).toBeUndefined();

        if (entry.op !== "unsubscribe_apply") {
          expect(
            harness.posters.length,
            `${entry.tool.name}.${entry.op} posted an unsubscribe target`,
          ).toBe(postersBefore);
        }
      }
    }, 120_000);
  }
});

/* ------------------------------------------------------------------ */
/* 2. Regression: argument validation gaps                             */
/* ------------------------------------------------------------------ */

describe("validation regressions", () => {
  it("rejects an unknown IANA zone instead of throwing from Intl", async () => {
    for (const timezone of ["Mars/Base", "../../etc/passwd", "America/Chicago; DROP"]) {
      for (const tool of [calendarTool, availabilityTool]) {
        const result = await tool.handle({ op: tool === calendarTool ? "list" : "freebusy", timezone });
        expect(result.isError).toBe(true);
        expect(textOf(result)).toContain("invalid_argument");
        expect(result.content[0]?.type === "text" ? result.content[0].text : "").not.toMatch(ONE_LINE_ERROR_STACK);
      }
    }
    expect(harness.calls).toHaveLength(0);
  });

  it("rejects OData filter injection in since/until before any request", async () => {
    for (const since of ["2026-01-01T00:00:00Z or IsRead eq true", "'; DROP", "2026-13-01"]) {
      const result = await mailTool.handle({ op: "list", since });
      expect(result.isError).toBe(true);
      expect(textOf(result)).toContain("invalid_argument");
    }
    expect(harness.calls).toHaveLength(0);
  });

  it("canonicalises a valid date filter and refuses since > until", async () => {
    const listed = await mailTool.handle({ op: "list", since: "2026-11-16", until: "2026-11-17" });
    expect(listed.isError).toBeUndefined();
    const url = harness.calls[0]?.url ?? "";
    expect(url).toContain("ReceivedDateTime+ge+2026-11-16T06%3A00%3A00.000Z");
    expect(url).toContain("ReceivedDateTime+le+2026-11-17T06%3A00%3A00.000Z");

    harness.calls.length = 0;
    const reversed = await mailTool.handle({ op: "list", since: "2026-11-18", until: "2026-11-17" });
    expect(reversed.isError).toBe(true);
    expect(textOf(reversed)).toContain("since must not be after until");
    expect(harness.calls).toHaveLength(0);
  });

  it("rejects a multi-address to[] element that would slip past the allowlist", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@allowed.com"]\n');
    resetConfigCache();
    for (const address of [
      "attacker@evil.com, victim@allowed.com",
      "victim@allowed.com\r\nBcc: attacker@evil.com",
      "a@b@c",
      "a@allowed.com;b@evil.com",
    ]) {
      harness.calls.length = 0;
      const result = await mailTool.handle({ op: "send", to: [address], subject: "x", body: "y" });
      expect(result.isError, address).toBe(true);
      expect(textOf(result)).toContain("invalid_recipient");
      expect(harness.calls, address).toHaveLength(0);
    }
  });

  it("keeps display-name recipients and forwards the bare address", async () => {
    const result = await mailTool.handle({
      op: "send",
      to: ["Test Student <jd@allowed.com>"],
      subject: "x",
      body: "y",
    });
    expect(result.isError).toBeUndefined();
    expect(member(harness.calls[0]?.body, "Message")).toMatchObject({
      ToRecipients: [{ EmailAddress: { Address: "jd@allowed.com" } }],
    });
  });

  it("refuses an availability window larger than a year without spinning", async () => {
    const started = Date.now();
    const result = await availabilityTool.handle({
      op: "freebusy",
      attendees: ["a@example.com"],
      start: "0001-01-01",
      end: "9999-12-31",
    });
    // Generous bound: the loop this guards would run for minutes, but the assert
    // must not flake when the suite's worker pool is busy.
    expect(Date.now() - started).toBeLessThan(10_000);
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("availability is limited to 366 days");
    expect(harness.calls).toHaveLength(0);
  });

  it("rejects impossible calendar dates instead of passing NaN to Intl", async () => {
    for (const start of ["2026-13-01", "2026-02-30", "2026-00-00"]) {
      for (const tool of [calendarTool, availabilityTool]) {
        const result = await tool.handle({
          op: tool === calendarTool ? "list" : "freebusy",
          start,
          end: "2026-12-31",
        });
        expect(result.isError, start).toBe(true);
        expect(textOf(result)).toContain("invalid_argument");
      }
    }
    expect(harness.calls).toHaveLength(0);
  });

  it("caps people.resolve fan-out at 50 inputs before any request", async () => {
    const emails = Array.from({ length: 51 }, (_, index) => `u${index}@example.com`);
    const result = await peopleTool.handle({ op: "resolve", emails });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("at most 50");
    expect(harness.calls).toHaveLength(0);
  });

  it("clamps a directly supplied negative top to 1", async () => {
    await mailTool.handle({ op: "list", top: -5 });
    expect(harness.calls[0]?.url).toContain("%24top=1");
    harness.calls.length = 0;
    await calendarTool.handle({ op: "list", top: -5 });
    expect(harness.calls[0]?.url).toContain("%24top=1");
    harness.calls.length = 0;
    await peopleTool.handle({ op: "search", query: "fuzz", top: -5 });
    expect(harness.calls[0]?.url).toContain("%24top=1");
  });

  it("rejects an unknown timezone in config.yaml and in the env override", async () => {
    await expect(setSetting("timezone", "Mars/Base")).rejects.toThrow(/IANA time zone/);
    const text = await readFile(join(dir, "config.yaml"), "utf8").catch(() => "");
    expect(text).not.toContain("Mars/Base");

    await writeFile(join(dir, "config.yaml"), "timezone: Mars/Base\n");
    resetConfigCache();
    expect(() => config()).toThrow(ConfigError);

    await writeFile(join(dir, "config.yaml"), "");
    process.env.TEAMS_MCP_TIMEZONE = "Nowhere/Nope";
    resetConfigCache();
    expect(isValidTimeZone(process.env.TEAMS_MCP_TIMEZONE)).toBe(false);
    expect(() => config()).toThrow(ConfigError);
  });

  it("turns a settings error into a typed result instead of throwing", async () => {
    const unknown = await settingsTool.handle({ op: "get", key: "no.such.key" });
    expect(unknown.isError).toBe(true);
    expect(textOf(unknown)).toContain("settings_invalid");
    expect(textOf(unknown)).not.toContain("\n");

    await writeFile(join(dir, "config.yaml"), "timezone: Mars/Base\n");
    resetConfigCache();
    for (const op of ["get", "reload"]) {
      const result = await settingsTool.handle({ op });
      expect(result.isError, op).toBe(true);
      expect(textOf(result), op).toContain("config_invalid");
    }
  });
});

/** Guarded object member read for assertions on parsed JSON. */
function member(value: unknown, key: string): unknown {
  if (value === null || typeof value !== "object" || !Object.hasOwn(value, key)) return undefined;
  const record = value as Record<string, unknown>; // guarded: own property exists
  return record[key];
}

/* ------------------------------------------------------------------ */
/* 3. State machines                                                   */
/* ------------------------------------------------------------------ */

describe("state machines", () => {
  it("create -> guarded update/delete: a wrong expected_subject never mutates", async () => {
    const created = payloadOf(
      await calendarTool.handle({
        op: "create",
        subject: "Fuzz Event",
        start: "2026-11-18T16:00:00Z",
        end: "2026-11-18T17:00:00Z",
      }),
    );
    expect(created).toMatchObject({ ok: true, op: "create", id: "created-1" });

    harness.calls.length = 0;
    const refusedUpdate = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "A different event",
      subject: "Renamed",
    });
    expect(refusedUpdate.isError).toBe(true);
    expect(textOf(refusedUpdate)).toContain("subject_mismatch");
    expect(harness.calls.map((call) => call.method)).toEqual(["GET"]);

    harness.calls.length = 0;
    const refusedDelete = await calendarTool.handle({
      op: "delete",
      event_id: "evt-1",
      expected_subject: "Nope",
    });
    expect(refusedDelete.isError).toBe(true);
    expect(harness.calls.map((call) => call.method)).toEqual(["GET"]);

    harness.calls.length = 0;
    const updated = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: " fuzz   EVENT ",
      subject: "Renamed",
    });
    expect(updated.isError).toBeUndefined();
    expect(harness.calls.map((call) => call.method)).toEqual(["GET", "PATCH"]);

    harness.calls.length = 0;
    const deleted = await calendarTool.handle({
      op: "delete",
      event_id: "evt-1",
      expected_subject: "Fuzz Event",
    });
    expect(deleted.isError).toBeUndefined();
    expect(harness.calls.map((call) => call.method)).toEqual(["GET", "DELETE"]);
  });

  it("reply returns the sent copy id, and forward reuses the conversation", async () => {
    const reply = payloadOf(await mailTool.handle({ op: "reply", message_id: "msg-1", body: "thanks" }));
    expect(reply).toMatchObject({ ok: true, op: "reply", id: "sent-1", recipient_count: 1 });

    harness.calls.length = 0;
    const forward = payloadOf(
      await mailTool.handle({ op: "forward", message_id: "msg-1", to: ["a@example.com"] }),
    );
    expect(forward).toMatchObject({ ok: true, op: "forward", id: "sent-1" });
    expect(harness.calls.some((call) => call.method === "POST" && call.url.includes("/forward"))).toBe(true);
  });

  it("move returns the new id and the new id is usable afterwards", async () => {
    const moved = payloadOf(
      await mailTool.handle({ op: "move", message_id: "msg-1", destination: "archive" }),
    );
    expect(moved).toMatchObject({ ok: true, id: "moved-1", from_id: "msg-1" });

    harness.calls.length = 0;
    const read = payloadOf(await mailTool.handle({ op: "get", message_id: String(moved.id) }));
    expect(read.id).toBe("msg-1");
    expect(harness.calls[0]?.url).toContain("/me/messages/moved-1");
  });

  it("coalesces parallel snapshot refreshes and leaves no temp files", async () => {
    const calls = Array.from({ length: 20 }, () =>
      availabilityTool.handle({
        op: "snapshot",
        attendees: ["a@example.com"],
        start: "2026-11-17",
        end: "2026-11-18",
      }),
    );
    const results = await Promise.all(calls);
    for (const result of results) expect(result.isError).toBeUndefined();
    const refreshed = results.filter((result) => payloadOf(result).refreshed === true);
    expect(refreshed).toHaveLength(1);

    const loaded = await readSnapshot(snapshotPath());
    expect(loaded).not.toBeNull();
    const names = await readdir(join(dir, "snapshots"));
    expect(names.filter((name) => name.endsWith(".tmp"))).toEqual([]);
    expect(names).toContain("calendar.yaml");
  });

  it("settings: invalid -> valid -> reload, and parallel writes all persist", async () => {
    const invalid = await settingsTool.handle({ op: "set", key: "timezone", value: "Mars/Base" });
    expect(invalid.isError).toBe(true);
    expect(textOf(invalid)).toContain("settings_invalid");

    const valid = await settingsTool.handle({ op: "set", key: "timezone", value: "Europe/Berlin" });
    expect(valid.isError).toBeUndefined();
    const reloaded = await settingsTool.handle({ op: "reload" });
    expect(payloadOf(reloaded).settings).toMatchObject({ timezone: "Europe/Berlin" });

    const keys = Array.from({ length: 30 }, (_, index) => `alias${index}`);
    const writes = await Promise.all(
      keys.map((key) =>
        settingsTool.handle({ op: "set", key: `aliases.${key}`, value: `user-${key}@example.com` }),
      ),
    );
    for (const write of writes) expect(write.isError).toBeUndefined();
    const text = await readFile(join(dir, "config.yaml"), "utf8");
    expect((text.match(/^ {2}alias\d+:/gm) ?? [])).toHaveLength(30);
  });

  it("surfaces a corrupt token cache without launching a login", async () => {
    await writeFile(join(dir, "token.json"), "{ not json");
    const status = payloadOf(await authTool.handle({ op: "status" }));
    expect(status).toMatchObject({ authenticated: false, corrupt: true });

    // Use the real token provider (no injected mock) so the corrupt-cache path
    // runs: it must fail with auth_required, never launch the browser flow.
    setOutlookClient(undefined);
    const before = harness.calls.length;
    const listed = await mailTool.handle({ op: "list" });
    expect(listed.isError).toBe(true);
    expect(textOf(listed)).toContain("auth_required");
    expect(harness.calls.length).toBe(before);
  });

  it("auth.login without credentials errors, and auth.logout only touches the temp cache", async () => {
    const login = await authTool.handle({ op: "login" });
    expect(login.isError).toBe(true);
    expect(textOf(login)).toContain("credentials_missing");

    const logout = payloadOf(await authTool.handle({ op: "logout" }));
    expect(logout.logged_out).toBe(true);
    expect(String(logout.token_cache)).toContain(dir);
    const status = payloadOf(await authTool.handle({ op: "status" }));
    expect(status.authenticated).toBe(false);
    expect(status.profile).toContain(dir);
  });

  it("keeps the people cache bounded under many distinct queries", async () => {
    clearPeopleCache();
    for (let index = 0; index < 150; index += 1) {
      const result = await peopleTool.handle({ op: "search", query: `person-${index}`, top: 1 });
      expect(result.isError).toBeUndefined();
    }
    const before = harness.calls.length;
    await peopleTool.handle({ op: "search", query: "person-0", top: 1 });
    expect(harness.calls.length).toBe(before + 1);
  });
});

/* ------------------------------------------------------------------ */
/* 4. Resilience                                                       */
/* ------------------------------------------------------------------ */

describe("resilience", () => {
  it("handles 20 parallel mixed reads without unhandled rejections or writes", async () => {
    const calls = [
      () => mailTool.handle({ op: "list", top: 2 }),
      () => mailTool.handle({ op: "search", query: "fuzz", top: 2 }),
      () => mailTool.handle({ op: "thread", thread_id: "conv-1", top: 2 }),
      () => calendarTool.handle({ op: "list", top: 2 }),
      () => calendarTool.handle({ op: "get", event_id: "evt-1" }),
      () => calendarTool.handle({ op: "instances", event_id: "evt-1", start: "2026-11-17", end: "2026-11-18" }),
      () => availabilityTool.handle({ op: "freebusy", attendees: ["a@example.com"], start: "2026-11-17", end: "2026-11-18" }),
      () => availabilityTool.handle({ op: "suggest", attendees: ["a@example.com"], start: "2026-11-17", end: "2026-11-18" }),
      () => peopleTool.handle({ op: "search", query: "fuzz" }),
      () => peopleTool.handle({ op: "resolve", emails: ["a@example.com"] }),
      () => authTool.handle({ op: "status" }),
      () => settingsTool.handle({ op: "get" }),
    ];
    const results = await Promise.all(Array.from({ length: 20 }, (_, index) => calls[index % calls.length]?.()));
    for (const result of results) {
      expect(result?.isError).toBeUndefined();
      expect(result?.content[0]?.type).toBe("text");
    }
    expect(harness.calls.every((call) => call.method === "GET" || call.method === "POST")).toBe(true);
    expect(harness.calls.filter((call) => mutationOf(call))).toEqual([]);
  });

  it("drops malformed upstream rows instead of failing the page", async () => {
    installMock({
      outlook: async (url) => {
        if (url.includes("calendarview")) {
          return json({ value: [null, 42, "x", eventRow("evt-ok")] });
        }
        if (url.includes("/messages")) {
          return json({ value: [null, [], messageRow("msg-ok")] });
        }
        if (url.includes("getschedule")) {
          return json({ value: [null, { ScheduleId: "a@example.com", ScheduleItems: [] }] });
        }
        return json({ value: [] });
      },
    });
    const events = payloadOf(await calendarTool.handle({ op: "list", top: 5 }));
    expect(Array.isArray(events.events) ? events.events.length : -1).toBe(1);
    const messages = payloadOf(await mailTool.handle({ op: "list", top: 5 }));
    expect(Array.isArray(messages.messages) ? messages.messages.length : -1).toBe(1);
    const busy = await availabilityTool.handle({
      op: "freebusy",
      attendees: ["a@example.com"],
      start: "2026-11-17",
      end: "2026-11-18",
    });
    expect(busy.isError).toBeUndefined();

    harness.calls.length = 0;
    const partial = await calendarTool.handle({ op: "get", event_id: "evt-empty" });
    expect(partial.isError).toBeUndefined();
  });

  it("handles 100 KB and deeply nested arguments without header injection", async () => {
    const payload = `line one\r\nInjected: yes\u0000${"x".repeat(100_000)}`;
    const sent = await mailTool.handle({
      op: "send",
      to: ["a@example.com"],
      subject: payload,
      body: payload,
      extra: deepObject(200),
    });
    expect(sent.isError).toBeUndefined();
    const request = harness.calls[0];
    expect(member(member(request?.body, "Message"), "Subject")).toBe(payload);
    expect(JSON.stringify(request?.headers)).not.toContain("\r\n");

    const created = await calendarTool.handle({
      op: "create",
      subject: payload,
      start: "2026-11-18T16:00:00Z",
      end: "2026-11-18T17:00:00Z",
    });
    expect(created.isError).toBeUndefined();
  });
});

function deepObject(depth: number): Record<string, unknown> {
  let value: Record<string, unknown> = { leaf: true };
  for (let index = 0; index < depth; index += 1) value = { child: value };
  return value;
}

/* ------------------------------------------------------------------ */
/* 5. Explicit schema coverage                                         */
/* ------------------------------------------------------------------ */

describe("schema surfaces", () => {
  it("rejects wrong types for every declared argument and accepts the op enum", () => {
    for (const tool of TOOLS) {
      expect(tool.shape.op?.safeParse("definitely-not-an-op").success).toBe(false);
      for (const [key, schema] of Object.entries(tool.shape)) {
        if (key === "op") continue;
        // `null` is never a valid tool argument.
        expect(schema.safeParse(null).success, `${tool.name}.${key}`).toBe(false);
      }
    }
  });

  it("is deterministic: the same seed produces the same generated calls", () => {
    const fingerprint = (seed: number): string => {
      const next = random(seed);
      const values: string[] = [];
      for (let index = 0; index < 50; index += 1) {
        const make = HOSTILE[Math.floor(next() * HOSTILE.length)] ?? (() => undefined);
        values.push(JSON.stringify(make() ?? null).slice(0, 40));
      }
      return values.join("|");
    };
    expect(fingerprint(42)).toBe(fingerprint(42));
    expect(fingerprint(42)).not.toBe(fingerprint(43));
  });
});
