#!/usr/bin/env node
/**
 * teams-mcp fuzz probe: seeded adversarial calls over real stdio.
 *
 *   node scripts/dev/fuzz-probe.mjs [--seed 1337] [--calls 200] [--live]
 *
 * Default mode starts a local mock Outlook/Graph server, points the child at it
 * with an isolated config/cache dir and a fabricated token, and runs every
 * mutating op there. The mock request log is asserted, so a write can never be
 * mistaken for a read. It also spawns short-lived children with
 * TEAMS_MCP_READ_ONLY=1, unsubscribe disabled, and a 1/minute write limit: each
 * blocked call must be refused AND leave the mock log untouched.
 *
 * `--live` uses the canonical config and ONLY the read ops listed in LIVE_OPS:
 * no send/reply/forward/mark/move, no calendar create/update/delete/respond, no
 * settings.set, no auth.login/logout, no unsubscribe_apply. Anything that could
 * mutate the account is absent from the live pool by construction. It adds two
 * read-only controls: `mail.list from=<sender>` (no InefficientFilter) and a
 * fake mailbox that must report `unknown`, never free.
 *
 * Exit code: 0 all checks passed, 1 any check failed.
 */
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const ROOT = new URL("../..", import.meta.url).pathname;
const argv = process.argv.slice(2);

function flag(name, fallback) {
  const index = argv.indexOf(name);
  return index >= 0 ? argv[index + 1] : fallback;
}

const SEED = Number(flag("--seed", "1337"));
const CALLS = Math.max(1, Math.min(2000, Number(flag("--calls", "200"))));
const LIVE = argv.includes("--live");
const HELP = argv.includes("--help") || argv.includes("-h");

if (HELP) {
  console.log(`usage: node scripts/dev/fuzz-probe.mjs [--seed N] [--calls N] [--live]

  --seed N    deterministic PRNG seed (default 1337)
  --calls N   random calls in the mock phase (default 200)
  --live      read-only calls against the canonical account (default off)

Mock mode is the default and is the only mode that exercises writes.`);
  process.exit(0);
}

let failures = 0;
let checks = 0;

function check(name, condition, detail) {
  checks += 1;
  if (!condition) failures += 1;
  const suffix = detail === undefined ? "" : ` — ${typeof detail === "string" ? detail : JSON.stringify(detail)}`;
  console.log(`${condition ? "PASS" : "FAIL"}  ${name}${suffix}`);
}

/* ------------------------------------------------------------------ */
/* Deterministic randomness                                            */
/* ------------------------------------------------------------------ */

function random(seed) {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

const next = random(SEED);

function pick(values) {
  return values[Math.floor(next() * values.length)];
}

function shuffled(values) {
  const copy = [...values];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swap = Math.floor(next() * (index + 1));
    const current = copy[index];
    copy[index] = copy[swap];
    copy[swap] = current;
  }
  return copy;
}

/* ------------------------------------------------------------------ */
/* Mock Outlook REST + Graph                                           */
/* ------------------------------------------------------------------ */

const mock = { log: [], server: undefined, base: undefined };

function messageRow(id = "msg-1") {
  return {
    Id: id,
    ConversationId: "conv-1",
    Subject: "Fuzz message",
    From: { EmailAddress: { Name: "Sender", Address: "sender@example.com" } },
    ToRecipients: [{ EmailAddress: { Address: "me@example.com" } }],
    ReceivedDateTime: "2026-11-16T12:00:00Z",
    IsRead: true,
    BodyPreview: "preview text",
    Body: { ContentType: "HTML", Content: "<p>body</p>" },
    WebLink: "https://outlook.office365.com/owa/?ItemID=msg-1",
  };
}

function eventRow(id = "evt-1", subject = "Fuzz Event") {
  return {
    Id: id,
    Subject: subject,
    Start: { DateTime: "2026-11-18T16:00:00.0000000", TimeZone: "UTC" },
    End: { DateTime: "2026-11-18T17:00:00.0000000", TimeZone: "UTC" },
    IsAllDay: false,
    ShowAs: "Busy",
    Organizer: { EmailAddress: { Address: "me@example.com" } },
    Attendees: [
      { EmailAddress: { Address: "a@example.com" }, Type: "Required", Status: { Response: "Accepted" } },
    ],
    WebLink: "https://outlook.office365.com/owa/?itemid=evt-1",
    BodyPreview: "event preview",
  };
}

function userRow(id = "user-1") {
  return {
    id,
    displayName: "Fuzz User",
    mail: "fuzz@example.com",
    userPrincipalName: "fuzz@example.com",
    jobTitle: "Tester",
    department: "QA",
  };
}

function send(res, status, body) {
  const text = body === undefined ? "" : JSON.stringify(body);
  res.writeHead(status, { "content-type": "application/json", "content-length": Buffer.byteLength(text) });
  res.end(text);
}

function outlookResponse(method, url, body) {
  const path = new URL(url, "http://mock").pathname;
  if (method === "POST" && path.endsWith("/me/sendmail")) return [202, undefined];
  if (method === "POST" && /\/me\/messages\/[^/]+\/(reply|forward)$/.test(path)) return [202, undefined];
  if (method === "POST" && /\/me\/messages\/[^/]+\/move$/.test(path)) return [200, { Id: "moved-1" }];
  if (method === "PATCH" && /\/me\/messages\/[^/]+$/.test(path)) return [200, messageRow("patched-1")];
  if (method === "GET" && /\/me\/mailfolders\/[^/]+\/messages$/.test(path)) {
    return [200, { value: [messageRow()] }];
  }
  if (method === "GET" && path.endsWith("/me/messages")) return [200, { value: [messageRow()] }];
  if (method === "GET" && /\/me\/messages\/[^/]+$/.test(path)) return [200, messageRow()];
  if (method === "GET" && path.endsWith("/me/calendarview")) return [200, { value: [eventRow()] }];
  if (method === "POST" && path.endsWith("/me/events")) return [201, eventRow("created-1")];
  if (method === "GET" && /\/me\/events\/[^/]+\/instances$/.test(path)) return [200, { value: [eventRow()] }];
  if (method === "POST" && /\/me\/events\/[^/]+\/(accept|decline|tentativelyaccept)$/.test(path)) {
    return [202, undefined];
  }
  if (method === "PATCH" && /\/me\/events\/[^/]+$/.test(path)) return [200, eventRow()];
  if (method === "DELETE" && /\/me\/events\/[^/]+$/.test(path)) return [204, undefined];
  if (method === "GET" && /\/me\/events\/[^/]+$/.test(path)) return [200, eventRow()];
  if (method === "POST" && path.endsWith("/me/calendar/getschedule")) {
    const schedules = Array.isArray(body?.Schedules) ? body.Schedules.filter((value) => typeof value === "string") : [];
    return [
      200,
      {
        value: schedules.map((email) => ({
          ScheduleId: email,
          AvailabilityView: "0",
          WorkingHours: { StartTime: "09:00", EndTime: "17:00" },
          ScheduleItems: [],
        })),
      },
    ];
  }
  if (method === "POST" && path.endsWith("/me/findmeetingtimes")) return [200, { MeetingTimeSuggestions: [] }];
  return [404, { error: { code: "NotFound", message: `unmocked ${method} ${path}` } }];
}

function graphResponse(method, url) {
  const parsed = new URL(url, "http://mock");
  if (method !== "GET") return [404, { error: { code: "NotFound", message: "read-only mock" } }];
  if (/\/users\/[^/]+$/.test(parsed.pathname)) return [200, userRow()];
  if (parsed.pathname.endsWith("/users")) return [200, { value: [userRow()] }];
  return [404, { error: { code: "NotFound", message: "unmocked graph path" } }];
}

async function startMock() {
  const server = createServer((req, res) => {
    const chunks = [];
    req.on("data", (chunk) => {
      if (chunks.reduce((size, part) => size + part.length, 0) > 20_000_000) return;
      chunks.push(chunk);
    });
    req.on("end", () => {
      const raw = Buffer.concat(chunks).toString("utf8");
      let body;
      try {
        body = raw ? JSON.parse(raw) : undefined;
      } catch {
        body = undefined;
      }
      const entry = { method: req.method ?? "GET", url: req.url ?? "/", body };
      mock.log.push(entry);
      const [status, payload] = (req.url ?? "").startsWith("/api/v2.0")
        ? outlookResponse(entry.method, `http://mock${entry.url}`, body)
        : graphResponse(entry.method, `http://mock${entry.url}`);
      send(res, status, payload);
    });
  });
  const listening = Promise.withResolvers();
  server.listen(0, "127.0.0.1", () => listening.resolve());
  await listening.promise;
  const address = server.address();
  if (address === null || typeof address === "string") throw new Error("mock server has no port");
  mock.server = server;
  mock.base = `http://127.0.0.1:${address.port}`;
}

/* ------------------------------------------------------------------ */
/* stdio JSON-RPC client                                               */
/* ------------------------------------------------------------------ */

function connect(env) {
  const child = spawn(process.execPath, ["dist/index.js"], {
    cwd: ROOT,
    env: { ...process.env, ...env },
    stdio: ["pipe", "pipe", "pipe"],
  });
  const pending = new Map();
  const frames = [];
  const badLines = [];
  let buffered = "";
  let stderr = "";
  let nextId = 1;

  child.stdout.setEncoding("utf8");
  child.stdout.on("data", (chunk) => {
    buffered += chunk;
    for (;;) {
      const newline = buffered.indexOf("\n");
      if (newline < 0) break;
      const line = buffered.slice(0, newline).replace(/\r$/, "");
      buffered = buffered.slice(newline + 1);
      if (line.trim() === "") continue;
      let frame;
      try {
        frame = JSON.parse(line);
      } catch {
        badLines.push(line.slice(0, 160));
        continue;
      }
      if (frame.jsonrpc !== "2.0") {
        badLines.push(line.slice(0, 160));
        continue;
      }
      frames.push(frame);
      if (typeof frame.id === "number" && pending.has(frame.id)) {
        pending.get(frame.id)(frame);
        pending.delete(frame.id);
      }
    }
  });
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
  });

  const send = (payload) => child.stdin.write(`${JSON.stringify(payload)}\n`);
  const request = (method, params, timeoutMs = 20_000) => {
    const id = nextId++;
    const { promise, resolve } = Promise.withResolvers();
    // Subprocess boundary: a real guard timer is the only way to fail fast.
    const timer = setTimeout(() => {
      pending.delete(id);
      resolve(undefined);
    }, timeoutMs);
    pending.set(id, (frame) => {
      clearTimeout(timer);
      resolve(frame);
    });
    send({ jsonrpc: "2.0", id, method, ...(params === undefined ? {} : { params }) });
    return promise;
  };
  const notify = (method, params) => send({ jsonrpc: "2.0", method, ...(params === undefined ? {} : { params }) });
  const raw = (text) => child.stdin.write(text);
  const close = async () => {
    child.stdin.end();
    const { promise, resolve } = Promise.withResolvers();
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      resolve();
    }, 5_000);
    child.once("exit", () => {
      clearTimeout(timer);
      resolve();
    });
    await promise;
  };

  return { child, frames, badLines, stderr: () => stderr, request, notify, raw, close };
}

function textOf(frame) {
  const content = frame?.result?.content;
  return Array.isArray(content) && content[0]?.type === "text" ? String(content[0].text ?? "") : "";
}

async function callTool(client, name, args) {
  const frame = await client.request("tools/call", { name, arguments: args });
  return { frame, text: textOf(frame), isError: frame?.result?.isError === true };
}

/* ------------------------------------------------------------------ */
/* Op pools and argument generation                                    */
/* ------------------------------------------------------------------ */

const MOCK_OPS = [
  { tool: "auth", op: "status" },
  ...["list", "search", "get", "thread", "send", "reply", "forward", "mark", "move", "unsubscribe_preview", "unsubscribe_apply"].map(
    (op) => ({ tool: "mail", op }),
  ),
  ...["list", "get", "create", "update", "delete", "respond", "instances"].map((op) => ({
    tool: "calendar",
    op,
  })),
  ...["freebusy", "suggest", "snapshot"].map((op) => ({ tool: "availability", op })),
  ...["search", "resolve"].map((op) => ({ tool: "people", op })),
  ...["get", "set", "validate", "reload", "path"].map((op) => ({ tool: "settings", op })),
];

/**
 * Live mode: read-only. Notably absent: mail send/reply/forward/mark/move/
 * unsubscribe_apply, calendar create/update/delete/respond, settings.set,
 * auth.login/logout.
 */
const LIVE_OPS = [
  { tool: "auth", op: "status" },
  { tool: "mail", op: "list" },
  { tool: "mail", op: "search" },
  { tool: "mail", op: "get" },
  { tool: "mail", op: "thread" },
  { tool: "mail", op: "unsubscribe_preview" },
  { tool: "calendar", op: "list" },
  { tool: "calendar", op: "get" },
  { tool: "calendar", op: "instances" },
  { tool: "availability", op: "freebusy" },
  { tool: "availability", op: "suggest" },
  { tool: "people", op: "search" },
  { tool: "people", op: "resolve" },
  { tool: "settings", op: "get" },
  { tool: "settings", op: "validate" },
  { tool: "settings", op: "path" },
];

const VALID_ARGS = {
  "auth.status": {},
  "mail.list": { top: 3 },
  "mail.search": { query: "fuzz", top: 3 },
  "mail.get": { message_id: "msg-1" },
  "mail.thread": { thread_id: "conv-1", top: 3 },
  "mail.send": { to: ["a@example.com"], subject: "Fuzz send", body: "body" },
  "mail.reply": { message_id: "msg-1", body: "reply" },
  "mail.forward": { message_id: "msg-1", to: ["a@example.com"] },
  "mail.mark": { message_id: "msg-1", mark: "read" },
  "mail.move": { message_id: "msg-1", destination: "archive" },
  "mail.unsubscribe_preview": { message_id: "msg-1" },
  "mail.unsubscribe_apply": { message_id: "msg-1" },
  "calendar.list": { top: 3 },
  "calendar.get": { event_id: "evt-1" },
  "calendar.create": { subject: "Fuzz Event", start: "2026-11-18T16:00:00Z", end: "2026-11-18T17:00:00Z" },
  "calendar.update": { event_id: "evt-1", expected_subject: "Fuzz Event", subject: "Renamed" },
  "calendar.delete": { event_id: "evt-1", expected_subject: "Fuzz Event" },
  "calendar.respond": { event_id: "evt-1", expected_subject: "Fuzz Event", response: "accept" },
  "calendar.instances": { event_id: "evt-1", start: "2026-11-17", end: "2026-11-18" },
  "availability.freebusy": { attendees: ["a@example.com"], start: "2026-11-17", end: "2026-11-18" },
  "availability.suggest": { attendees: ["a@example.com"], start: "2026-11-17", end: "2026-11-18" },
  "availability.snapshot": { attendees: ["a@example.com"], start: "2026-11-17", end: "2026-11-18" },
  "people.search": { query: "fuzz", top: 3 },
  "people.resolve": { emails: ["a@example.com"] },
  "settings.get": {},
  "settings.set": { key: "aliases.fuzz", value: "fuzz@example.com" },
  "settings.validate": {},
  "settings.reload": {},
  "settings.path": {},
};

const HOSTILE = [
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
  () => "2026-13-01",
  () => "Mars/Base",
  () => "0001-01-01",
  () => "9999-12-31",
  () => [],
  () => [null],
  () => ["a@b@c"],
  () => ["a\r\nBcc: b@example.com"],
  () => Array.from({ length: 1000 }, (_, index) => `u${index}@example.com`),
  () => ({}),
  () => ({ nested: { deeper: { deepest: true } } }),
  () => true,
  () => "__proto__",
];

const SHAPE_KEYS = {
  auth: ["audience", "force"],
  mail: [
    "folder",
    "query",
    "top",
    "cursor",
    "unread_only",
    "from",
    "since",
    "until",
    "body_format",
    "body_chars",
    "include_body_preview",
    "include_web_link",
    "include_unsubscribe",
    "include_unsubscribe_targets",
    "message_id",
    "thread_id",
    "to",
    "cc",
    "bcc",
    "subject",
    "body",
    "importance",
    "mark",
    "destination",
    "confirm",
  ],
  calendar: [
    "event_id",
    "start",
    "end",
    "timezone",
    "top",
    "cursor",
    "brief",
    "include_web_link",
    "subject",
    "body",
    "location",
    "attendees",
    "all_day",
    "show_as",
    "sensitivity",
    "response",
    "comment",
    "expected_subject",
  ],
  availability: [
    "attendees",
    "start",
    "end",
    "duration_min",
    "interval_min",
    "buffer_min",
    "working_hours_only",
    "max_candidates",
    "timezone",
    "refresh",
  ],
  people: ["query", "ids", "emails", "top", "refresh"],
  settings: ["key", "value"],
};

function randomArgs(tool, op) {
  const args = { ...(VALID_ARGS[`${tool}.${op}`] ?? {}) };
  for (const key of SHAPE_KEYS[tool] ?? []) {
    const roll = next();
    if (roll < 0.22) args[key] = pick(HOSTILE)();
    else if (roll < 0.4) delete args[key];
  }
  return args;
}

/* ------------------------------------------------------------------ */
/* Write classification (mock assertions)                              */
/* ------------------------------------------------------------------ */

function writeKind(entry) {
  if (entry.method === "GET") return undefined;
  const path = new URL(entry.url, "http://mock").pathname;
  if (path.endsWith("/me/calendar/getschedule") || path.endsWith("/me/findmeetingtimes")) return undefined;
  if (entry.method === "POST" && path.endsWith("/me/sendmail")) return "mail.send";
  if (entry.method === "POST" && /\/me\/messages\/[^/]+\/reply$/.test(path)) return "mail.reply";
  if (entry.method === "POST" && /\/me\/messages\/[^/]+\/forward$/.test(path)) return "mail.forward";
  if (entry.method === "POST" && /\/me\/messages\/[^/]+\/move$/.test(path)) return "mail.move";
  if (entry.method === "PATCH" && /\/me\/messages\/[^/]+$/.test(path)) return "mail.mark";
  if (entry.method === "POST" && path.endsWith("/me/events")) return "calendar.create";
  if (entry.method === "PATCH" && /\/me\/events\/[^/]+$/.test(path)) return "calendar.update";
  if (entry.method === "DELETE" && /\/me\/events\/[^/]+$/.test(path)) return "calendar.delete";
  if (entry.method === "POST" && /\/me\/events\/[^/]+\/(accept|decline|tentativelyaccept)$/.test(path)) {
    return "calendar.respond";
  }
  return `unexpected ${entry.method} ${path}`;
}

const ALLOWED_WRITES = {
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
  "mail.unsubscribe_apply": [],
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
  "settings.set": [],
  "settings.validate": [],
  "settings.reload": [],
  "settings.path": [],
};

const STACK = /\n\s+at\s+[^\n]+:\d+:\d+/;
const BEARER = /Bearer\s+[A-Za-z0-9._~+/=-]{10,}/;

function silentCheck(name, condition, detail) {
  checks += 1;
  if (condition) return;
  failures += 1;
  const suffix = detail === undefined ? "" : ` — ${typeof detail === "string" ? detail : JSON.stringify(detail)}`;
  console.log(`FAIL  ${name}${suffix}`);
}

/**
 * Assert one tool response is safe, and no write outside the op's allowance.
 * `quiet` aggregates the random phase instead of printing four lines per call.
 */
function assertSafeCall(label, call, logBefore, opKey = label, quiet = false) {
  const report = quiet ? silentCheck : check;
  report(`${label}: response received`, call.frame !== undefined);
  report(`${label}: no stack in tool text`, !STACK.test(call.text));
  report(`${label}: no bearer leak`, !BEARER.test(call.text));
  // The SDK's own -32602 schema rejection serializes the Zod issues as a
  // multi-line JSON array before our handler runs; tool errors from our code
  // are always one line. Require one line only for the latter.
  const sdkValidation = call.text.startsWith("MCP error -32602");
  if (call.isError) {
    report(`${label}: one-line error`, sdkValidation || !call.text.includes("\n"), call.text.slice(0, 100));
  }
  const produced = mock.log.slice(logBefore).map(writeKind).filter(Boolean);
  const allowed = ALLOWED_WRITES[opKey] ?? [];
  for (const kind of produced) {
    report(`${label}: write ${kind} is allowed for the op`, kind.startsWith("unexpected") || allowed.includes(kind), kind);
  }
}

/* ------------------------------------------------------------------ */
/* Mock phase                                                          */
/* ------------------------------------------------------------------ */

async function mockPhase() {
  const base = await mkdtemp(join(tmpdir(), "teams-mcp-fuzz-probe-"));
  const configDir = join(base, "config");
  await mkdir(configDir, { recursive: true });
  await writeFile(
    join(configDir, "token.json"),
    `${JSON.stringify({
      version: 1,
      account: { upn: "fuzz@example.com" },
      tokens: {
        outlook: { access_token: "fuzz-outlook-token", expires_at: "2099-01-01T00:00:00.000Z" },
        graph: { access_token: "fuzz-graph-token", expires_at: "2099-01-01T00:00:00.000Z" },
      },
      source: "headless",
      updated_at: new Date().toISOString(),
    })}\n`,
    { mode: 0o600 },
  );
  await startMock();
  const client = connect({
    TEAMS_MCP_CONFIG_DIR: configDir,
    TEAMS_MCP_CACHE_DIR: join(base, "cache"),
    TEAMS_MCP_PROFILE_DIR: join(base, "profile"),
    TEAMS_MCP_SNAPSHOT_DIR: join(base, "snapshots"),
    TEAMS_MCP_OUTLOOK_BASE_URL: `${mock.base}/api/v2.0`,
    TEAMS_MCP_GRAPH_BASE_URL: `${mock.base}/v1.0`,
    TEAMS_MCP_MAX_ATTEMPTS: "1",
    TEAMS_MCP_REQUEST_TIMEOUT_MS: "5000",
  });

  try {
    const init = await client.request("initialize", {
      protocolVersion: "2025-06-18",
      capabilities: { tools: {}, resources: {} },
      clientInfo: { name: "fuzz-probe", version: "0.0.0" },
    });
    check("initialize answers", init?.result?.serverInfo?.name === "teams-mcp");
    client.notify("notifications/initialized");

    /* Protocol abuse */
    for (const line of ["{", "[", "null", '{"jsonrpc":"2.0"}', "not json", '{"jsonrpc":"2.0","id":']) {
      client.raw(`${line}\n`);
    }
    client.raw('[{"jsonrpc":"2.0","id":900,"method":"ping"}]\n');
    check("unknown method is a JSON-RPC error", (await client.request("bogus/method"))?.error !== undefined);
    check(
      "malformed args set isError",
      (await callTool(client, "mail", { op: "send", to: "a@example.com", subject: "x", body: "y" })).isError,
    );
    check("stdout is JSON-RPC only after malformed input", client.badLines.length === 0, client.badLines);

    /* Explicit write smoke: every mutating op through the mock */
    const writes = [
      ["mail.send", { op: "send", to: ["a@example.com"], subject: "AI TEST — fuzz probe", body: "body" }],
      ["mail.reply", { op: "reply", message_id: "msg-1", body: "reply" }],
      ["mail.forward", { op: "forward", message_id: "msg-1", to: ["a@example.com"] }],
      ["mail.mark", { op: "mark", message_id: "msg-1", mark: "read" }],
      ["mail.move", { op: "move", message_id: "msg-1", destination: "archive" }],
      [
        "calendar.create",
        { op: "create", subject: "AI TEST — teams-mcp — fuzz", start: "2026-11-18T16:00:00Z", end: "2026-11-18T17:00:00Z" },
      ],
      ["calendar.update", { op: "update", event_id: "evt-1", expected_subject: "Fuzz Event", subject: "Renamed" }],
      ["calendar.delete", { op: "delete", event_id: "evt-1", expected_subject: "Fuzz Event" }],
      ["calendar.respond", { op: "respond", event_id: "evt-1", expected_subject: "Fuzz Event", response: "accept" }],
    ];
    for (const [label, args] of writes) {
      const before = mock.log.length;
      const call = await callTool(client, label.split(".")[0], args);
      assertSafeCall(label, call, before);
      check(`${label}: mock saw the write`, mock.log.slice(before).some((entry) => writeKind(entry) === label));
    }

    /* expected_subject guard: a mismatch must not PATCH or DELETE */
    for (const op of ["update", "delete"]) {
      const before = mock.log.length;
      const call = await callTool(client, "calendar", {
        op,
        event_id: "evt-1",
        expected_subject: "Definitely wrong",
        ...(op === "update" ? { subject: "Renamed" } : {}),
      });
      check(`calendar.${op} guard: refused`, call.isError && call.text.includes("subject_mismatch"));
      check(
        `calendar.${op} guard: only GET reached the mock`,
        mock.log.slice(before).every((entry) => entry.method === "GET"),
        mock.log.slice(before).map((entry) => `${entry.method} ${entry.url}`),
      );
    }

    /* Settings: invalid timezone refused; valid alias written to the temp config */
    const invalidSetting = await callTool(client, "settings", {
      op: "set",
      key: "timezone",
      value: "Mars/Base",
    });
    check("settings.set refuses an invalid timezone", invalidSetting.isError);
    const validSetting = await callTool(client, "settings", {
      op: "set",
      key: "aliases.probe",
      value: "probe@example.com",
    });
    check("settings.set writes an alias in the temp config", !validSetting.isError);
    const configText = await readFile(join(configDir, "config.yaml"), "utf8");
    check("temp config contains the alias", configText.includes("probe: probe@example.com"));

    /* Capability scoping: read-only, unsubscribe opt-in, rate limit.
       Each blocked call must be visible to the model AND absent from the mock log. */
    const safetyEnv = {
      TEAMS_MCP_CONFIG_DIR: configDir,
      TEAMS_MCP_CACHE_DIR: join(base, "cache"),
      TEAMS_MCP_PROFILE_DIR: join(base, "profile"),
      TEAMS_MCP_SNAPSHOT_DIR: join(base, "snapshots"),
      TEAMS_MCP_OUTLOOK_BASE_URL: `${mock.base}/api/v2.0`,
      TEAMS_MCP_GRAPH_BASE_URL: `${mock.base}/v1.0`,
      TEAMS_MCP_MAX_ATTEMPTS: "1",
      TEAMS_MCP_REQUEST_TIMEOUT_MS: "5000",
    };
    const openChild = async (extra) => {
      const child = connect({ ...safetyEnv, ...extra });
      const init = await child.request("initialize", {
        protocolVersion: "2025-06-18",
        capabilities: { tools: {}, resources: {} },
        clientInfo: { name: "fuzz-probe-safety", version: "0.0.0" },
      });
      child.notify("notifications/initialized");
      check("safety child initializes", init?.result?.serverInfo?.name === "teams-mcp");
      return child;
    };

    const blockedCalls = [
      ["mail.send", "mail", { op: "send", to: ["a@example.com"], subject: "s", body: "b" }],
      ["mail.reply", "mail", { op: "reply", message_id: "msg-1", body: "b" }],
      ["mail.forward", "mail", { op: "forward", message_id: "msg-1", to: ["a@example.com"] }],
      ["mail.mark", "mail", { op: "mark", message_id: "msg-1", mark: "read" }],
      ["mail.move", "mail", { op: "move", message_id: "msg-1", destination: "archive" }],
      ["mail.unsubscribe_apply", "mail", { op: "unsubscribe_apply", message_id: "msg-1", confirm: true }],
      ["calendar.create", "calendar", { op: "create", subject: "s", start: "2026-11-18T16:00:00Z" }],
      ["calendar.update", "calendar", { op: "update", event_id: "evt-1", expected_subject: "Fuzz Event", subject: "t" }],
      ["calendar.delete", "calendar", { op: "delete", event_id: "evt-1", expected_subject: "Fuzz Event" }],
      ["calendar.respond", "calendar", { op: "respond", event_id: "evt-1", expected_subject: "Fuzz Event", response: "accept" }],
      ["settings.set", "settings", { op: "set", key: "timezone", value: "UTC" }],
      ["auth.logout", "auth", { op: "logout" }],
      // Teams writes: refused before any token read, so no cache/harvest either.
      ["teams.send", "teams", { op: "send", chat_id: "19:probe", text: "x", confirm: true }],
      ["teams.create", "teams", { op: "create", kind: "channel", team_id: "t1", name: "x", private: true, confirm: true }],
      ["teams.delete", "teams", { op: "delete", team_id: "t1", channel_id: "19:probe", expected_name: "x", confirm: true }],
    ];

    const readOnly = await openChild({ TEAMS_MCP_READ_ONLY: "1" });
    try {
      for (const [label, tool, args] of blockedCalls) {
        const before = mock.log.length;
        const call = await callTool(readOnly, tool, args);
        check(`read-only ${label} is refused`, call.isError && call.text.includes("read_only"), call.text.slice(0, 90));
        check(`read-only ${label} issued no request`, mock.log.length === before);
      }
      const beforeRead = mock.log.length;
      const read = await callTool(readOnly, "mail", { op: "list", top: 1 });
      check("read-only mail.list still reads", !read.isError && mock.log.length > beforeRead);
    } finally {
      await readOnly.close();
    }
    const auditText = await readFile(join(base, "cache", "write-audit.jsonl"), "utf8").catch(() => "");
    check("read-only attempts were audited", auditText.includes('"outcome":"read_only"'));

    const optOut = await openChild({ TEAMS_MCP_READ_ONLY: "0", TEAMS_MCP_UNSUBSCRIBE_ENABLED: "0" });
    try {
      const before = mock.log.length;
      const disabled = await callTool(optOut, "mail", {
        op: "unsubscribe_apply",
        message_id: "msg-1",
        confirm: true,
        expected_sender: "sender@example.com",
      });
      check(
        "unsubscribe_apply is disabled by default",
        disabled.isError && disabled.text.includes("unsubscribe_disabled"),
        disabled.text.slice(0, 90),
      );
      check("disabled unsubscribe issued no request", mock.log.length === before);
    } finally {
      await optOut.close();
    }

    const limited = await openChild({
      TEAMS_MCP_READ_ONLY: "0",
      TEAMS_MCP_UNSUBSCRIBE_ENABLED: "0",
      TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE: "1",
    });
    try {
      const beforeFirst = mock.log.length;
      const first = await callTool(limited, "mail", { op: "mark", message_id: "msg-1", mark: "read" });
      check(
        "rate limit: first write passes",
        !first.isError && mock.log.slice(beforeFirst).some((entry) => writeKind(entry) === "mail.mark"),
      );
      const afterFirst = mock.log.length;
      const second = await callTool(limited, "mail", { op: "mark", message_id: "msg-1", mark: "read" });
      check(
        "rate limit: second write is refused",
        second.isError && second.text.includes("rate_limited"),
        second.text.slice(0, 90),
      );
      check("rate limit: refused write made no request", mock.log.length === afterFirst);
    } finally {
      await limited.close();
    }

    /* Seeded random sequence over all ops */
    const pool = shuffled(
      Array.from({ length: Math.ceil(CALLS / MOCK_OPS.length) }, () => MOCK_OPS).flat(),
    ).slice(0, CALLS);
    for (const [index, entry] of pool.entries()) {
      const args = randomArgs(entry.tool, entry.op);
      const before = mock.log.length;
      const call = await callTool(client, entry.tool, { ...args, op: entry.op });
      assertSafeCall(
        `${entry.tool}.${entry.op} [${index}]`,
        call,
        before,
        `${entry.tool}.${entry.op}`,
        true,
      );
    }

    /* Auth ops against the temp config only: no credentials -> no browser, and
       logout removes the fabricated cache (never the canonical one). */
    const login = await callTool(client, "auth", { op: "login" });
    check("auth.login without credentials errors before a browser launch", login.isError && login.text.includes("credentials_missing"));
    const logout = await callTool(client, "auth", { op: "logout" });
    check("auth.logout removes only the temp token cache", !logout.isError && logout.text.includes(configDir));
    const statusAfter = await callTool(client, "auth", { op: "status" });
    check("auth.status reports unauthenticated after the temp logout", !statusAfter.isError && statusAfter.text.includes('"authenticated":false'));

    check("server still answers ping after the sequence", (await client.request("ping"))?.result !== undefined);
    check("stdout is JSON-RPC only", client.badLines.length === 0, client.badLines);
    check("stderr carries no bearer token", !BEARER.test(client.stderr()));
    check(
      "every mock request stayed on the loopback mock",
      mock.log.every((entry) => entry.url.startsWith("/api/v2.0") || entry.url.startsWith("/v1.0")),
    );
    check(
      "no account-mutating endpoint was outside the classified write set",
      mock.log.every((entry) => {
        const kind = writeKind(entry);
        return kind === undefined || !kind.startsWith("unexpected");
      }),
    );
    console.log(`\nmock requests: ${mock.log.length}; random calls: ${pool.length}; seed ${SEED}`);
  } finally {
    await client.close();
    const closed = Promise.withResolvers();
    mock.server?.close(() => closed.resolve());
    await closed.promise;
    await rm(base, { recursive: true, force: true });
  }
}

/* ------------------------------------------------------------------ */
/* Live read-only phase                                                */
/* ------------------------------------------------------------------ */

async function livePhase() {
  const client = connect({});
  try {
    const init = await client.request("initialize", {
      protocolVersion: "2025-06-18",
      capabilities: { tools: {}, resources: {} },
      clientInfo: { name: "fuzz-probe-live", version: "0.0.0" },
    });
    check("live: initialize answers", init?.result?.serverInfo?.name === "teams-mcp");
    client.notify("notifications/initialized");

    const status = await callTool(client, "auth", { op: "status" });
    let authenticated = false;
    try {
      authenticated = JSON.parse(status.text).authenticated === true;
    } catch {
      authenticated = false;
    }
    if (!authenticated) {
      console.log("live reads skipped: not signed in (the mock phase still ran)");
      return;
    }

    /* Build ids from reads instead of guessing account data. */
    const list = await callTool(client, "mail", { op: "list", top: 3 });
    check("live: mail.list responds", list.frame !== undefined);
    let messageId;
    let threadId;
    try {
      const rows = JSON.parse(list.text).messages;
      if (Array.isArray(rows) && rows[0]) {
        messageId = rows[0].id;
        threadId = rows[0].conversation_id;
      }
    } catch {
      messageId = undefined;
    }

    const calendar = await callTool(client, "calendar", { op: "list", top: 3 });
    check("live: calendar.list responds", calendar.frame !== undefined);
    let eventId;
    try {
      const rows = JSON.parse(calendar.text).events;
      if (Array.isArray(rows) && rows[0]) eventId = rows[0].id;
    } catch {
      eventId = undefined;
    }

    /* mail.list + from: Exchange used to reject $orderby with a From filter. */
    let sender;
    try {
      const rows = JSON.parse(list.text).messages;
      if (Array.isArray(rows) && rows[0]?.from) sender = String(rows[0].from);
    } catch {
      sender = undefined;
    }
    if (sender) {
      const before = mock.log.length;
      const filtered = await callTool(client, "mail", { op: "list", from: sender, top: 5 });
      assertSafeCall("mail.list+from", filtered, before);
      check(
        "live: mail.list with from is accepted",
        !filtered.isError && !filtered.text.includes("InefficientFilter"),
        filtered.text.slice(0, 120),
      );
      try {
        const rows = JSON.parse(filtered.text).messages ?? [];
        const times = rows.map((row) => String(row.received ?? ""));
        const ordered = times.every((time, index) => index === 0 || times[index - 1] >= time);
        check("live: from-page is newest first", ordered, times.slice(0, 3));
      } catch {
        check("live: from-page parsed", false, filtered.text.slice(0, 120));
      }
    }

    /* A nonexistent mailbox must be `unknown`, never a free day. */
    const FAKE = "teams-mcp-no-such-mailbox-9f3a@outlook.com";
    const fakeBefore = mock.log.length;
    const fake = await callTool(client, "availability", { op: "freebusy", attendees: [FAKE] });
    assertSafeCall("availability.freebusy fake", fake, fakeBefore);
    try {
      const out = JSON.parse(fake.text);
      const unknown = Array.isArray(out.unknown) ? out.unknown.map((email) => String(email).toLowerCase()) : [];
      check("live: fake mailbox is unknown", unknown.includes(FAKE.toLowerCase()), out.unknown);
      const viewKeys = Object.keys(out.availability_view ?? {}).map((key) => key.toLowerCase());
      check("live: fake mailbox has no availability view", !viewKeys.includes(FAKE.toLowerCase()));
    } catch {
      check("live: fake-mailbox result parsed", false, fake.text.slice(0, 120));
    }

    const reads = [];
    if (messageId) {
      reads.push(["mail.get", "mail", { op: "get", message_id: messageId }]);
      reads.push(["mail.unsubscribe_preview", "mail", { op: "unsubscribe_preview", message_id: messageId }]);
    }
    if (threadId) reads.push(["mail.thread", "mail", { op: "thread", thread_id: threadId, top: 2 }]);
    if (eventId) {
      reads.push(["calendar.get", "calendar", { op: "get", event_id: eventId }]);
      reads.push([
        "calendar.instances",
        "calendar",
        { op: "instances", event_id: eventId, start: new Date().toISOString().slice(0, 10), end: new Date(Date.now() + 7 * 86_400_000).toISOString().slice(0, 10) },
      ]);
    }
    reads.push(["people.search", "people", { op: "search", query: "a", top: 2 }]);
    reads.push(["availability.freebusy", "availability", { op: "freebusy" }]);
    reads.push(["settings.path", "settings", { op: "path" }]);
    reads.push(["settings.validate", "settings", { op: "validate" }]);

    for (const [label, tool, args] of reads) {
      const before = mock.log.length;
      const call = await callTool(client, tool, args);
      assertSafeCall(label, call, before);
    }
    check("live: stdout is JSON-RPC only", client.badLines.length === 0, client.badLines);
    check("live: stderr carries no bearer token", !BEARER.test(client.stderr()));
    check("live: server still answers ping", (await client.request("ping"))?.result !== undefined);
    console.log(`\nlive read calls: ${reads.length}; no write op is present in the live pool`);
  } finally {
    await client.close();
  }
}

/* ------------------------------------------------------------------ */

try {
  if (LIVE) await livePhase();
  else await mockPhase();
} catch (error) {
  failures += 1;
  console.log(`FAIL  probe crashed — ${error instanceof Error ? error.message : String(error)}`);
}

console.log(`\n${checks - failures}/${checks} checks passed (seed ${SEED}, mode ${LIVE ? "live read-only" : "mock"})`);
if (failures > 0) process.exitCode = 1;
