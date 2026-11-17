/**
 * Protocol-level fuzzing over real stdio: malformed JSON-RPC, unknown methods,
 * batches, notifications, cancellation, oversized/nested params, and several
 * concurrent clients.
 *
 * The child runs with an isolated config/cache dir, a fabricated token, and API
 * base URLs pointed at a closed loopback port, so it cannot reach the live
 * account. Every byte on stdout must be a JSON-RPC frame.
 */
import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { z } from "zod";

const ROOT = new URL("..", import.meta.url).pathname;

const FrameSchema = z
  .object({
    jsonrpc: z.literal("2.0"),
    id: z.union([z.number(), z.string(), z.null()]).optional(),
    result: z.unknown().optional(),
    error: z.unknown().optional(),
    method: z.string().optional(),
  })
  .passthrough();
type Frame = z.infer<typeof FrameSchema>;

const ErrorSchema = z.object({ code: z.number(), message: z.string() }).passthrough();

const ToolResultSchema = z
  .object({
    content: z.array(z.object({ type: z.string(), text: z.string().optional() }).passthrough()).optional(),
    isError: z.boolean().optional(),
  })
  .passthrough();
type ToolResult = z.infer<typeof ToolResultSchema>;

interface Harness {
  child: ChildProcessWithoutNullStreams;
  frames: Frame[];
  badLines: string[];
  stderr: () => string;
  request: (method: string, params?: unknown, timeoutMs?: number) => Promise<Frame | undefined>;
  notify: (method: string, params?: unknown) => void;
  raw: (text: string) => void;
  close: () => Promise<void>;
}

function toolResultOf(frame: Frame | undefined): ToolResult | undefined {
  const parsed = ToolResultSchema.safeParse(frame?.result);
  return parsed.success ? parsed.data : undefined;
}

function textOfResult(frame: Frame | undefined): string {
  const first = toolResultOf(frame)?.content?.[0];
  return first?.type === "text" && typeof first.text === "string" ? first.text : "";
}

function connect(env: Record<string, string>): Harness {
  const child = spawn(process.execPath, ["--import", "tsx", "src/index.ts"], {
    cwd: ROOT,
    env: { ...process.env, ...env },
    stdio: ["pipe", "pipe", "pipe"],
  });
  const frames: Frame[] = [];
  const badLines: string[] = [];
  const pending = new Map<number, (frame: Frame) => void>();
  let buffered = "";
  let stderr = "";
  let nextId = 1;

  child.stdout.setEncoding("utf8");
  child.stdout.on("data", (chunk: string) => {
    buffered += chunk;
    for (;;) {
      const newline = buffered.indexOf("\n");
      if (newline < 0) break;
      const line = buffered.slice(0, newline).replace(/\r$/, "");
      buffered = buffered.slice(newline + 1);
      if (line.trim() === "") continue;
      let decoded: unknown;
      try {
        decoded = JSON.parse(line);
      } catch {
        badLines.push(line.slice(0, 160));
        continue;
      }
      const parsed = FrameSchema.safeParse(decoded);
      if (!parsed.success) {
        badLines.push(line.slice(0, 160));
        continue;
      }
      const frame = parsed.data;
      frames.push(frame);
      const id = frame.id;
      if (typeof id === "number" && pending.has(id)) {
        const resolve = pending.get(id);
        pending.delete(id);
        resolve?.(frame);
      }
    }
  });
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk: string) => {
    stderr += chunk;
  });

  const send = (payload: unknown): void => {
    child.stdin.write(`${JSON.stringify(payload)}\n`);
  };
  const request: Harness["request"] = (method, params, timeoutMs = 20_000) => {
    const id = nextId++;
    const { promise, resolve } = Promise.withResolvers<Frame | undefined>();
    // Integration boundary: the child process is outside fake timers, so a real
    // guard timer is the only way to fail fast instead of hanging the suite.
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
  const notify: Harness["notify"] = (method, params) => {
    send({ jsonrpc: "2.0", method, ...(params === undefined ? {} : { params }) });
  };
  const raw: Harness["raw"] = (text) => child.stdin.write(text);
  const close = async (): Promise<void> => {
    child.stdin.end();
    const { promise, resolve } = Promise.withResolvers<void>();
    // Subprocess exit is a real OS event; keep a hard guard so a wedged child
    // cannot hang the suite.
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

async function handshake(client: Harness): Promise<void> {
  const init = await client.request("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: { tools: {}, resources: {} },
    clientInfo: { name: "protocol-fuzz", version: "0.0.0" },
  });
  expect(init?.result).toBeDefined();
  client.notify("notifications/initialized");
}

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

const TOOL_OPS: ReadonlyArray<{ name: string; ops: readonly string[] }> = [
  { name: "auth", ops: ["status"] },
  {
    name: "mail",
    ops: ["list", "search", "get", "thread", "send", "reply", "forward", "mark", "move", "unsubscribe_preview", "unsubscribe_apply"],
  },
  { name: "calendar", ops: ["list", "get", "create", "update", "delete", "respond", "instances"] },
  { name: "availability", ops: ["freebusy", "suggest", "snapshot"] },
  { name: "people", ops: ["search", "resolve"] },
  { name: "settings", ops: ["get", "validate", "reload", "path"] },
];

const MALFORMED = [
  "{",
  "[",
  "null",
  '"x"',
  '{"jsonrpc":"2.0"}',
  '{"jsonrpc":"2.0","id":',
  "not json at all",
  '{"jsonrpc":"2.0","id":1,"method":"ping","params":',
];

const HOSTILE_VALUES: ReadonlyArray<() => unknown> = [
  () => null,
  () => 0,
  () => -1,
  () => Number.NaN,
  () => "",
  () => "\u0000\r\n💥",
  () => "../../etc/passwd",
  () => "attacker@evil.com, victim@allowed.com",
  () => "2026-13-01",
  () => "Mars/Base",
  () => [],
  () => [null],
  () => Array.from({ length: 1000 }, (_, index) => `u${index}@example.com`),
  () => ({}),
  () => true,
];

function randomArguments(next: () => number): Record<string, unknown> {
  const args: Record<string, unknown> = {};
  const count = 1 + Math.floor(next() * 4);
  for (let index = 0; index < count; index += 1) {
    const key = ["op", "top", "cursor", "timezone", "message_id", "to", "attendees", "query", "key", "value"][
      Math.floor(next() * 10)
    ];
    if (!key) continue;
    const make = HOSTILE_VALUES[Math.floor(next() * HOSTILE_VALUES.length)] ?? (() => null);
    args[key] = make();
  }
  return args;
}

let baseDir: string;
let client: Harness;

beforeAll(async () => {
  baseDir = await mkdtemp(join(tmpdir(), "teams-mcp-protocol-fuzz-"));
  const configDir = join(baseDir, "config");
  await mkdir(configDir, { recursive: true });
  await writeFile(
    join(configDir, "token.json"),
    `${JSON.stringify({
      version: 1,
      account: { upn: "probe@example.com" },
      tokens: {
        outlook: { access_token: "probe-token", expires_at: "2099-01-01T00:00:00.000Z" },
        graph: { access_token: "probe-token", expires_at: "2099-01-01T00:00:00.000Z" },
      },
      source: "headless",
      updated_at: new Date().toISOString(),
    })}\n`,
    { mode: 0o600 },
  );
  client = connect({
    TEAMS_MCP_CONFIG_DIR: configDir,
    TEAMS_MCP_CACHE_DIR: join(baseDir, "cache"),
    TEAMS_MCP_PROFILE_DIR: join(baseDir, "profile"),
    TEAMS_MCP_SNAPSHOT_DIR: join(baseDir, "snapshots"),
    TEAMS_MCP_OUTLOOK_BASE_URL: "http://127.0.0.1:1/api/v2.0",
    TEAMS_MCP_GRAPH_BASE_URL: "http://127.0.0.1:1/v1.0",
    TEAMS_MCP_MAX_ATTEMPTS: "1",
    TEAMS_MCP_REQUEST_TIMEOUT_MS: "2000",
  });
  await handshake(client);
}, 60_000);

afterAll(async () => {
  await client?.close();
  await rm(baseDir, { recursive: true, force: true });
});

describe("protocol abuse", () => {
  it("keeps stdout JSON-RPC-only through malformed lines and unknown methods", async () => {
    for (const line of MALFORMED) client.raw(`${line}\n`);
    client.raw('[{"jsonrpc":"2.0","id":900,"method":"ping"}]\n');
    client.notify("bogus/notification", { anything: true });

    const unknownMethod = await client.request("bogus/method", {});
    expect(ErrorSchema.safeParse(unknownMethod?.error).success).toBe(true);

    const ping = await client.request("ping");
    expect(ping?.result).toBeDefined();
    expect(client.badLines).toEqual([]);
  }, 30_000);

  it("answers wrong-typed tool arguments with isError, never a crash", async () => {
    const cases: Array<{ name: string; args: unknown }> = [
      { name: "mail", args: null },
      { name: "mail", args: {} },
      { name: "mail", args: { op: "bogus" } },
      { name: "mail", args: { op: 5 } },
      { name: "mail", args: { op: "list", top: "many" } },
      { name: "mail", args: { op: "send", to: "a@example.com", subject: "x", body: "y" } },
      { name: "calendar", args: { op: "list", top: -1 } },
      { name: "calendar", args: { op: "create", subject: 42 } },
      { name: "people", args: { op: "search", query: { nested: true } } },
      { name: "settings", args: { op: "set", key: ["aliases", "x"], value: {} } },
      { name: "nope", args: {} },
    ];
    for (const item of cases) {
      const response = await client.request("tools/call", { name: item.name, arguments: item.args });
      expect(response, `${item.name}: no response`).toBeDefined();
      const result = toolResultOf(response);
      if (result) {
        expect(result.isError, `${item.name} ${JSON.stringify(item.args)}`).toBe(true);
        expect(textOfResult(response)).not.toMatch(/\n\s+at\s+[^\n]+:\d+:\d+/);
      } else {
        // Protocol-level validation (e.g. `arguments: null`) may answer with a
        // JSON-RPC error instead of a tool result.
        expect(response?.error).toBeDefined();
      }
    }
    expect(client.badLines).toEqual([]);
  }, 60_000);

  for (const seed of [5, 99, 2026]) {
    it(`survives a seeded burst of 40 protocol frames (seed ${seed})`, async () => {
      const next = random(seed);
      for (let index = 0; index < 40; index += 1) {
        const roll = next();
        if (roll < 0.15) {
          const line = MALFORMED[Math.floor(next() * MALFORMED.length)] ?? "{";
          client.raw(`${line}\n`);
          continue;
        }
        if (roll < 0.25) {
          client.raw(
            `${JSON.stringify([
              { jsonrpc: "2.0", id: 1000 + index, method: "ping" },
              { jsonrpc: "2.0", id: 2000 + index, method: "bogus/method" },
            ])}\n`,
          );
          continue;
        }
        if (roll < 0.35) {
          client.notify(index % 2 === 0 ? "notifications/cancelled" : "notifications/initialized", {
            requestId: 99_999 + index,
          });
          continue;
        }
        const tool = TOOL_OPS[Math.floor(next() * TOOL_OPS.length)] ?? TOOL_OPS[0];
        const op = tool?.ops[Math.floor(next() * (tool?.ops.length ?? 1))] ?? "list";
        const response = await client.request("tools/call", {
          name: tool?.name,
          arguments: { ...randomArguments(next), op },
        });
        expect(response, `seed ${seed} frame ${index}: no response`).toBeDefined();
        const text = textOfResult(response);
        expect(text).not.toMatch(/\n\s+at\s+[^\n]+:\d+:\d+/);
        expect(text).not.toMatch(/Bearer\s+[A-Za-z0-9._~+/=-]{10,}/);
      }
      const ping = await client.request("ping");
      expect(ping?.result).toBeDefined();
      expect(client.badLines).toEqual([]);
    }, 90_000);
  }

  it("accepts oversized and deeply nested params without dying", async () => {
    const deep = (() => {
      let value: Record<string, unknown> = { leaf: true };
      for (let index = 0; index < 300; index += 1) value = { child: value };
      return value;
    })();
    const cases: unknown[] = [
      { name: "people", arguments: { op: "search", query: "x".repeat(1_000_000) } },
      { name: "mail", arguments: { op: "send", to: Array.from({ length: 1000 }, () => "a@example.com"), subject: "x", body: "y", extra: deep } },
      { name: "calendar", arguments: { op: "create", subject: "x", start: "2026-11-18T16:00:00Z", extra: deep } },
    ];
    for (const params of cases) {
      const response = await client.request("tools/call", params, 20_000);
      expect(response).toBeDefined();
      expect(response?.result ?? response?.error).toBeDefined();
    }
    const ping = await client.request("ping");
    expect(ping?.result).toBeDefined();
    expect(client.badLines).toEqual([]);
  }, 60_000);

  it("ignores a cancellation for an unknown request and stays responsive", async () => {
    client.notify("notifications/cancelled", { requestId: 4_242_424, reason: "fuzz" });
    const ping = await client.request("ping");
    expect(ping?.result).toBeDefined();
    expect(client.badLines).toEqual([]);
  });

  it("serves two concurrent clients independently", async () => {
    const second = connect({
      TEAMS_MCP_CONFIG_DIR: join(baseDir, "config"),
      TEAMS_MCP_CACHE_DIR: join(baseDir, "cache-2"),
      TEAMS_MCP_PROFILE_DIR: join(baseDir, "profile-2"),
      TEAMS_MCP_OUTLOOK_BASE_URL: "http://127.0.0.1:1/api/v2.0",
      TEAMS_MCP_GRAPH_BASE_URL: "http://127.0.0.1:1/v1.0",
      TEAMS_MCP_MAX_ATTEMPTS: "1",
      TEAMS_MCP_REQUEST_TIMEOUT_MS: "2000",
    });
    try {
      await handshake(second);
      const [firstPing, secondPing] = await Promise.all([client.request("ping"), second.request("ping")]);
      expect(firstPing?.result).toBeDefined();
      expect(secondPing?.result).toBeDefined();
      expect(second.badLines).toEqual([]);
    } finally {
      await second.close();
    }
  }, 60_000);

  it("never leaks a bearer token to stdout or stderr", () => {
    const stdout = JSON.stringify(client.frames);
    expect(stdout).not.toMatch(/Bearer\s+[A-Za-z0-9._~+/=-]{10,}/);
    expect(client.stderr()).not.toMatch(/Bearer\s+[A-Za-z0-9._~+/=-]{10,}/);
  });
});
