/**
 * Capability scoping for autonomous use.
 *
 * A live-account test demonstrated the failure mode this module exists to stop:
 * an autonomous agent performed bulk mailbox actions (mass moves and irreversible
 * unsubscribes) because the server offered no capability switch and
 * `unsubscribe_apply` needed only a model-supplied `confirm:true`.
 *
 * Four layers, all enforced in-process before the mutating request:
 *
 *  1. `readOnlyGuard` — every write op returns `read_only` with zero network.
 *  2. `unsubscribe` policy — opt-in (default off), sender-bound, daily-capped
 *     per external action (an https POST and a mailto send each take a slot).
 *  3. `writeRateGuard` — one global token bucket for every mutating op
 *     (mail/calendar/settings/auth), consumed before the first request.
 *  4. `auditWrite` — append-only JSONL journal of each write attempt (0600).
 *
 * None of this is a sandbox: a caller with the file system can edit config.yaml.
 * It is a capability boundary for an agent that only sees the MCP surface.
 */
import { randomUUID } from "node:crypto";
import {
  appendFileSync,
  chmodSync,
  mkdirSync,
  renameSync,
  rmSync,
  statSync,
} from "node:fs";
import { chmod, mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { config } from "./config.js";
import { errorResult } from "./tools/shared.js";
import { localDate, localMidnightOfDate } from "./time.js";

/* ------------------------------------------------------------------ */
/* Mutating op classification                                          */
/* ------------------------------------------------------------------ */

/**
 * Ops that can change the account or the local credential/config state.
 *
 * `auth.login` is intentionally absent: it writes only the token cache and is
 * required for read-only operation (an unauthenticated server cannot read).
 * `settings.reload`/`validate` are reads; `availability.snapshot` writes only a
 * read-derived cache file, which is allowed in read-only mode.
 */
const MUTATING_OPS: Record<string, readonly string[]> = {
  mail: ["send", "draft", "send_draft", "reply", "forward", "mark", "move", "unsubscribe_apply"],
  calendar: ["create", "update", "delete", "respond"],
  settings: ["set"],
  auth: ["logout"],
  teams: ["send", "create", "delete"],
};

export function isMutatingOp(tool: string, op: string): boolean {
  return (MUTATING_OPS[tool] ?? []).includes(op);
}

/**
 * Ids/counts for a blocked attempt's audit line. Deliberately a small allowlist:
 * never a body, subject or token.
 */
function targetSummaryOf(
  args: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  if (!args) return undefined;
  const summary: Record<string, unknown> = {};
  for (const key of ["message_id", "draft_id", "event_id", "thread_id", "destination", "key", "chat_id", "channel_id", "team_id", "name", "expected_name"]) {
    const value = args[key];
    if (typeof value === "string" && value !== "") summary[key] = value;
  }
  if (Array.isArray(args.to)) summary.recipients = args.to.length;
  if (Array.isArray(args.attendees)) summary.attendees = args.attendees.length;
  return Object.keys(summary).length > 0 ? summary : undefined;
}

/**
 * Reject a mutating op when the server is read-only.
 *
 * Called at the top of each tool handler (not only in the registry wrapper) so
 * direct handler calls — the test surface, and any future transport — get the
 * same guarantee: no HTTP request precedes this refusal.
 */
export function readOnlyGuard(
  tool: string,
  op: string,
  args?: Record<string, unknown>,
): CallToolResult | undefined {
  if (!isMutatingOp(tool, op)) return undefined;
  if (!config().readOnly) return undefined;
  auditWrite({ tool, op, target: targetSummaryOf(args), outcome: "read_only" });
  return errorResult(
    "read_only",
    `${tool}.${op} is a write and this server is in read-only mode (safety.read_only / TEAMS_MCP_READ_ONLY); nothing was sent or changed`,
    "an operator must unset TEAMS_MCP_READ_ONLY or set safety.read_only: false; reads, availability and auth.status remain available",
  );
}

/* ------------------------------------------------------------------ */
/* Write audit journal                                                 */
/* ------------------------------------------------------------------ */

const AUDIT_FILE = "write-audit.jsonl";
const MAX_AUDIT_BYTES = 5 * 1024 * 1024;

export interface WriteAuditEntry {
  /** ISO 8601 UTC. */
  ts: string;
  tool: string;
  op: string;
  /** Ids / recipient counts / destination only — never bodies or tokens. */
  target?: string | Record<string, unknown>;
  outcome: string;
}

/** Absolute path of the append-only journal (tests point TEAMS_MCP_CACHE_DIR at a temp dir). */
export function writeAuditPath(): string {
  return join(config().cacheDir, AUDIT_FILE);
}

/**
 * Append one line. Best effort by design: a full or unwritable disk must never
 * change whether a write was attempted, so failures are swallowed.
 *
 * Rotation happens before the append: at 5 MB the current file becomes
 * `write-audit.jsonl.1` (replacing any previous generation), so the journal is
 * bounded at ~10 MB on disk.
 */
export function auditWrite(entry: Omit<WriteAuditEntry, "ts">): void {
  try {
    const path = writeAuditPath();
    mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
    try {
      if (statSync(path).size >= MAX_AUDIT_BYTES) {
        rmSync(`${path}.1`, { force: true });
        renameSync(path, `${path}.1`);
      }
    } catch {
      // Missing file: nothing to rotate.
    }
    const line = JSON.stringify({ ts: new Date().toISOString(), ...entry });
    appendFileSync(path, `${line}\n`, { mode: 0o600 });
    // `mode` above applies only when the file is created; an existing 0644 file
    // (or a permissive umask on a pre-created one) is tightened here.
    chmodSync(path, 0o600);
  } catch {
    // Audit never blocks or fails an operation.
  }
}

/** Stable error code for an audit outcome, without messages (which may carry content). */
export function errorOutcome(error: unknown): string {
  if (error !== null && typeof error === "object" && "code" in error) {
    const code = error.code;
    if (typeof code === "string" && code !== "") return code;
  }
  return "error";
}

/**
 * Run one mutation and journal its outcome. Returns the value unchanged and
 * rethrows failures after logging the code (never the message).
 */
export async function auditedWrite<T>(
  tool: string,
  op: string,
  target: string | Record<string, unknown> | undefined,
  run: () => Promise<T>,
): Promise<T> {
  try {
    const value = await run();
    auditWrite({ tool, op, target, outcome: "ok" });
    return value;
  } catch (error) {
    auditWrite({ tool, op, target, outcome: errorOutcome(error) });
    throw error;
  }
}

/* ------------------------------------------------------------------ */
/* Write rate limit (per process token bucket)                         */
/* ------------------------------------------------------------------ */

export type WriteToken = { ok: true } | { ok: false; retryAfterSeconds: number };

interface Bucket {
  tokens: number;
  rate: number;
  last: number;
}

let bucket: Bucket | undefined;

/** Test seam: forget the bucket so each case starts with a full allowance. */
export function resetWriteLimiter(): void {
  bucket = undefined;
}

function refill(ratePerMinute: number, now: number): Bucket {
  const rate = Math.max(1, Math.trunc(ratePerMinute));
  if (!bucket) {
    bucket = { tokens: rate, rate, last: now };
    return bucket;
  }
  const elapsed = Math.max(0, now - bucket.last);
  const perMs = bucket.rate / 60_000;
  // `Math.min(rate, …)` also applies a lowered limit immediately.
  bucket.tokens = Math.min(rate, bucket.tokens + elapsed * perMs);
  bucket.rate = rate;
  bucket.last = now;
  return bucket;
}

function secondsUntilToken(current: Bucket): number {
  return Math.max(1, Math.ceil((1 - current.tokens) / (current.rate / 60_000) / 1000));
}

/** True when a token is available, without consuming it. */
export function peekWriteToken(ratePerMinute = config().writeRateLimitPerMinute, now = Date.now()): WriteToken {
  const current = refill(ratePerMinute, now);
  return current.tokens >= 1 ? { ok: true } : { ok: false, retryAfterSeconds: secondsUntilToken(current) };
}

/** Consume one token, or report when the next one becomes available. */
export function acquireWriteToken(
  ratePerMinute = config().writeRateLimitPerMinute,
  now = Date.now(),
): WriteToken {
  const current = refill(ratePerMinute, now);
  if (current.tokens < 1) return { ok: false, retryAfterSeconds: secondsUntilToken(current) };
  current.tokens -= 1;
  return { ok: true };
}

/** Give an unused token back (a request refused before its mutating call). */
export function refundWriteToken(ratePerMinute = config().writeRateLimitPerMinute, now = Date.now()): void {
  const current = refill(ratePerMinute, now);
  current.tokens = Math.min(current.rate, current.tokens + 1);
}

/**
 * Rate-limit gate for every mutating op. Callers validate arguments first so an
 * invalid call does not consume a token; the token is consumed before the first
 * request, so a rate-limited call issues no network request at all. A caller
 * that refuses after acquiring (bad recipient, allowlist, a failed guard read)
 * must `refundWriteToken()`.
 */
export function writeRateGuard(
  tool: string,
  op: string,
  target?: string | Record<string, unknown>,
): CallToolResult | undefined {
  const token = acquireWriteToken();
  if (token.ok) return undefined;
  auditWrite({ tool, op, target, outcome: "rate_limited" });
  return errorResult(
    "rate_limited",
    `write rate limit reached (${config().writeRateLimitPerMinute}/minute); nothing was sent`,
    `retry in ${token.retryAfterSeconds}s`,
  );
}

/* ------------------------------------------------------------------ */
/* Unsubscribe daily cap                                               */
/* ------------------------------------------------------------------ */

const QUOTA_FILE = "unsubscribe-state.json";

const storedQuotaSchema = z.object({
  date: z.string(),
  count: z.number().int().nonnegative(),
});

export interface UnsubscribeQuota {
  allowed: boolean;
  used: number;
  limit: number;
  /** ISO 8601 UTC instant when the counter resets (next local midnight). */
  reset_at: string;
}

type StoredQuota = z.output<typeof storedQuotaSchema>;

function nextResetAt(now: number, timezone: string): string {
  const today = localDate(now, timezone);
  const tomorrow = new Date(Date.parse(`${today}T00:00:00Z`) + 86_400_000).toISOString().slice(0, 10);
  return new Date(localMidnightOfDate(tomorrow, timezone)).toISOString();
}

async function readQuota(now: number): Promise<StoredQuota> {
  const timezone = config().timezone;
  const today = localDate(now, timezone);
  try {
    const parsed = storedQuotaSchema.safeParse(
      JSON.parse(await readFile(join(config().cacheDir, QUOTA_FILE), "utf8")) as unknown,
    );
    if (parsed.success && parsed.data.date === today) return parsed.data;
  } catch {
    // Missing or unreadable state: start a fresh counter for today.
  }
  return { date: today, count: 0 };
}

async function writeQuota(state: StoredQuota): Promise<void> {
  const path = join(config().cacheDir, QUOTA_FILE);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const tmp = `${path}.${process.pid}.${randomUUID()}.tmp`;
  await writeFile(tmp, `${JSON.stringify(state)}\n`, { mode: 0o600 });
  await rename(tmp, path);
  await chmod(path, 0o600).catch(() => undefined);
}

/** Serialise read-modify-write of the counter inside this process. */
const quotaQueues = new Map<string, Promise<unknown>>();

async function serialise<T>(path: string, task: () => Promise<T>): Promise<T> {
  const prior = quotaQueues.get(path) ?? Promise.resolve();
  const run = prior.then(task, task);
  const tail = run.then(
    () => undefined,
    () => undefined,
  );
  quotaQueues.set(path, tail);
  try {
    return await run;
  } finally {
    if (quotaQueues.get(path) === tail) quotaQueues.delete(path);
  }
}

/** Read-only quota check; used before any message fetch so a capped call is offline. */
export async function peekUnsubscribeQuota(now = Date.now()): Promise<UnsubscribeQuota> {
  const limit = config().unsubscribe.dailyCap;
  const state = await readQuota(now);
  return {
    allowed: state.count < limit,
    used: state.count,
    limit,
    reset_at: nextResetAt(now, config().timezone),
  };
}

/**
 * Atomically re-check and reserve `count` slots immediately before the external
 * actions. One slot per external action (https POST or mailto send), so a
 * message carrying both targets consumes two and can never push the day past
 * the cap. A refusal here means the remaining allowance is smaller than the
 * requested actions; nothing has been sent yet.
 */
export async function reserveUnsubscribeSlots(
  count = 1,
  now = Date.now(),
): Promise<UnsubscribeQuota> {
  const wanted = Math.max(1, Math.trunc(count));
  const path = join(config().cacheDir, QUOTA_FILE);
  return serialise(path, async () => {
    const limit = config().unsubscribe.dailyCap;
    const state = await readQuota(now);
    const reset_at = nextResetAt(now, config().timezone);
    if (state.count + wanted > limit) {
      return { allowed: false, used: state.count, limit, reset_at };
    }
    const next: StoredQuota = { date: state.date, count: state.count + wanted };
    await writeQuota(next);
    return { allowed: true, used: next.count, limit, reset_at };
  });
}
