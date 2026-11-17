/**
 * Calendar snapshot layer.
 *
 * Why: a scheduling question should not cost N Graph round-trips per call.
 * One refresh fetches the signed-in user's events (`/me/calendarView`) plus every
 * attendee's busy blocks (`/me/calendar/getSchedule`) for a range, then writes a
 * compact YAML document:
 *
 *   ${TEAMS_MCP_CACHE_DIR:-~/.cache/teams-mcp}/snapshots/calendar.yaml   (0600)
 *
 * Structure:
 *   version: 2
 *   metadata: {fetched_at, range_start, range_end, timezone, ttl_seconds, stale_after, source, attendees[], unknown[]}
 *   events:   [{id, subject, start, end, all_day, show_as, location, is_online}]
 *   busy:     { "alice@example.com": [{start, end, status}] }   # never "free"
 *
 * `metadata.unknown` is REQUIRED from version 2 on. A version-1 snapshot (or any
 * snapshot without the field) is rejected on read and rebuilt, because otherwise
 * an unrecorded unresolved mailbox reads as "free all day".
 *
 * Freshness: the user's own calendar is short-lived (own_calendar_ttl_minutes,
 * default 5) while other people's busy data may be older (busy_ttl_minutes,
 * default 30) — busy data is kept if the snapshot is younger than that longer
 * TTL and still covers the requested range.
 *
 * Every write is atomic (tmp file + rename) and owner-only. Reads never invent
 * data: a missing/corrupt snapshot is reported as unavailable.
 */
import { randomUUID } from "node:crypto";
import { chmod, mkdir, readdir, readFile, rename, stat, unlink, writeFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";
import { parse as parseYaml, stringify as stringifyYaml } from "yaml";
import { z } from "zod";
import { config } from "../config.js";
import { isoTime } from "../format.js";
import { formatLocal, localDate, localParts, parseInstant, startOfLocalDay } from "../time.js";

export const SNAPSHOT_VERSION = 2 as const;

export const BUSY_STATUSES = ["free", "tentative", "busy", "oof", "working_elsewhere"] as const;
export type BusyStatus = (typeof BUSY_STATUSES)[number];

export interface BusyBlock {
  /** ISO 8601 UTC. */
  start: string;
  /** ISO 8601 UTC. */
  end: string;
  status: BusyStatus;
}

export interface SnapshotEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  all_day?: boolean;
  show_as?: BusyStatus;
  location?: string;
  is_online?: boolean;
}

export interface SnapshotMetadata {
  fetched_at: string;
  range_start: string;
  range_end: string;
  timezone: string;
  ttl_seconds: number;
  /** fetched_at + ttl_seconds. */
  stale_after: string;
  source: "outlook" | "shared";
  attendees: string[];
  /** Mailboxes the schedule source could not resolve; `busy: []` here is unknown. */
  unknown?: string[];
  /**
   * Mailboxes whose schedule response actually carried data (working hours,
   * schedule items or an availability view). Absent on older snapshots; readers
   * fall back to "not unknown".
   */
  data_available?: string[];
}

export interface CalendarSnapshot {
  version: typeof SNAPSHOT_VERSION;
  metadata: SnapshotMetadata;
  events: SnapshotEvent[];
  busy: Record<string, BusyBlock[]>;
}

const busyBlockSchema = z.object({
  start: z.string(),
  end: z.string(),
  status: z.enum(BUSY_STATUSES),
});

const snapshotSchema = z.object({
  version: z.literal(SNAPSHOT_VERSION),
  metadata: z.object({
    fetched_at: z.string(),
    range_start: z.string(),
    range_end: z.string(),
    timezone: z.string(),
    ttl_seconds: z.number().int().nonnegative(),
    stale_after: z.string(),
    source: z.enum(["outlook", "shared"]),
    attendees: z.array(z.string()),
    /**
     * Attendees whose mailbox the service could not resolve. Required from
     * version 2: a legacy snapshot without it is invalid and gets rebuilt.
     */
    unknown: z.array(z.string()),
    /** Additive: absent on version-2 snapshots written before it existed. */
    data_available: z.array(z.string()).optional(),
  }),
  events: z.array(
    z.object({
      id: z.string(),
      subject: z.string(),
      start: z.string(),
      end: z.string(),
      all_day: z.boolean().optional(),
      show_as: z.enum(BUSY_STATUSES).optional(),
      location: z.string().optional(),
      is_online: z.boolean().optional(),
    }),
  ),
  busy: z.record(z.string(), z.array(busyBlockSchema)),
});

export class SnapshotError extends Error {
  readonly code: string;
  constructor(code: string, message: string, readonly hint?: string) {
    super(message);
    this.name = "SnapshotError";
    this.code = code;
  }
}

/* ------------------------------------------------------------------ */
/* Graph-facing data source                                            */
/* ------------------------------------------------------------------ */

export interface Range {
  start: string;
  end: string;
}

export interface ScheduleAttendee {
  email: string;
  blocks: BusyBlock[];
  /** Optional `availabilityView` returned by Graph (per-interval digits). */
  availability_view?: string;
  /**
   * The service returned no working hours for this mailbox (typo, foreign
   * mailbox, or no schedule data). An empty result is *not* evidence of a free
   * person, so callers must not render it as free.
   */
  unknown?: true;
  /**
   * The service returned real schedule data for this mailbox (working hours,
   * schedule items or an availability view). Separate from `unknown`: an error
   * row can still carry items.
   */
  data_available?: boolean;
  /** Provider-side per-mailbox error, when one was returned. */
  error?: { code?: string; message?: string };
}

export interface CalendarSource {
  /** The signed-in user's events for the range. */
  ownEvents(range: Range): Promise<SnapshotEvent[]>;
  /** Busy blocks for every mailbox in `attendees` (organizer included). */
  schedules(attendees: string[], range: Range, timezone: string, intervalMinutes: number): Promise<ScheduleAttendee[]>;
}

/* ------------------------------------------------------------------ */
/* Building                                                            */
/* ------------------------------------------------------------------ */

/** Merge overlapping/adjacent blocks of the same status. */
export function mergeBlocks(blocks: BusyBlock[]): BusyBlock[] {
  const sorted = [...blocks].sort((a, b) => a.start.localeCompare(b.start));
  const merged: BusyBlock[] = [];
  for (const block of sorted) {
    const last = merged[merged.length - 1];
    if (last && last.status === block.status && Date.parse(block.start) <= Date.parse(last.end)) {
      if (Date.parse(block.end) > Date.parse(last.end)) last.end = block.end;
      continue;
    }
    merged.push({ ...block });
  }
  return merged;
}

export interface BuildSnapshotOptions {
  source: CalendarSource;
  attendees: string[];
  range: Range;
  timezone?: string;
  /** Longer of the own/busy TTLs is stored; callers apply the granular policy. */
  ttlSeconds?: number;
  intervalMinutes?: number;
  now?: number;
}

export async function buildSnapshot(options: BuildSnapshotOptions): Promise<CalendarSnapshot> {
  const cfg = config();
  const timezone = options.timezone ?? cfg.timezone;
  const ttlSeconds = options.ttlSeconds ?? cfg.cache.busyTtlSeconds;
  const intervalMinutes = options.intervalMinutes ?? 30;
  const attendees = dedupe(options.attendees);

  const [events, schedules] = await Promise.all([
    options.source.ownEvents(options.range),
    attendees.length > 0
      ? options.source.schedules(attendees, options.range, timezone, intervalMinutes)
      : Promise.resolve([] as ScheduleAttendee[]),
  ]);

  const busy: Record<string, BusyBlock[]> = {};
  const unknown: string[] = [];
  const dataAvailable: string[] = [];
  const byEmail = new Map<string, ScheduleAttendee>();
  for (const schedule of schedules) {
    if (schedule.email) byEmail.set(schedule.email.trim().toLowerCase(), schedule);
  }
  // Iterate the requested attendees, not the response: a provider may omit a
  // row entirely for an address it cannot resolve. That absence is unknown, not
  // "free", and must be recorded so `freebusy` can refuse to render it.
  for (const email of attendees) {
    const schedule = byEmail.get(email.trim().toLowerCase());
    if (!schedule) {
      busy[email] = [];
      unknown.push(email);
      continue;
    }
    // `busy` must never claim "free": a free block is the absence of one. Free
    // intervals are still encoded as 0 digits by `availabilityView`.
    busy[schedule.email] = schedule.blocks.filter((block) => block.status !== "free");
    if (schedule.unknown) unknown.push(schedule.email);
    // Older/custom sources do not set the flag: any resolved row with blocks
    // proves data, and a resolved row without `unknown` is treated as answered.
    if (schedule.data_available ?? (schedule.blocks.length > 0 || !schedule.unknown)) {
      dataAvailable.push(schedule.email);
    }
  }

  const fetchedAt = new Date(options.now ?? Date.now()).toISOString();
  return {
    version: SNAPSHOT_VERSION,
    metadata: {
      fetched_at: fetchedAt,
      range_start: isoTime(options.range.start) ?? options.range.start,
      range_end: isoTime(options.range.end) ?? options.range.end,
      timezone,
      ttl_seconds: ttlSeconds,
      stale_after: new Date(Date.parse(fetchedAt) + ttlSeconds * 1000).toISOString(),
      source: "outlook",
      attendees,
      unknown: dedupe(unknown),
      data_available: dedupe(dataAvailable),
    },
    events: events.sort((a, b) => a.start.localeCompare(b.start)),
    busy,
  };
}

function dedupe(values: string[]): string[] {
  const seen: Record<string, true> = {};
  const out: string[] = [];
  for (const value of values) {
    const key = value.trim().toLowerCase();
    if (!key || seen[key]) continue;
    seen[key] = true;
    out.push(value.trim());
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Persistence                                                         */
/* ------------------------------------------------------------------ */

export function snapshotPath(): string {
  return config().snapshotPath;
}

/** Temp files older than this are leftovers from a killed writer. */
const STALE_TEMP_MS = 60 * 60 * 1000;

/**
 * Remove abandoned `<snapshot>.<pid>.<uuid>.tmp` siblings.
 *
 * Temp names are unique per write, so a crashed writer would otherwise leave one
 * file behind forever.
 */
async function cleanStaleTemps(path: string, now = Date.now()): Promise<void> {
  const dir = dirname(path);
  const prefix = `${basename(path)}.`;
  let names: string[];
  try {
    names = await readdir(dir);
  } catch {
    return;
  }
  await Promise.all(
    names
      .filter((name) => name.startsWith(prefix) && name.endsWith(".tmp"))
      .map(async (name) => {
        const full = join(dir, name);
        const info = await stat(full).catch(() => undefined);
        if (info && now - info.mtimeMs > STALE_TEMP_MS) await unlink(full).catch(() => undefined);
      }),
  );
}

/**
 * Atomically replace the snapshot file.
 *
 * The temp name carries pid **and** a UUID: two writers in one process (parallel
 * tool calls) previously shared `<path>.<pid>.tmp`, so the second `rename` lost
 * its source and surfaced `ENOENT … rename`. The target is only ever replaced by
 * an atomic rename, so the previous snapshot survives any failure.
 */
export async function writeSnapshot(snapshot: CalendarSnapshot, path = snapshotPath()): Promise<string> {
  const parsed = snapshotSchema.parse(snapshot);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  await cleanStaleTemps(path);
  const tmp = `${path}.${process.pid}.${randomUUID()}.tmp`;
  const body = [
    "# teams-mcp calendar snapshot — generated, safe to delete",
    `# fetched_at: ${parsed.metadata.fetched_at}  range: ${parsed.metadata.range_start}..${parsed.metadata.range_end}`,
    stringifyYaml(parsed, { lineWidth: 200 }),
  ].join("\n");
  try {
    await writeFile(tmp, body, { mode: 0o600 });
    try {
      await rename(tmp, path);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      // The temp vanished under us (a cleaner or a losing rename): write it once
      // more and retry. A second failure is real and propagates.
      await writeFile(tmp, body, { mode: 0o600 });
      await rename(tmp, path);
    }
  } catch (error) {
    await unlink(tmp).catch(() => undefined);
    throw error;
  }
  await chmod(path, 0o600).catch(() => undefined);
  return path;
}

/** Serialise snapshot refreshes per path inside this process. */
const refreshQueues = new Map<string, Promise<unknown>>();

async function exclusive<T>(path: string, task: () => Promise<T>): Promise<T> {
  const prior = refreshQueues.get(path) ?? Promise.resolve();
  const run = prior.then(task, task);
  const tail = run.then(
    () => undefined,
    () => undefined,
  );
  refreshQueues.set(path, tail);
  try {
    return await run;
  } finally {
    if (refreshQueues.get(path) === tail) refreshQueues.delete(path);
  }
}

export interface LoadedSnapshot {
  snapshot: CalendarSnapshot;
  path: string;
  mtimeMs: number;
}

/** Read and validate a snapshot; returns null when absent or unreadable. */
export async function readSnapshot(path = snapshotPath()): Promise<LoadedSnapshot | null> {
  let text: string;
  try {
    text = await readFile(path, "utf8");
  } catch {
    return null;
  }
  try {
    const parsed = snapshotSchema.parse(parseYaml(text));
    const info = await stat(path).catch(() => undefined);
    return { snapshot: parsed, path, mtimeMs: info?.mtimeMs ?? 0 };
  } catch {
    return null;
  }
}

export async function removeSnapshot(path = snapshotPath()): Promise<boolean> {
  try {
    await unlink(path);
    return true;
  } catch {
    return false;
  }
}

/** True when `fetched_at + ttl_seconds` has passed. */
export function isStale(snapshot: CalendarSnapshot, now = Date.now(), ttlOverrideSeconds?: number): boolean {
  const fetched = Date.parse(snapshot.metadata.fetched_at);
  if (Number.isNaN(fetched)) return true;
  const ttl =
    ttlOverrideSeconds ??
    (snapshot.metadata.source === "outlook"
      ? Math.min(snapshot.metadata.ttl_seconds, config().cache.ownCalendarTtlSeconds)
      : snapshot.metadata.ttl_seconds);
  return fetched + ttl * 1000 <= now;
}

/** True when the snapshot's recorded range covers [start, end]. */
export function rangeCovers(snapshot: CalendarSnapshot, start: string, end: string): boolean {
  const snapStart = Date.parse(snapshot.metadata.range_start);
  const snapEnd = Date.parse(snapshot.metadata.range_end);
  const wantStart = Date.parse(start);
  const wantEnd = Date.parse(end);
  if ([snapStart, snapEnd, wantStart, wantEnd].some((value) => Number.isNaN(value))) return false;
  return snapStart <= wantStart && snapEnd >= wantEnd;
}

/* ------------------------------------------------------------------ */
/* In-memory cache + orchestration                                     */
/* ------------------------------------------------------------------ */

let memory: { path: string; snapshot: CalendarSnapshot; loadedAt: number } | undefined;

export function cacheSnapshot(snapshot: CalendarSnapshot, path = snapshotPath()): void {
  memory = { path, snapshot, loadedAt: Date.now() };
}

export function clearSnapshotCache(): void {
  memory = undefined;
}

export interface EnsureSnapshotOptions {
  /** Required only when a refresh is needed. */
  source?: CalendarSource;
  /** Defaults to the signed-in user. */
  attendees?: string[];
  start: string;
  end: string;
  timezone?: string;
  /** Busy TTL; the shorter own-calendar TTL is applied on top. */
  ttlSeconds?: number;
  /** Force a fetch even when the cached copy is usable. */
  refresh?: boolean;
  now?: number;
}

export interface EnsureSnapshotResult {
  snapshot: CalendarSnapshot;
  path: string;
  /** True when the snapshot was rebuilt during this call. */
  refreshed: boolean;
  /** True when the returned data is past its own-calendar TTL. */
  stale: boolean;
}

/**
 * Return a snapshot covering [start, end], refreshing when:
 *  - there is nothing cached, or
 *  - the cached copy is past its own-calendar TTL, or
 *  - the cached copy does not cover the requested range, or
 *  - `refresh` is set.
 * A refresh without a `source` raises `SnapshotError("snapshot_unavailable")`.
 */
export async function ensureSnapshot(options: EnsureSnapshotOptions): Promise<EnsureSnapshotResult> {
  const cfg = config();
  const path = cfg.snapshotPath;
  const now = options.now ?? Date.now();
  const busyTtl = options.ttlSeconds ?? cfg.cache.busyTtlSeconds;
  const attendees = dedupe(options.attendees ?? []);

  const usable = (snapshot: CalendarSnapshot): boolean =>
    !options.refresh &&
    !isStale(snapshot, now) &&
    // Version 2 requires this array; a legacy object without it cannot tell
    // "free" from "no data" and must be rebuilt.
    Array.isArray(snapshot.metadata.unknown) &&
    rangeCovers(snapshot, options.start, options.end) &&
    attendees.every((email) =>
      Object.keys(snapshot.busy).some((key) => key.toLowerCase() === email.toLowerCase()),
    );

  if (memory && memory.path === path && usable(memory.snapshot)) {
    return { snapshot: memory.snapshot, path, refreshed: false, stale: false };
  }

  const fromDisk = await readSnapshot(path);
  if (fromDisk && usable(fromDisk.snapshot)) {
    cacheSnapshot(fromDisk.snapshot, path);
    return { snapshot: fromDisk.snapshot, path, refreshed: false, stale: false };
  }

  if (!options.source) {
    if (fromDisk) {
      return {
        snapshot: fromDisk.snapshot,
        path,
        refreshed: false,
        stale: isStale(fromDisk.snapshot, now),
      };
    }
    throw new SnapshotError(
      "snapshot_unavailable",
      "no calendar snapshot on disk and no refresh source available",
      "sign in and call availability op=snapshot",
    );
  }

  const source = options.source;
  // One refresh at a time per snapshot path. Parallel availability calls used to
  // build and rename concurrently; the loser hit ENOENT and lost the write.
  const seen = memory?.snapshot;
  return await exclusive(path, async () => {
    // Double-check inside the lock: another call may have refreshed while we
    // waited. A *newer* snapshot that covers this request is reused even when
    // the caller passed refresh=true (one forced refresh, not one per caller).
    const rebuilt = memory && memory.path === path && memory.snapshot !== seen;
    if (rebuilt && memory && !isStale(memory.snapshot, now) && rangeCovers(memory.snapshot, options.start, options.end)) {
      return { snapshot: memory.snapshot, path, refreshed: false, stale: false };
    }
    if (!options.refresh) {
      const disk = await readSnapshot(path);
      if (disk && usable(disk.snapshot)) {
        cacheSnapshot(disk.snapshot, path);
        return { snapshot: disk.snapshot, path, refreshed: false, stale: false };
      }
    }

    const snapshot = await buildSnapshot({
      source,
      attendees,
      range: { start: options.start, end: options.end },
      timezone: options.timezone ?? cfg.timezone,
      ttlSeconds: busyTtl,
      now,
    });
    const written = await writeSnapshot(snapshot, path);
    cacheSnapshot(snapshot, written);
    return { snapshot, path: written, refreshed: true, stale: false };
  });
}

/** Attendees the schedule source could not resolve for this snapshot. */
export function unknownAttendeesOf(snapshot: CalendarSnapshot): string[] {
  return snapshot.metadata.unknown ?? [];
}

/** Case-insensitive membership test for the snapshot's unknown list. */
export function isUnknownAttendee(snapshot: CalendarSnapshot, email: string): boolean {
  const key = email.trim().toLowerCase();
  return unknownAttendeesOf(snapshot).some((value) => value.trim().toLowerCase() === key);
}

/**
 * True when the schedule service returned real data for this mailbox (working
 * hours, schedule items or an availability view). Older snapshots lack the
 * field, so anything not marked unknown is treated as answered.
 */
export function hasScheduleData(snapshot: CalendarSnapshot, email: string): boolean {
  const key = email.trim().toLowerCase();
  const flagged = snapshot.metadata.data_available;
  if (Array.isArray(flagged)) {
    return flagged.some((value) => value.trim().toLowerCase() === key);
  }
  return !isUnknownAttendee(snapshot, email);
}

/** Compact summary for tool output (no full event/busy arrays). */
export function summarizeSnapshot(snapshot: CalendarSnapshot, path: string, now = Date.now()): Record<string, unknown> {
  const parts = localParts(snapshot.metadata.range_start, snapshot.metadata.timezone);
  return {
    path,
    fetched_at: snapshot.metadata.fetched_at,
    stale: isStale(snapshot, now),
    range: {
      start: formatLocal(snapshot.metadata.range_start, snapshot.metadata.timezone),
      end: formatLocal(snapshot.metadata.range_end, snapshot.metadata.timezone),
      timezone: snapshot.metadata.timezone,
      first_day: parts.date,
    },
    ttl_seconds: snapshot.metadata.ttl_seconds,
    events: snapshot.events.length,
    attendees: Object.entries(snapshot.busy).map(([email, blocks]) => ({
      email,
      busy_blocks: blocks.length,
      ...(isUnknownAttendee(snapshot, email) ? { unknown: true } : {}),
    })),
    ...(unknownAttendeesOf(snapshot).length > 0 ? { unknown: unknownAttendeesOf(snapshot) } : {}),
  };
}

/**
 * Default range for scheduling calls: 09:00 today .. 17:00 local seven days out.
 *
 * The bounds are local wall clocks, not "midnight + 9 hours": on a DST
 * transition day the latter lands at 08:00 or 10:00 local.
 */
export function defaultRange(timezone: string, now = Date.now()): Range {
  const startDate = localDate(now, timezone);
  const endDate = localDate(now + 7 * 24 * 60 * 60 * 1000, timezone);
  const start =
    parseInstant(`${startDate}T09:00`, timezone) ?? startOfLocalDay(now, timezone) + 9 * 60 * 60 * 1000;
  const end =
    parseInstant(`${endDate}T17:00`, timezone) ??
    startOfLocalDay(now + 7 * 24 * 60 * 60 * 1000, timezone) + 17 * 60 * 60 * 1000;
  return { start: new Date(start).toISOString(), end: new Date(end).toISOString() };
}
