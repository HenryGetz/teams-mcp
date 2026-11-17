/**
 * Slot finding — pure computation over snapshot busy data.
 *
 * The server does the search: the model gets a handful of ranked candidates plus
 * a one-line conflict summary per slot, never the full busy arrays.
 *
 * Timezone/DST: candidates advance in absolute milliseconds on an interval grid,
 * so a 09:00 local slot stays 09:00 local across DST changes. Working-hour and
 * blocked-time rules compare *local wall clock* values derived with Intl.
 * getSchedule returns times in the timezone requested by the caller (we request
 * the configured IANA zone and normalise to UTC ISO internally).
 */
import type { BlockedTime, WorkHours } from "../config.js";
import type { BusyBlock, BusyStatus } from "../snapshot/calendar.js";
import {
  alignUp,
  formatLocal,
  localDate,
  localParts,
  minutesOfDay,
  parseInstant,
  withinBlockedTime,
  withinWorkingHours,
} from "../time.js";

export interface SuggestOptions {
  /** Busy blocks keyed by attendee email (case-insensitive match). */
  busy: Record<string, BusyBlock[]>;
  attendees: string[];
  start: string;
  end: string;
  durationMin: number;
  intervalMin: number;
  /** Minutes of padding required on both sides of a candidate. */
  bufferMin: number;
  workingHoursOnly: boolean;
  timezone: string;
  workHours: WorkHours;
  blockedTimes: BlockedTime[];
  maxCandidates: number;
  /** Ignore candidates that start before this instant. */
  now?: number;
}

export interface SlotConflict {
  attendee: string;
  reason: "busy" | "tentative" | "oof" | "working_elsewhere" | "working_hours" | "blocked" | "past";
}

export interface SlotCandidate {
  /** ISO 8601 UTC. */
  start: string;
  end: string;
  /** `2026-11-17 10:00` in the configured zone, for display. */
  local_start: string;
  local_end: string;
  conflicts: SlotConflict[];
  free_count: number;
  attendee_count: number;
}

/** Merge overlapping/adjacent blocks regardless of status (for conflict checks). */
function mergeAll(blocks: BusyBlock[]): BusyBlock[] {
  const sorted = [...blocks].sort((a, b) => a.start.localeCompare(b.start));
  const merged: BusyBlock[] = [];
  for (const block of sorted) {
    const last = merged[merged.length - 1];
    if (last && Date.parse(block.start) <= Date.parse(last.end)) {
      if (Date.parse(block.end) > Date.parse(last.end)) last.end = block.end;
      continue;
    }
    merged.push({ ...block });
  }
  return merged;
}

function statusAt(blocks: BusyBlock[], startMs: number, endMs: number): BusyStatus | undefined {
  let worst: BusyStatus | undefined;
  const weight: Record<BusyStatus, number> = {
    free: 0,
    tentative: 1,
    working_elsewhere: 2,
    busy: 3,
    oof: 4,
  };
  for (const block of blocks) {
    const blockStart = Date.parse(block.start);
    const blockEnd = Date.parse(block.end);
    if (blockStart < endMs && blockEnd > startMs) {
      if (!worst || weight[block.status] > weight[worst]) worst = block.status;
    }
  }
  return worst;
}

/** Case-insensitive lookup of an attendee's busy blocks. */
export function busyFor(busy: Record<string, BusyBlock[]>, attendee: string): BusyBlock[] {
  const direct = busy[attendee];
  if (direct) return direct;
  const key = Object.keys(busy).find((candidate) => candidate.toLowerCase() === attendee.toLowerCase());
  return key ? (busy[key] ?? []) : [];
}

/**
 * Worst non-free status overlapping [startMs, endMs), or undefined when the
 * interval is clear. `free` blocks are ignored: they are not evidence of busy.
 */
export function busyStatusIn(blocks: BusyBlock[], startMs: number, endMs: number): BusyStatus | undefined {
  const occupying = blocks.filter(
    (block) =>
      block.status !== "free" &&
      Date.parse(block.start) < endMs &&
      Date.parse(block.end) > startMs,
  );
  return statusAt(occupying, startMs, endMs);
}

/** Non-free blocks that overlap [startMs, endMs), sorted earliest first. */
export function overlappingBusy(blocks: BusyBlock[], startMs: number, endMs: number): BusyBlock[] {
  return blocks
    .filter(
      (block) =>
        block.status !== "free" &&
        Date.parse(block.start) < endMs &&
        Date.parse(block.end) > startMs,
    )
    .sort((a, b) => a.start.localeCompare(b.start));
}

/**
 * Rank candidate slots. Ordering: fewest conflicts, then earliest. Candidates
 * that would overlap an already-returned candidate are dropped so the list is a
 * menu of distinct options.
 */
export function suggestSlots(options: SuggestOptions): SlotCandidate[] {
  const startMs = Date.parse(options.start);
  const endMs = Date.parse(options.end);
  if (Number.isNaN(startMs) || Number.isNaN(endMs) || endMs <= startMs) return [];

  const durationMs = Math.max(1, options.durationMin) * 60_000;
  const bufferMs = Math.max(0, options.bufferMin) * 60_000;
  const intervalMs = Math.max(1, options.intervalMin) * 60_000;
  const nowMs = options.now ?? Date.now();
  const attendees = options.attendees.length > 0 ? options.attendees : [""];
  const merged = new Map<string, BusyBlock[]>();
  for (const attendee of attendees) merged.set(attendee, mergeAll(busyFor(options.busy, attendee)));

  const candidates: SlotCandidate[] = [];
  let cursor = alignUp(Math.max(startMs, options.workingHoursOnly ? startMs : startMs), options.intervalMin);

  for (; cursor + durationMs <= endMs; cursor += intervalMs) {
    if (cursor < nowMs) continue;
    const slotEnd = cursor + durationMs;
    const conflicts: SlotConflict[] = [];

    if (options.workingHoursOnly) {
      if (!withinWorkingHours(cursor, slotEnd, options.timezone, options.workHours)) {
        conflicts.push({ attendee: "(work_hours)", reason: "working_hours" });
      }
    }
    if (withinBlockedTime(cursor, slotEnd, options.timezone, options.blockedTimes)) {
      conflicts.push({ attendee: "(blocked)", reason: "blocked" });
    }
    for (const attendee of attendees) {
      const blocks = merged.get(attendee) ?? [];
      const status = statusAt(blocks, cursor - bufferMs, slotEnd + bufferMs);
      if (status && status !== "free") {
        conflicts.push({ attendee, reason: statusTag(status) });
      }
    }

    const freeCount = attendees.length - conflicts.filter((c) => !c.attendee.startsWith("(")).length;
    candidates.push({
      start: new Date(cursor).toISOString(),
      end: new Date(slotEnd).toISOString(),
      local_start: formatLocal(cursor, options.timezone),
      local_end: formatLocal(slotEnd, options.timezone),
      conflicts,
      free_count: Math.max(0, freeCount),
      attendee_count: attendees.length,
    });
  }

  candidates.sort((a, b) => {
    if (a.conflicts.length !== b.conflicts.length) return a.conflicts.length - b.conflicts.length;
    return Date.parse(a.start) - Date.parse(b.start);
  });

  const chosen: SlotCandidate[] = [];
  const limit = Math.max(1, options.maxCandidates);
  for (const candidate of candidates) {
    if (chosen.length >= limit) break;
    const overlaps = chosen.some(
      (existing) => Date.parse(existing.end) > Date.parse(candidate.start) && Date.parse(existing.start) < Date.parse(candidate.end),
    );
    if (!overlaps) chosen.push(candidate);
  }
  return chosen;
}

function statusTag(status: BusyStatus): SlotConflict["reason"] {
  switch (status) {
    case "tentative":
      return "tentative";
    case "oof":
      return "oof";
    case "working_elsewhere":
      return "working_elsewhere";
    default:
      return "busy";
  }
}

export interface AvailabilityViewOptions {
  start: string;
  end: string;
  /** Granularity of the digit string, Graph default 30 minutes. */
  intervalMinutes?: number;
}

const DIGIT_BY_STATUS: Record<BusyStatus, string> = {
  free: "0",
  tentative: "1",
  busy: "2",
  oof: "3",
  working_elsewhere: "4",
};

/**
 * Compact per-attendee availability view: one digit per interval
 * (0 free, 1 tentative, 2 busy, 3 out-of-office, 4 working elsewhere).
 */
export function buildAvailabilityView(
  busy: Record<string, BusyBlock[]>,
  attendees: string[],
  options: AvailabilityViewOptions,
): Record<string, string> {
  const startMs = Date.parse(options.start);
  const endMs = Date.parse(options.end);
  const intervalMs = Math.max(1, options.intervalMinutes ?? 30) * 60_000;
  if (Number.isNaN(startMs) || Number.isNaN(endMs) || endMs <= startMs) return {};

  const slots = Math.floor((endMs - startMs) / intervalMs);
  const out: Record<string, string> = {};
  for (const attendee of attendees) {
    const blocks = busyFor(busy, attendee);
    let digits = "";
    for (let index = 0; index < slots; index += 1) {
      const slotStart = startMs + index * intervalMs;
      const status = statusAt(blocks, slotStart, slotStart + intervalMs) ?? "free";
      digits += DIGIT_BY_STATUS[status];
    }
    out[attendee] = digits;
  }
  return out;
}

/** Compact busy summary: hour-resolution ranges in local time, max `limit` blocks. */
export type BusyRow = { start: string; end: string; status: BusyStatus } | { omitted: number };

export function summarizeBusyBlocks(
  blocks: BusyBlock[],
  timezone: string,
  limit = 20,
): BusyRow[] {
  // A `busy` list must not contain free intervals — old snapshots may still hold
  // them, so they are dropped on read too.
  const merged = mergeAll(blocks.filter((block) => block.status !== "free"));
  const rows: BusyRow[] = merged.slice(0, limit).map((block) => ({
    start: formatLocal(block.start, timezone),
    end: formatLocal(block.end, timezone),
    status: block.status,
  }));
  if (merged.length > limit) rows.push({ omitted: merged.length - limit });
  return rows;
}

export interface FreeWindow {
  /** ISO 8601 UTC. */
  start: string;
  end: string;
  local_start: string;
  local_end: string;
}

export interface FreeWindowOptions {
  /** Busy blocks keyed by attendee email (case-insensitive match). */
  busy: Record<string, BusyBlock[]>;
  attendees: string[];
  start: string;
  end: string;
  durationMin: number;
  /** Padding required between a free window and any meeting. */
  bufferMin: number;
  workingHoursOnly: boolean;
  timezone: string;
  workHours: WorkHours;
  blockedTimes: BlockedTime[];
  maxSlots: number;
}

export interface FreeWindowResult {
  windows: FreeWindow[];
  /** Matches before the cap. */
  total: number;
  truncated: boolean;
}

/** Local calendar dates touched by [startMs, endMs], inclusive. */
function localDatesBetween(startMs: number, endMs: number, timezone: string): string[] {
  const last = localDate(endMs, timezone);
  const dates: string[] = [];
  let date = localDate(startMs, timezone);
  for (let guard = 0; guard < 400; guard += 1) {
    dates.push(date);
    if (date >= last) break;
    date = new Date(Date.parse(`${date}T00:00:00Z`) + 86_400_000).toISOString().slice(0, 10);
  }
  return dates;
}

interface Interval {
  start: number;
  end: number;
}

/** Allowed search windows: working hours per day, or the raw range. */
function allowedWindows(options: FreeWindowOptions, startMs: number, endMs: number): Interval[] {
  if (!options.workingHoursOnly) return [{ start: startMs, end: endMs }];
  const windows: Interval[] = [];
  for (const date of localDatesBetween(startMs, endMs, options.timezone)) {
    const from = parseInstant(`${date}T${options.workHours.start}`, options.timezone);
    const to = parseInstant(`${date}T${options.workHours.end}`, options.timezone);
    if (from === undefined || to === undefined || to <= from) continue;
    if (!options.workHours.days.includes(localParts(from, options.timezone).weekday)) continue;
    const start = Math.max(from, startMs);
    const end = Math.min(to, endMs);
    if (end > start) windows.push({ start, end });
  }
  return windows;
}

/** Recurring blocked-time windows materialised inside [startMs, endMs). */
function blockedIntervals(options: FreeWindowOptions, startMs: number, endMs: number): Interval[] {
  if (options.blockedTimes.length === 0) return [];
  const out: Interval[] = [];
  for (const date of localDatesBetween(startMs, endMs, options.timezone)) {
    for (const window of options.blockedTimes) {
      const from = parseInstant(`${date}T${window.start}`, options.timezone);
      const to = parseInstant(`${date}T${window.end}`, options.timezone);
      if (from === undefined || to === undefined || to <= from) continue;
      if (window.days && !window.days.includes(localParts(from, options.timezone).weekday)) continue;
      const start = Math.max(from, startMs);
      const end = Math.min(to, endMs);
      if (end > start) out.push({ start, end });
    }
  }
  return out;
}

/**
 * Free windows of at least `durationMin` inside [start, end), computed purely
 * from snapshot busy blocks. `working_hours_only` limits the search to the
 * configured work hours per local weekday; `buffer_min` widens every busy
 * interval so a returned window never sits flush against a meeting.
 */
export function freeWindows(options: FreeWindowOptions): FreeWindowResult {
  const rangeStart = Date.parse(options.start);
  const rangeEnd = Date.parse(options.end);
  if (Number.isNaN(rangeStart) || Number.isNaN(rangeEnd) || rangeEnd <= rangeStart) {
    return { windows: [], total: 0, truncated: false };
  }
  const durationMs = Math.max(1, options.durationMin) * 60_000;
  const bufferMs = Math.max(0, options.bufferMin) * 60_000;

  const busy: Interval[] = [];
  for (const attendee of options.attendees) {
    for (const block of busyFor(options.busy, attendee)) {
      if (block.status === "free") continue;
      const start = Date.parse(block.start);
      const end = Date.parse(block.end);
      if (Number.isNaN(start) || Number.isNaN(end) || end <= start) continue;
      busy.push({ start: Math.max(start, rangeStart), end: Math.min(end, rangeEnd) });
    }
  }
  busy.push(...blockedIntervals(options, rangeStart, rangeEnd));
  const merged = mergeAll(
    busy
      .filter((interval) => interval.end > interval.start)
      .map((interval) => ({
        start: new Date(interval.start).toISOString(),
        end: new Date(interval.end).toISOString(),
        status: "busy" as const,
      })),
  ).map((block) => ({
    start: Date.parse(block.start) - bufferMs,
    end: Date.parse(block.end) + bufferMs,
  }));

  const free: FreeWindow[] = [];
  for (const allowed of allowedWindows(options, rangeStart, rangeEnd)) {
    let cursor = allowed.start;
    for (const block of merged) {
      if (block.end <= cursor) continue;
      if (block.start >= allowed.end) break;
      if (block.start > cursor) {
        const end = Math.min(block.start, allowed.end);
        if (end - cursor >= durationMs) free.push(windowOf(cursor, end, options.timezone));
      }
      cursor = Math.max(cursor, block.end);
      if (cursor >= allowed.end) break;
    }
    if (cursor < allowed.end && allowed.end - cursor >= durationMs) {
      free.push(windowOf(cursor, allowed.end, options.timezone));
    }
  }

  const limit = Math.max(1, options.maxSlots);
  return {
    windows: free.slice(0, limit),
    total: free.length,
    truncated: free.length > limit,
  };
}

function windowOf(startMs: number, endMs: number, timezone: string): FreeWindow {
  return {
    start: new Date(startMs).toISOString(),
    end: new Date(endMs).toISOString(),
    local_start: formatLocal(startMs, timezone),
    local_end: formatLocal(endMs, timezone),
  };
}

/** Minutes remaining in the working day, used to reject impossible durations. */
export function workingDayCapacity(workHours: WorkHours): number {
  return Math.max(0, minutesOfDay(workHours.end) - minutesOfDay(workHours.start));
}

/** Local weekday of an instant, for diagnostics. */
export function localWeekday(instant: string, timezone: string): string {
  return localParts(instant, timezone).weekday;
}
