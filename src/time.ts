/**
 * Timezone-aware formatting helpers.
 *
 * All absolute instants are ISO 8601 UTC strings. "Local" means the configured
 * IANA zone (America/Chicago by default) and is only used for display and for
 * working-hours / blocked-time rules. Stepping through time is done in absolute
 * milliseconds, which keeps DST transitions correct (a 09:00 local meeting is
 * still 09:00 local after a clock change).
 */
import type { Weekday } from "./config.js";

const WEEKDAY_ORDER: Weekday[] = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];

const formatters = new Map<string, Intl.DateTimeFormat>();

function partsFormatter(timeZone: string): Intl.DateTimeFormat {
  let formatter = formatters.get(timeZone);
  if (!formatter) {
    formatter = new Intl.DateTimeFormat("en-US", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
      weekday: "short",
    });
    formatters.set(timeZone, formatter);
  }
  return formatter;
}

export interface LocalParts {
  /** YYYY-MM-DD in the target zone. */
  date: string;
  /** HH:MM in the target zone. */
  time: string;
  /** HH:MM:SS in the target zone. */
  clock: string;
  hour: number;
  minute: number;
  second: number;
  /** Three-letter lowercase weekday. */
  weekday: Weekday;
}

export function localParts(instant: string | number | Date, timeZone: string): LocalParts {
  const date = instant instanceof Date ? instant : new Date(instant);
  const fields: Record<string, string> = {};
  for (const part of partsFormatter(timeZone).formatToParts(date)) {
    if (part.type !== "literal") fields[part.type] = part.value;
  }
  const weekday = (fields.weekday ?? "sun").toLowerCase().slice(0, 3) as Weekday;
  return {
    date: `${fields.year}-${fields.month}-${fields.day}`,
    time: `${fields.hour}:${fields.minute}`,
    clock: `${fields.hour}:${fields.minute}:${fields.second}`,
    hour: Number(fields.hour ?? 0),
    minute: Number(fields.minute ?? 0),
    second: Number(fields.second ?? 0),
    weekday: WEEKDAY_ORDER.includes(weekday) ? weekday : "sun",
  };
}

/**
 * Wall-clock literal for Outlook REST bodies: `2026-11-16T09:00:00` in the given
 * zone, with no offset (the request carries `TimeZone` separately).
 */
export function wallClock(instant: string | number | Date, timeZone: string): string {
  const parts = localParts(instant, timeZone);
  return `${parts.date}T${parts.clock}`;
}

/** Calendar date (`YYYY-MM-DD`) in the given zone. */
export function localDate(instant: string | number | Date, timeZone: string): string {
  return localParts(instant, timeZone).date;
}

/** `2026-11-16 09:15` in the given zone. */
export function formatLocal(instant: string | number | Date, timeZone: string): string {
  const parts = localParts(instant, timeZone);
  return `${parts.date} ${parts.time}`;
}

/** Minutes since local midnight for an `HH:MM` clock string. */
export function minutesOfDay(clock: string): number {
  const [hour, minute] = clock.split(":").map(Number);
  return (hour ?? 0) * 60 + (minute ?? 0);
}

/** True when [start, end) falls inside local working hours for that weekday. */
export function withinWorkingHours(
  startMs: number,
  endMs: number,
  timeZone: string,
  workHours: { days: Weekday[]; start: string; end: string },
): boolean {
  const from = localParts(startMs, timeZone);
  const to = localParts(endMs, timeZone);
  if (!workHours.days.includes(from.weekday)) return false;
  if (from.date !== to.date) return false;
  const open = minutesOfDay(workHours.start);
  const close = minutesOfDay(workHours.end);
  const startMinute = from.hour * 60 + from.minute;
  const endMinute = to.hour * 60 + to.minute;
  return startMinute >= open && endMinute <= close;
}

/** True when [start, end) intersects a blocked window on that weekday. */
export function withinBlockedTime(
  startMs: number,
  endMs: number,
  timeZone: string,
  blocked: Array<{ start: string; end: string; days?: Weekday[] }>,
): boolean {
  const from = localParts(startMs, timeZone);
  const startMinute = from.hour * 60 + from.minute;
  const endMinute = startMinute + Math.round((endMs - startMs) / 60_000);
  return blocked.some((window) => {
    if (window.days && !window.days.includes(from.weekday)) return false;
    const open = minutesOfDay(window.start);
    const close = minutesOfDay(window.end);
    return startMinute < close && endMinute > open;
  });
}

/** Round an instant up to the next `intervalMinutes` boundary. */
export function alignUp(ms: number, intervalMinutes: number): number {
  const step = intervalMinutes * 60_000;
  return Math.ceil(ms / step) * step;
}

/**
 * Offset of `timeZone` at an instant, in minutes (local wall clock minus UTC).
 *
 * Computed from the actual instant rather than from a guessed noon, so the value
 * is the offset *in force at that instant* — the only thing that stays correct
 * across a DST transition.
 */
export function zoneOffsetMinutes(instantMs: number, timeZone: string): number {
  const parts = localParts(instantMs, timeZone);
  const utc = new Date(instantMs);
  let offset = parts.hour * 60 + parts.minute - (utc.getUTCHours() * 60 + utc.getUTCMinutes());
  if (offset > 720) offset -= 1440;
  if (offset < -720) offset += 1440;
  return offset;
}

/** `YYYY-MM-DDTHH:MM:SS` for a local wall-clock input, seconds padded. */
function canonicalWallClock(text: string): string {
  return text.length === 16 ? `${text}:00` : text;
}

/**
 * Resolve a bare local wall clock to a UTC instant.
 *
 * Two offset passes are required: the first uses the offset at the UTC guess,
 * which for a wall clock just after a spring-forward can still be the pre-gap
 * offset (one hour off). Re-probing at the candidate instant converges because
 * offsets are stable except within a transition.
 *
 * Returns `undefined` when the wall clock does not exist in the zone (the DST
 * gap, e.g. 02:30 in America/Chicago on a spring-forward day). Ambiguous wall
 * clocks (fall-back hour) resolve to the first occurrence, i.e. the pre-transition
 * offset.
 */
function resolveWallClock(text: string, timeZone: string): number | undefined {
  const wall = canonicalWallClock(text);
  const guess = Date.parse(`${wall}Z`);
  if (Number.isNaN(guess)) return undefined;
  let candidate = guess - zoneOffsetMinutes(guess, timeZone) * 60_000;
  candidate = guess - zoneOffsetMinutes(candidate, timeZone) * 60_000;
  const at = localParts(candidate, timeZone);
  return `${at.date}T${at.clock}` === wall ? candidate : undefined;
}

/** First instant of local date `date` — midnight, or the first instant that exists. */
function midnightOfDate(date: string, timeZone: string): number {
  const guess = Date.parse(`${date}T00:00:00Z`);
  if (Number.isNaN(guess)) return Number.NaN;
  let candidate = guess - zoneOffsetMinutes(guess, timeZone) * 60_000;
  candidate = guess - zoneOffsetMinutes(candidate, timeZone) * 60_000;
  const at = localParts(candidate, timeZone);
  if (at.date === date && at.clock === "00:00:00") return candidate;

  // A few zones (e.g. America/Santiago) shift *at* midnight, so 00:00 local does
  // not exist on the transition day. Walk the surrounding three hours and take
  // exact midnight when it exists, otherwise the first instant of that date.
  let firstOfDate: number | undefined;
  for (let step = -180; step <= 180; step += 1) {
    const probe = candidate + step * 60_000;
    const parts = localParts(probe, timeZone);
    if (parts.date !== date) continue;
    if (parts.clock === "00:00:00") return probe;
    firstOfDate ??= probe;
  }
  return firstOfDate ?? candidate;
}

/** Start of the local day containing `instant`, expressed as a UTC instant. */
export function startOfLocalDay(instant: string | number | Date, timeZone: string): number {
  return midnightOfDate(localParts(instant, timeZone).date, timeZone);
}

export type InstantParse =
  | { instant: number }
  /** `nonexistent`: a valid wall clock that the zone skips (DST gap). */
  | { error: "unrecognized" | "nonexistent" };

/**
 * True for a real calendar date (`YYYY-MM-DD`). `Date.parse` silently normalises
 * `2026-02-30` to March 2 and returns `NaN` for month 13; neither is a valid
 * caller input, and the `NaN` case used to reach `Intl` as a `RangeError`.
 */
function validCalendarDate(date: string): boolean {
  const utc = Date.parse(`${date}T00:00:00Z`);
  return !Number.isNaN(utc) && new Date(utc).toISOString().slice(0, 10) === date;
}

/**
 * Parse with a reason, so a DST-gap wall clock is reported as such instead of
 * silently sliding an hour (or looking like a typo).
 */
export function parseInstantDetailed(value: string, timeZone: string): InstantParse {
  const text = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    if (!validCalendarDate(text)) return { error: "unrecognized" };
    const instant = localMidnightOfDate(text, timeZone);
    return Number.isNaN(instant) ? { error: "unrecognized" } : { instant };
  }
  if (/[zZ]|[+-]\d{2}:\d{2}$/.test(text)) {
    if (!validCalendarDate(text.slice(0, 10))) return { error: "unrecognized" };
    const parsed = Date.parse(text);
    return Number.isNaN(parsed) ? { error: "unrecognized" } : { instant: parsed };
  }
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/.test(text)) return { error: "unrecognized" };
  if (!validCalendarDate(text.slice(0, 10))) return { error: "unrecognized" };
  if (Number.isNaN(Date.parse(`${canonicalWallClock(text)}Z`))) return { error: "unrecognized" };
  const instant = resolveWallClock(text, timeZone);
  return instant === undefined ? { error: "nonexistent" } : { instant };
}

/**
 * Parse a caller-supplied datetime into a UTC instant.
 * Accepts ISO 8601 with an offset, or a bare local wall clock (`YYYY-MM-DDTHH:MM[:SS]`)
 * which is interpreted in `timeZone`, or a date-only `YYYY-MM-DD` (local midnight).
 */
export function parseInstant(value: string, timeZone: string): number | undefined {
  const parsed = parseInstantDetailed(value, timeZone);
  return "instant" in parsed ? parsed.instant : undefined;
}

/** UTC instant of 00:00 local on `YYYY-MM-DD`. */
export function localMidnightOfDate(date: string, timeZone: string): number {
  return midnightOfDate(date, timeZone);
}
