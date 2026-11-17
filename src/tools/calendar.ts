/**
 * `calendar` tool: Outlook REST v2.0 calendar (`/me/calendarview`, `/me/events`).
 *
 * All ops are live. Writes read the event first and compare the caller's
 * `expected_subject` before touching it; a mismatch aborts before any PATCH,
 * DELETE or RSVP is issued.
 */
import { z } from "zod";
import { config } from "../config.js";
import {
  buildEventPatch,
  createEvent,
  deleteEvent,
  getEvent,
  listCalendarView,
  listInstances,
  normalizeSubject,
  patchEvent,
  respondToEvent,
  type EventResponse,
  type EventSensitivity,
  type EventShowAs,
  type EventUpdate,
  type CalendarQuery,
} from "../outlook/calendar.js";
import { outlookClient, normalizeEvent, readCollection, type CalendarEvent, type OutlookClient } from "../outlook/client.js";
import { defaultRange } from "../snapshot/calendar.js";
import {
  alignUp,
  formatLocal,
  localDate,
  localMidnightOfDate,
  localParts,
  parseInstantDetailed,
} from "../time.js";
import { truncate } from "../format.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import {
  arrayArg,
  booleanArg,
  disallowedRecipients,
  encodeCursor,
  errorResult,
  numberArg,
  opOf,
  pageSize,
  resolveCursor,
  stringArg,
  textResult,
  timezoneOf,
  toolError,
  unknownOp,
  type ToolModule,
} from "./shared.js";
import { resolveRecipients } from "./recipients.js";
import { auditWrite, auditedWrite, readOnlyGuard, refundWriteToken, writeRateGuard } from "../safety.js";

const OPS = ["list", "agenda", "get", "create", "update", "delete", "respond", "instances"] as const;

const SHOW_AS: Record<string, EventShowAs> = {
  free: "Free",
  tentative: "Tentative",
  busy: "Busy",
  oof: "Oof",
  working_elsewhere: "WorkingElsewhere",
};

const SENSITIVITY: Record<string, EventSensitivity> = {
  normal: "Normal",
  personal: "Personal",
  private: "Private",
  confidential: "Confidential",
};

/** Fields `update` can change; checked before the guard read so an empty update is offline. */
const UPDATE_FIELDS = [
  "subject",
  "body",
  "location",
  "attendees",
  "all_day",
  "show_as",
  "sensitivity",
  "start",
  "end",
  "duration_min",
  "shift_min",
] as const;

interface EventRow {
  id: string;
  subject: string;
  start: string;
  end: string;
  local: string;
  all_day?: boolean;
  show_as?: string;
  location?: string;
  organizer?: string;
  attendees?: string[];
  online?: boolean;
  web_link?: string;
}

/**
 * All-day events are returned as UTC midnight with an exclusive end date, so a
 * timezone conversion shows "the previous evening". The date is the payload:
 * label it by date instead.
 */
function allDayLocal(start: string, end: string): string {
  const startDate = start.slice(0, 10);
  const parsed = Date.parse(`${end.slice(0, 10)}T00:00:00Z`);
  const last = Number.isNaN(parsed) ? startDate : new Date(parsed - 86_400_000).toISOString().slice(0, 10);
  return last <= startDate ? `${startDate} (all day)` : `${startDate} → ${last} (all day)`;
}

/**
 * Local rendering of a provider instant. A malformed upstream payload must not
 * turn one row into a `RangeError` for the whole page: show the raw value.
 */
function eventRow(
  event: CalendarEvent,
  timezone: string,
  detailed = false,
  includeWebLink = true,
): EventRow {
  const local = event.allDay
    ? allDayLocal(event.start, event.end)
    : `${Number.isNaN(Date.parse(event.start)) ? event.start : formatLocal(event.start, timezone)} → ${
        Number.isNaN(Date.parse(event.end)) ? event.end : formatLocal(event.end, timezone)
      }`;
  return {
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    local,
    ...(event.allDay === true ? { all_day: true } : {}),
    ...(event.showAs ? { show_as: event.showAs } : {}),
    ...(event.location ? { location: truncate(event.location, 120) } : {}),
    ...(event.organizer?.email ? { organizer: event.organizer.email } : {}),
    ...(event.attendees
      ? {
          attendees: event.attendees.map((attendee) =>
            detailed && attendee.status
              ? `${attendee.email ?? attendee.name ?? "?"} (${attendee.status})`
              : (attendee.email ?? attendee.name ?? "?"),
          ),
        }
      : {}),
    ...(event.isOnlineMeeting !== undefined ? { online: event.isOnlineMeeting } : {}),
    ...(includeWebLink && event.webLink ? { web_link: event.webLink } : {}),
  };
}

/** `brief: true` keeps only the fields needed to pick and reference an event. */
function listEventRow(
  event: CalendarEvent,
  timezone: string,
  brief: boolean,
  includeWebLink: boolean,
): EventRow {
  const row = eventRow(event, timezone, false, includeWebLink);
  if (!brief) return row;
  return {
    id: row.id,
    subject: row.subject,
    start: row.start,
    end: row.end,
    local: row.local,
    ...(row.show_as ? { show_as: row.show_as } : {}),
  };
}

function rangeFrom(args: Record<string, unknown>, timezone: string): { start: string; end: string } {
  const fallback = defaultRange(timezone);
  return {
    start: stringArg(args, "start") ?? fallback.start,
    end: stringArg(args, "end") ?? fallback.end,
  };
}

/**
 * Resolve the caller's window to UTC ISO instants. A bare date means local
 * midnight and a bare wall clock is read in `timezone`, so the API never has to
 * guess a zone (and a DST-gap wall clock is reported rather than slid an hour).
 */
function normalizeRange(
  args: Record<string, unknown>,
  timezone: string,
): { start: string; end: string } | { error: string; hint: string } {
  const raw = rangeFrom(args, timezone);
  const start = parseInstantDetailed(raw.start, timezone);
  if ("error" in start) return instantError("start", raw.start, timezone, start.error);
  const end = parseInstantDetailed(raw.end, timezone);
  if ("error" in end) return instantError("end", raw.end, timezone, end.error);
  if (end.instant <= start.instant) {
    return { error: "end must be after start", hint: "check the requested window" };
  }
  return { start: new Date(start.instant).toISOString(), end: new Date(end.instant).toISOString() };
}

/** Parse one required datetime argument, keeping DST-gap diagnostics distinct. */
function parseField(
  label: string,
  input: string,
  timezone: string,
): { ms: number } | { error: string; hint: string } {
  const parsed = parseInstantDetailed(input, timezone);
  if ("error" in parsed) return instantError(label, input, timezone, parsed.error);
  return { ms: parsed.instant };
}

/** A DST-gap wall clock is a real input error, not a typo: say which one it is. */
function instantError(
  label: string,
  input: string,
  timezone: string,
  reason: "unrecognized" | "nonexistent",
): { error: string; hint: string } {
  if (reason === "nonexistent") {
    return {
      error: `${label} "${input}" does not exist in ${timezone} (DST spring-forward gap)`,
      hint: "choose a wall-clock time that exists, or pass an ISO 8601 timestamp with an offset",
    };
  }
  return { error: `cannot parse ${label} "${input}"`, hint: "use ISO 8601 or YYYY-MM-DD" };
}

/** Returns the actual subject when the guard fails. */
function subjectGuard(event: CalendarEvent, expected: string | undefined): string | undefined {
  if (expected === undefined) return event.subject;
  return normalizeSubject(event.subject) === normalizeSubject(expected) ? undefined : event.subject;
}

const CONFLICT_POLICIES = ["warn", "refuse", "ignore"] as const;
type ConflictPolicy = (typeof CONFLICT_POLICIES)[number];

interface ConflictRow {
  id: string;
  subject: string;
  start: string;
  end: string;
  local: string;
}

interface ConflictReport {
  conflicts: ConflictRow[];
  total: number;
}

/** `conflict_policy` defaults to warn; an unknown value is a caller error. */
function conflictPolicyOf(args: Record<string, unknown>): ConflictPolicy | { error: string } {
  const raw = stringArg(args, "conflict_policy");
  if (raw === undefined) return "warn";
  return (CONFLICT_POLICIES as readonly string[]).includes(raw)
    ? (raw as ConflictPolicy)
    : { error: 'conflict_policy must be "warn", "refuse" or "ignore"' };
}

/**
 * Overlapping events on the user's own calendar for [start, end). `free` events
 * are excluded (they do not occupy the time) and the event being updated is
 * excluded by id. The row list is capped at five; `total` keeps the true count.
 */
async function findConflicts(
  client: OutlookClient,
  timezone: string,
  start: string,
  end: string,
  excludeId?: string,
): Promise<ConflictReport> {
  const rangeStart = Date.parse(start);
  const rangeEnd = Date.parse(end);
  const { events } = await listCalendarView(client, { range: { start, end }, timezone, top: 50 });
  const overlapping = events
    .filter((event) => {
      if (excludeId !== undefined && event.id === excludeId) return false;
      if ((event.showAs ?? "").toLowerCase() === "free") return false;
      const eventStart = Date.parse(event.start);
      const eventEnd = Date.parse(event.end);
      if ([eventStart, eventEnd].some((value) => Number.isNaN(value))) return false;
      return eventStart < rangeEnd && eventEnd > rangeStart;
    })
    .sort((a, b) => a.start.localeCompare(b.start));
  return {
    conflicts: overlapping.slice(0, 5).map((event) => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      local: event.allDay
        ? allDayLocal(event.start, event.end)
        : `${formatLocal(event.start, timezone)} → ${formatLocal(event.end, timezone)}`,
    })),
    total: overlapping.length,
  };
}

/** `refuse`: structured error, no POST/PATCH. */
function conflictRefusal(report: ConflictReport): CallToolResult {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({
          error: "conflict_detected",
          message: `${report.total} event(s) already overlap the proposed time; nothing was written`,
          conflicts: report.conflicts,
          conflict_count: report.total,
          hint: 'pass conflict_policy:"warn" to book anyway, or choose another time',
        }),
      },
    ],
    isError: true,
  };
}

/** Fields a successful create/update adds when a conflict check ran. */
function conflictFields(policy: ConflictPolicy, report: ConflictReport | undefined): Record<string, unknown> {
  if (policy === "ignore") return { conflict_policy: policy };
  if (!report) return { conflict_policy: policy, conflict_check: "unavailable" };
  return {
    conflict_policy: policy,
    conflicts: report.conflicts,
    ...(report.total > 0 ? { conflict_count: report.total, conflict_warning: true } : {}),
  };
}

/** `changed` block: before/after start plus durations, in whole minutes. */
function changedReport(beforeStart: string, beforeEnd: string, afterStart: string, afterEnd: string) {
  const beforeMs = Date.parse(beforeEnd) - Date.parse(beforeStart);
  const afterMs = Date.parse(afterEnd) - Date.parse(afterStart);
  return {
    start_before: beforeStart,
    start_after: afterStart,
    duration_before_min: Number.isFinite(beforeMs) ? Math.round(beforeMs / 60_000) : 0,
    duration_after_min: Number.isFinite(afterMs) ? Math.round(afterMs / 60_000) : 0,
  };
}

/** Default agenda window: the caller's start (today by default) to the next local midnight. */
function agendaRange(
  args: Record<string, unknown>,
  timezone: string,
): { start: string; end: string } | { error: string; hint: string } {
  const today = localDate(Date.now(), timezone);
  const rawStart = stringArg(args, "start") ?? today;
  const start = parseInstantDetailed(rawStart, timezone);
  if ("error" in start) return instantError("start", rawStart, timezone, start.error);
  const rawEnd = stringArg(args, "end");
  let endInstant: number;
  if (rawEnd !== undefined) {
    const end = parseInstantDetailed(rawEnd, timezone);
    if ("error" in end) return instantError("end", rawEnd, timezone, end.error);
    endInstant = end.instant;
  } else {
    // Whole local day from the requested start, DST-safe across midnight.
    const nextDate = new Date(
      Date.parse(`${localDate(start.instant, timezone)}T00:00:00Z`) + 86_400_000,
    )
      .toISOString()
      .slice(0, 10);
    endInstant = localMidnightOfDate(nextDate, timezone);
  }
  if (endInstant <= start.instant) {
    return { error: "end must be after start", hint: "check the requested agenda window" };
  }
  return { start: new Date(start.instant).toISOString(), end: new Date(endInstant).toISOString() };
}

/** One compact agenda row: date, clock bounds, subject, optional id. */
function agendaRow(event: CalendarEvent, timezone: string, includeIds: boolean) {
  const startLocal = event.allDay ? undefined : localParts(event.start, timezone);
  const endLocal = event.allDay ? undefined : localParts(event.end, timezone);
  const date = event.allDay ? event.start.slice(0, 10) : (startLocal?.date ?? event.start.slice(0, 10));
  const time = event.allDay
    ? "all day"
    : `${startLocal?.time ?? event.start}–${endLocal?.time ?? event.end}`;
  return {
    date,
    start_local: event.allDay ? "all day" : (startLocal?.time ?? event.start),
    end_local: event.allDay ? "all day" : (endLocal?.time ?? event.end),
    time,
    subject: event.subject,
    ...(includeIds ? { id: event.id } : {}),
  };
}

export const calendarTool: ToolModule = {
  name: "calendar",
  title: "Calendar",
  description:
    "Outlook calendar (REST v2). op=list (start/end; default today..+7d; top<=50; brief) | agenda (compact rows; ids only with include_ids) | get | create (subject; start default now+15m; end = start+duration_min; no Z/offset = timezone; attendees resolve through the directory and are allowlist-checked; attendees are emailed invites) | update (event_id + expected_subject; patches only supplied fields; one bound alone preserves the other; shift_min) | delete | respond | instances. Conflicts are reported; conflict_policy=refuse writes nothing.",
  shape: {
    op: z.enum(OPS),
    event_id: z.string().optional().describe("event id from a prior read"),
    start: z
      .string()
      .optional()
      .describe("ISO 8601 or YYYY-MM-DD; no Z/offset = timezone; default today/now+15m"),
    end: z.string().optional().describe("ISO 8601 or YYYY-MM-DD; default +7d or start+duration"),
    duration_min: z
      .number()
      .int()
      .min(1)
      .max(10_080)
      .optional()
      .describe("create: minutes if end absent; update: fills the missing bound"),
    shift_min: z
      .number()
      .int()
      .min(-10_080)
      .max(10_080)
      .optional()
      .describe("update: shift start and end (keeps duration)"),
    conflict_policy: z
      .enum(CONFLICT_POLICIES)
      .optional()
      .describe("create/update: warn (default), refuse, ignore"),
    timezone: z.string().optional().describe("IANA zone, default from config"),
    top: z.number().int().min(1).max(50).optional().describe("max events, default 25, cap 50"),
    cursor: z.string().optional().describe("opaque next_cursor from a prior list"),
    brief: z.boolean().optional().describe("list: compact rows (id, subject, start, end, local, show_as)"),
    include_ids: z.boolean().optional().describe("agenda: include event id in each row, default true"),
    include_web_link: z.boolean().optional().describe("list: attach web_link, default false"),
    subject: z.string().optional().describe("create/update: subject"),
    body: z.string().optional().describe("create/update: notes"),
    location: z.string().optional().describe("create/update: location"),
    attendees: z.array(z.string()).optional().describe("create/update: attendee addresses"),
    all_day: z.boolean().optional().describe("create/update: all-day (dates are local)"),
    show_as: z
      .enum(["free", "tentative", "busy", "oof", "working_elsewhere"])
      .optional()
      .describe("free|tentative|busy|oof|working_elsewhere"),
    sensitivity: z
      .enum(["normal", "personal", "private", "confidential"])
      .optional()
      .describe("create/update: sensitivity"),
    response: z.enum(["accept", "decline", "tentative"]).optional().describe("respond: RSVP"),
    comment: z.string().optional().describe("respond: note to organizer"),
    expected_subject: z.string().optional().describe("update/delete/respond guard: current subject must match"),
  },
  async handle(args) {
    const op = opOf(args, OPS);
    if (!op) return unknownOp("calendar", String(args.op ?? ""), OPS);
    const blocked = readOnlyGuard("calendar", op, args);
    if (blocked) return blocked;

    const cfg = config();
    const timezoneLookup = timezoneOf(args, cfg.timezone);
    if ("result" in timezoneLookup) return timezoneLookup.result;
    const timezone = timezoneLookup.timezone;
    const client = outlookClient();

    try {
      if (op === "list") {
        const cursorLookup = resolveCursor(args);
        if (cursorLookup.result) return cursorLookup.result;
        const cursor = cursorLookup.link;
        if (cursor && !client.ownsUrl(cursor)) {
          return errorResult("invalid_cursor", "cursor does not belong to Outlook REST");
        }
        const normalized = normalizeRange(args, timezone);
        if ("error" in normalized) return errorResult("invalid_argument", normalized.error, normalized.hint);
        const range = normalized;
        const query: CalendarQuery = {
          range,
          top: pageSize(args, "top", 25, cfg.maxResults.calendar),
          timezone,
        };
        const brief = booleanArg(args, "brief") === true;
        const includeWebLink = booleanArg(args, "include_web_link") === true;
        if (cursor) {
          const payload = await client.get<unknown>(cursor);
          const collection = readCollection(payload, normalizeEvent);
          return textResult({
            events: collection.items.map((event) => listEventRow(event, timezone, brief, includeWebLink)),
            next_cursor: collection.nextLink ? encodeCursor(collection.nextLink) : undefined,
            truncated: collection.nextLink !== undefined,
          });
        }
        const result = await listCalendarView(client, query);
        return textResult({
          range: {
            start: formatLocal(range.start, timezone),
            end: formatLocal(range.end, timezone),
            timezone,
          },
          count: result.events.length,
          events: result.events.map((event) => listEventRow(event, timezone, brief, includeWebLink)),
          next_cursor: result.nextLink ? encodeCursor(result.nextLink) : undefined,
          truncated: result.nextLink !== undefined,
        });
      }

      if (op === "agenda") {
        const cursorLookup = resolveCursor(args);
        if (cursorLookup.result) return cursorLookup.result;
        const cursor = cursorLookup.link;
        if (cursor && !client.ownsUrl(cursor)) {
          return errorResult("invalid_cursor", "cursor does not belong to Outlook REST");
        }
        const normalized = agendaRange(args, timezone);
        if ("error" in normalized) return errorResult("invalid_argument", normalized.error, normalized.hint);
        const range = normalized;
        const includeIds = booleanArg(args, "include_ids") !== false;
        const top = pageSize(args, "top", 25, cfg.maxResults.calendar);
        const rangeOut = {
          start: formatLocal(range.start, timezone),
          end: formatLocal(range.end, timezone),
          timezone,
        };
        if (cursor) {
          const payload = await client.get<unknown>(cursor);
          const collection = readCollection(payload, normalizeEvent);
          return textResult({
            range: rangeOut,
            count: collection.items.length,
            truncated: collection.nextLink !== undefined,
            events: collection.items.map((event) => agendaRow(event, timezone, includeIds)),
            next_cursor: collection.nextLink ? encodeCursor(collection.nextLink) : undefined,
          });
        }
        const result = await listCalendarView(client, { range, top, timezone });
        return textResult({
          range: rangeOut,
          count: result.events.length,
          truncated: result.nextLink !== undefined,
          events: result.events.map((event) => agendaRow(event, timezone, includeIds)),
          next_cursor: result.nextLink ? encodeCursor(result.nextLink) : undefined,
        });
      }

      if (op === "get") {
        const eventId = stringArg(args, "event_id");
        if (!eventId) return errorResult("invalid_argument", "get requires event_id");
        const event = await getEvent(client, eventId);
        return textResult({
          ...eventRow(event, timezone, true),
          preview: truncate(event.preview, cfg.maxResults.previewChars),
          online_meeting_url: event.onlineMeetingUrl,
          attendees_detail: event.attendees,
        });
      }

      if (op === "create") {
        const subject = stringArg(args, "subject");
        if (!subject) return errorResult("invalid_argument", "create requires subject");
        const startInput = stringArg(args, "start");
        const endInput = stringArg(args, "end");
        const allDay = booleanArg(args, "all_day") ?? false;

        let startMs: number;
        if (startInput) {
          const parsed = parseField("start", startInput, timezone);
          if ("error" in parsed) return errorResult("invalid_argument", parsed.error, parsed.hint);
          startMs = parsed.ms;
        } else {
          startMs = alignUp(Date.now(), 15);
        }
        const durationInput = numberArg(args, "duration_min");
        if (durationInput !== undefined) {
          if (!Number.isInteger(durationInput) || durationInput < 1) {
            return errorResult("invalid_argument", "duration_min must be a positive whole number of minutes");
          }
          if (startInput !== undefined && endInput !== undefined) {
            return errorResult(
              "invalid_argument",
              "duration_min cannot be combined with both start and end",
              "pass start+duration_min, or start+end",
            );
          }
        }
        let endMs: number;
        if (endInput) {
          const parsed = parseField("end", endInput, timezone);
          if ("error" in parsed) return errorResult("invalid_argument", parsed.error, parsed.hint);
          endMs = parsed.ms;
        } else {
          endMs = startMs + (durationInput ?? cfg.defaultDurationMin) * 60_000;
        }
        if (endMs <= startMs) {
          return errorResult("invalid_argument", "end must be after start");
        }
        const policy = conflictPolicyOf(args);
        if (typeof policy === "object") return errorResult("invalid_argument", policy.error);

        // Rate gate before directory resolution: a refused create sends no
        // invite and makes no request. Invalid arguments above consumed nothing.
        const limited = writeRateGuard("calendar", op, {
          attendees: (arrayArg(args, "attendees") ?? []).length,
        });
        if (limited) return limited;
        const resolved = await resolveRecipients(arrayArg(args, "attendees") ?? []);
        if (!resolved.ok) {
          refundWriteToken();
          return resolved.result;
        }
        const attendees = resolved.addresses;
        const blocked = disallowedRecipients(attendees);
        if (blocked.length > 0) {
          refundWriteToken();
          auditWrite({
            tool: "calendar",
            op,
            target: { attendees: blocked.length },
            outcome: "recipient_not_allowed",
          });
          return errorResult(
            "recipient_not_allowed",
            `attendee(s) outside invite_allowlist: ${blocked.join(", ")}`,
            "edit invite_allowlist in settings",
          );
        }
        const showAs = SHOW_AS[stringArg(args, "show_as") ?? ""];
        const sensitivity = SENSITIVITY[stringArg(args, "sensitivity") ?? ""];

        // Read-before-write: the guard/rate/allowlist checks above run first, so a
        // refused create has made no mutating request. `ignore` skips the read.
        let reported: ConflictReport | undefined;
        if (policy !== "ignore") {
          try {
            reported = await findConflicts(
              client,
              timezone,
              new Date(startMs).toISOString(),
              new Date(endMs).toISOString(),
            );
          } catch (error) {
            if (policy === "refuse") {
              refundWriteToken();
              auditWrite({
                tool: "calendar",
                op,
                target: { attendees: attendees.length },
                outcome: "conflict_check_failed",
              });
              throw error;
            }
          }
          if (policy === "refuse" && reported && reported.total > 0) {
            refundWriteToken();
            auditWrite({
              tool: "calendar",
              op,
              target: { attendees: attendees.length },
              outcome: "conflict_detected",
            });
            return conflictRefusal(reported);
          }
        }

        const event = await auditedWrite("calendar", op, { attendees: attendees.length }, () =>
          createEvent(client, {
            subject,
            start: new Date(startMs).toISOString(),
            end: new Date(endMs).toISOString(),
            timezone,
            ...(typeof args.body === "string" ? { body: args.body } : {}),
            ...(stringArg(args, "location") ? { location: stringArg(args, "location") as string } : {}),
            ...(attendees.length > 0 ? { attendees } : {}),
            allDay,
            ...(showAs ? { showAs } : {}),
            ...(sensitivity ? { sensitivity } : {}),
          }),
        );
        return textResult({
          ok: true,
          op: "create",
          ...eventRow(event, timezone),
          ...conflictFields(policy, reported),
          attendees: event.attendees?.map((attendee) => attendee.email ?? attendee.name ?? "?") ?? [],
          // Same detail `get` returns, so the caller need not re-fetch.
          attendees_detail: event.attendees ?? [],
        });
      }

      const guardedOps = op === "update" || op === "delete" || op === "respond";
      if (guardedOps) {
        const eventId = stringArg(args, "event_id");
        const expected = stringArg(args, "expected_subject");
        if (!eventId) return errorResult("invalid_argument", `${op} requires event_id`);
        if (!expected) {
          return errorResult(
            "invalid_argument",
            `${op} requires expected_subject`,
            "read the event first; the guard prevents acting on the wrong object",
          );
        }
        const response = args.response;
        if (
          op === "respond" &&
          response !== "accept" &&
          response !== "decline" &&
          response !== "tentative"
        ) {
          return errorResult("invalid_argument", "respond requires response=accept|decline|tentative");
        }
        if (op === "update" && !UPDATE_FIELDS.some((key) => args[key] !== undefined)) {
          return errorResult("invalid_argument", "update requires at least one field to change");
        }
        const startInput = stringArg(args, "start");
        const endInput = stringArg(args, "end");
        const durationInput = numberArg(args, "duration_min");
        const shiftInput = numberArg(args, "shift_min");
        const policy = conflictPolicyOf(args);
        if (typeof policy === "object") return errorResult("invalid_argument", policy.error);
        if (op === "update") {
          // Ambiguous combinations are rejected before the rate token and the
          // guard read: the caller must say which bound is authoritative.
          if (shiftInput !== undefined) {
            if (startInput !== undefined || endInput !== undefined || durationInput !== undefined) {
              return errorResult(
                "invalid_argument",
                "shift_min cannot be combined with start, end or duration_min",
                "pass shift_min alone; it moves both bounds and keeps the duration",
              );
            }
            if (!Number.isInteger(shiftInput) || shiftInput === 0) {
              return errorResult("invalid_argument", "shift_min must be a non-zero whole number of minutes");
            }
          }
          if (durationInput !== undefined) {
            if (startInput !== undefined && endInput !== undefined) {
              return errorResult(
                "invalid_argument",
                "duration_min cannot be combined with both start and end",
                "pass start+duration_min, end+duration_min, or start+end",
              );
            }
            if (!Number.isInteger(durationInput) || durationInput < 1) {
              return errorResult("invalid_argument", "duration_min must be a positive whole number of minutes");
            }
          }
        }
        // Rate gate before the guard read: a refused op makes no request at all,
        // and invalid arguments above consumed no token.
        const limited = writeRateGuard("calendar", op, { event_id: eventId });
        if (limited) return limited;
        let existing: CalendarEvent;
        try {
          existing = await getEvent(client, eventId);
        } catch (error) {
          refundWriteToken();
          throw error;
        }
        const actual = subjectGuard(existing, expected);
        if (actual !== undefined) {
          refundWriteToken();
          auditWrite({
            tool: "calendar",
            op,
            target: { event_id: eventId },
            outcome: "subject_mismatch",
          });
          return errorResult(
            "subject_mismatch",
            `event ${eventId} subject is "${actual}", not "${expected}"; nothing was changed`,
          );
        }

        if (op === "delete") {
          await auditedWrite("calendar", op, { event_id: eventId }, () => deleteEvent(client, eventId));
          return textResult({ deleted: true, id: eventId, subject: existing.subject });
        }

        if (op === "respond") {
          const comment = typeof args.comment === "string" ? args.comment : undefined;
          await auditedWrite("calendar", op, { event_id: eventId, response }, () =>
            respondToEvent(client, eventId, response as EventResponse, comment),
          );
          return textResult({ responded: true, response, id: eventId });
        }

        const update: EventUpdate = { timezone };
        const subject = stringArg(args, "subject");
        if (subject !== undefined) update.subject = subject;
        if (typeof args.body === "string") update.body = args.body;
        if (typeof args.location === "string") update.location = args.location;
        if (startInput !== undefined) update.start = startInput;
        if (endInput !== undefined) update.end = endInput;
        const allDayInput = booleanArg(args, "all_day");
        if (allDayInput !== undefined) update.allDay = allDayInput;
        const showAsInput = SHOW_AS[stringArg(args, "show_as") ?? ""];
        if (showAsInput) update.showAs = showAsInput;
        const sensitivityInput = SENSITIVITY[stringArg(args, "sensitivity") ?? ""];
        if (sensitivityInput) update.sensitivity = sensitivityInput;
        const rawAttendees = arrayArg(args, "attendees");
        if (rawAttendees) {
          const resolvedAttendees = await resolveRecipients(rawAttendees);
          if (!resolvedAttendees.ok) {
            refundWriteToken();
            return resolvedAttendees.result;
          }
          const blocked = disallowedRecipients(resolvedAttendees.addresses);
          if (blocked.length > 0) {
            refundWriteToken();
            auditWrite({
              tool: "calendar",
              op,
              target: { event_id: eventId, attendees: blocked.length },
              outcome: "recipient_not_allowed",
            });
            return errorResult(
              "recipient_not_allowed",
              `attendee(s) outside invite_allowlist: ${blocked.join(", ")}`,
              "edit invite_allowlist in settings",
            );
          }
          update.attendees = resolvedAttendees.addresses;
        }
        const timeChanging =
          startInput !== undefined ||
          endInput !== undefined ||
          durationInput !== undefined ||
          shiftInput !== undefined;
        if (Object.keys(update).length === 1 && !timeChanging) {
          refundWriteToken();
          return errorResult("invalid_argument", "update requires at least one field to change");
        }
        let changed = changedReport(existing.start, existing.end, existing.start, existing.end);
        if (timeChanging) {
          const oldStartMs = Date.parse(existing.start);
          const oldEndMs = Date.parse(existing.end);
          if (Number.isNaN(oldStartMs) || Number.isNaN(oldEndMs) || oldEndMs <= oldStartMs) {
            refundWriteToken();
            return errorResult(
              "invalid_argument",
              "cannot read the existing event times",
              "read the event with op=get before moving it",
            );
          }
          const oldDurationMs = oldEndMs - oldStartMs;
          let startMs: number;
          let endMs: number;
          if (shiftInput !== undefined) {
            startMs = oldStartMs + shiftInput * 60_000;
            endMs = oldEndMs + shiftInput * 60_000;
          } else if (durationInput !== undefined) {
            const durationMs = durationInput * 60_000;
            if (startInput !== undefined) {
              const parsed = parseField("start", startInput, timezone);
              if ("error" in parsed) {
                refundWriteToken();
                return errorResult("invalid_argument", parsed.error, parsed.hint);
              }
              startMs = parsed.ms;
              endMs = startMs + durationMs;
            } else if (endInput !== undefined) {
              const parsed = parseField("end", endInput, timezone);
              if ("error" in parsed) {
                refundWriteToken();
                return errorResult("invalid_argument", parsed.error, parsed.hint);
              }
              endMs = parsed.ms;
              startMs = endMs - durationMs;
            } else {
              startMs = oldStartMs;
              endMs = startMs + durationMs;
            }
          } else if (startInput !== undefined) {
            const parsed = parseField("start", startInput, timezone);
            if ("error" in parsed) {
              refundWriteToken();
              return errorResult("invalid_argument", parsed.error, parsed.hint);
            }
            // Moving the start keeps the existing DURATION, not the old end.
            startMs = parsed.ms;
            endMs = startMs + oldDurationMs;
          } else {
            const parsed = parseField("end", endInput as string, timezone);
            if ("error" in parsed) {
              refundWriteToken();
              return errorResult("invalid_argument", parsed.error, parsed.hint);
            }
            endMs = parsed.ms;
            startMs = oldStartMs;
          }
          if (endMs <= startMs) {
            refundWriteToken();
            return errorResult("invalid_argument", "end must be after start");
          }
          update.start = new Date(startMs).toISOString();
          update.end = new Date(endMs).toISOString();
          changed = changedReport(existing.start, existing.end, update.start, update.end);
        }

        // A subject-only edit cannot create a new overlap, so the conflict read
        // runs only when the occupied window changes. It excludes this event.
        let reported: ConflictReport | undefined;
        if (timeChanging && policy !== "ignore") {
          try {
            reported = await findConflicts(
              client,
              timezone,
              update.start as string,
              update.end as string,
              eventId,
            );
          } catch (error) {
            if (policy === "refuse") {
              refundWriteToken();
              auditWrite({
                tool: "calendar",
                op,
                target: { event_id: eventId },
                outcome: "conflict_check_failed",
              });
              throw error;
            }
          }
          if (policy === "refuse" && reported && reported.total > 0) {
            refundWriteToken();
            auditWrite({
              tool: "calendar",
              op,
              target: { event_id: eventId },
              outcome: "conflict_detected",
            });
            return conflictRefusal(reported);
          }
        }

        const updated = await auditedWrite("calendar", op, { event_id: eventId }, () =>
          patchEvent(client, eventId, buildEventPatch(update, existing)),
        );
        return textResult({
          ok: true,
          op: "update",
          ...eventRow(updated, timezone, true),
          changed,
          ...(timeChanging ? conflictFields(policy, reported) : {}),
          attendees_detail: updated.attendees ?? [],
        });
      }

      if (op === "instances") {
        const eventId = stringArg(args, "event_id");
        if (!eventId) return errorResult("invalid_argument", "instances requires event_id");
        const cursorLookup = resolveCursor(args);
        if (cursorLookup.result) return cursorLookup.result;
        const cursor = cursorLookup.link;
        if (cursor && !client.ownsUrl(cursor)) {
          return errorResult("invalid_cursor", "cursor does not belong to Outlook REST");
        }
        const normalized = normalizeRange(args, timezone);
        if ("error" in normalized) return errorResult("invalid_argument", normalized.error, normalized.hint);
        const range = normalized;
        const page = await listInstances(
          client,
          eventId,
          range,
          pageSize(args, "top", 25, cfg.maxResults.calendar),
          cursor,
        );
        return textResult({
          event_id: eventId,
          range: {
            start: formatLocal(range.start, timezone),
            end: formatLocal(range.end, timezone),
            timezone,
          },
          count: page.instances.length,
          instances: page.instances.map((event) => eventRow(event, timezone)),
          next_cursor: page.nextLink ? encodeCursor(page.nextLink) : undefined,
          truncated: page.nextLink !== undefined,
        });
      }

      return unknownOp("calendar", op, OPS);
    } catch (error) {
      const mapped = toolError(error);
      if (mapped) return mapped;
      throw error;
    }
  },
};
