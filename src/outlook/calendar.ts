/**
 * Outlook REST v2.0 calendar and scheduling.
 *
 * Verified behaviours on the tested tenant:
 *  - `GET  /me/calendarview` works; there is NO delta support (405) — snapshots
 *    are time-window polls;
 *  - `POST /me/calendar/getschedule` and `POST /users/{address}/calendar/getschedule`
 *    return busy blocks for other mailboxes;
 *  - `POST /me/findmeetingtimes` returns ranked meeting suggestions.
 *
 * Request bodies are PascalCase (REST v2), responses are normalised through
 * `camelizeKeys` before mapping. All instants leave this module as UTC ISO 8601:
 * we request `TimeZone: "UTC"` and append `Z` when the service omits an offset.
 */
import {
  camelizeKeys,
  normalizeEvent,
  objectRows,
  readCollection,
  restDateTimeToIso,
  type CalendarEvent,
  type OutlookClient,
} from "./client.js";
export { restDateTimeToIso } from "./client.js";
import { localDate, wallClock } from "../time.js";
import type { BusyBlock, BusyStatus, CalendarSource, Range, ScheduleAttendee, SnapshotEvent } from "../snapshot/calendar.js";
import type { EmailAddress } from "./client.js";

export interface CalendarQuery {
  range: Range;
  top?: number;
  /** IANA zone used for working-hours reasoning only; the API is asked for UTC. */
  timezone?: string;
}

const STATUS_BY_REST: Record<string, BusyStatus> = {
  free: "free",
  tentative: "tentative",
  busy: "busy",
  oof: "oof",
  workingElsewhere: "working_elsewhere",
  unknown: "busy",
};

function statusOf(value: unknown): BusyStatus {
  if (typeof value !== "string") return "busy";
  return STATUS_BY_REST[value] ?? STATUS_BY_REST[value.toLowerCase()] ?? "busy";
}

/** `GET /me/calendarview` — own events in a time window. */
export async function listCalendarView(
  client: OutlookClient,
  query: CalendarQuery,
): Promise<{ events: CalendarEvent[]; nextLink?: string }> {
  const params = new URLSearchParams({
    startDateTime: query.range.start,
    endDateTime: query.range.end,
    $select:
      "Id,Subject,Start,End,IsAllDay,ShowAs,Location,Organizer,Attendees,IsOnlineMeeting,OnlineMeetingUrl,WebLink,BodyPreview",
    $orderby: "Start/DateTime",
    $top: String(query.top ?? 100),
  });
  const payload = await client.get<Record<string, unknown>>(`me/calendarview?${params.toString()}`);
  const rows = objectRows(payload?.value);
  const events = rows.map((row) => normalizeEvent(camelizeKeys<Record<string, unknown>>(row)));
  const nextLink = typeof payload?.["@odata.nextLink"] === "string" ? payload["@odata.nextLink"] : undefined;
  return nextLink ? { events, nextLink } : { events };
}

/** `GET /me/events/{id}` — one event with body. */
export async function getEvent(client: OutlookClient, eventId: string): Promise<CalendarEvent> {
  const payload = await client.get<Record<string, unknown>>(
    `me/events/${encodeURIComponent(eventId)}?$select=` +
      encodeURIComponent(
        "Id,Subject,Start,End,IsAllDay,ShowAs,Location,Organizer,Attendees,IsOnlineMeeting,OnlineMeetingUrl,WebLink,BodyPreview",
      ),
  );
  return normalizeEvent(camelizeKeys<Record<string, unknown>>(payload));
}

export interface ScheduleRequest {
  attendees: string[];
  range: Range;
  timezone?: string;
  intervalMinutes?: number;
}

/**
 * `POST /me/calendar/getschedule` — busy blocks per mailbox. Times are requested
 * in UTC so no offset guessing is needed.
 */
export async function getSchedule(
  client: OutlookClient,
  request: ScheduleRequest,
): Promise<ScheduleAttendee[]> {
  const body = {
    Schedules: request.attendees,
    StartTime: { DateTime: toRestUtc(request.range.start), TimeZone: "UTC" },
    EndTime: { DateTime: toRestUtc(request.range.end), TimeZone: "UTC" },
    AvailabilityViewInterval: request.intervalMinutes ?? 30,
  };
  const payload = await client.post<Record<string, unknown>>("me/calendar/getschedule", body);
  const rows = objectRows(payload?.value);
  return rows.map((raw): ScheduleAttendee => {
    const row = camelizeKeys<Record<string, unknown>>(raw);
    const items = Array.isArray(row.scheduleItems) ? row.scheduleItems : [];
    const blocks: BusyBlock[] = items.flatMap((entry): BusyBlock[] => {
      const item = camelizeKeys<Record<string, unknown>>(entry);
      const blockStart = item.start as Record<string, unknown> | undefined;
      const blockEnd = item.end as Record<string, unknown> | undefined;
      const start = restDateTimeToIso(blockStart?.dateTime, blockStart?.timeZone as string | undefined);
      const end = restDateTimeToIso(blockEnd?.dateTime, blockEnd?.timeZone as string | undefined);
      if (!start || !end) return [];
      return [{ start, end, status: statusOf(item.status) }];
    });
    // Exchange omits `WorkingHours` when it cannot resolve the mailbox (typo,
    // foreign address). An empty `ScheduleItems` array alone is indistinguishable
    // from a genuinely free person, so this is the only safe signal.
    const unknown = row.workingHours === undefined || row.workingHours === null;
    const availabilityView =
      typeof row.availabilityView === "string" && row.availabilityView !== "" ? row.availabilityView : undefined;
    // WorkingHours is the primary signal, but a row that carries schedule items
    // or an availability view is also real data even when Exchange omitted the
    // hours. `unknown` stays as-is for callers that must not treat it as free.
    const dataAvailable = !unknown || items.length > 0 || availabilityView !== undefined;
    const errorRow =
      row.error && typeof row.error === "object" ? camelizeKeys<Record<string, unknown>>(row.error) : undefined;
    const code = errorRow?.responseCode;
    const error =
      errorRow && (typeof code === "string" || typeof code === "number" || typeof errorRow.message === "string")
        ? {
            ...(typeof code === "string" || typeof code === "number" ? { code: String(code) } : {}),
            ...(typeof errorRow.message === "string" ? { message: errorRow.message } : {}),
          }
        : undefined;
    return {
      email: typeof row.scheduleId === "string" ? row.scheduleId : "",
      blocks,
      ...(availabilityView ? { availability_view: availabilityView } : {}),
      data_available: dataAvailable,
      ...(unknown ? { unknown: true as const } : {}),
      ...(error ? { error } : {}),
    };
  });
}

/** REST DateTime literal: `2026-11-16T15:00:00` without the trailing `Z`. */
function toRestUtc(iso: string): string {
  return iso.replace(/\.\d{3}Z$/, "").replace(/Z$/, "");
}

export interface MeetingSuggestion {
  /** ISO 8601 UTC. */
  start: string;
  end: string;
  confidence?: number;
  reason?: string;
  attendeeAvailability: Array<EmailAddress & { availability: BusyStatus }>;
}

export interface FindMeetingTimesRequest {
  attendees: string[];
  range: Range;
  durationMin: number;
  maxCandidates?: number;
  /** Return slots where some attendees are busy instead of dropping them. */
  allowConflicts?: boolean;
}

/**
 * `POST /me/findmeetingtimes` — service-side candidate suggestions.
 *
 * `IsOrganizerOptional: false` because the signed-in user's own calendar must
 * also be free; `MinimumAttendeePercentage` drops candidates where an attendee is
 * unavailable.
 */
export async function findMeetingTimes(
  client: OutlookClient,
  request: FindMeetingTimesRequest,
): Promise<MeetingSuggestion[]> {
  const body = {
    Attendees: request.attendees.map((address) => ({
      Type: "Required",
      EmailAddress: { Address: address },
    })),
    TimeConstraint: {
      ActivityDomain: "Unrestricted",
      Timeslots: [
        {
          Start: { DateTime: toRestUtc(request.range.start), TimeZone: "UTC" },
          End: { DateTime: toRestUtc(request.range.end), TimeZone: "UTC" },
        },
      ],
    },
    MeetingDuration: `PT${Math.max(1, Math.round(request.durationMin))}M`,
    MaxCandidates: String(request.maxCandidates ?? 5),
    IsOrganizerOptional: false,
    ReturnSuggestionReasons: true,
    MinimumAttendeePercentage: request.allowConflicts ? 1 : 100,
  };
  const payload = await client.post<Record<string, unknown>>("me/findmeetingtimes", body);
  // REST v2 returns MeetingTimeSuggestions; capture tooling and some proxies
  // wrap it in `value`. Accept both.
  const container =
    payload?.meetingTimeSuggestions ??
    payload?.MeetingTimeSuggestions ??
    payload?.value;
  const rows = objectRows(container);
  return rows.flatMap((raw): MeetingSuggestion[] => {
    const row = camelizeKeys<Record<string, unknown>>(raw);
    const slot = row.meetingTimeSlot as Record<string, unknown> | undefined;
    const slotStart = slot?.start as Record<string, unknown> | undefined;
    const slotEnd = slot?.end as Record<string, unknown> | undefined;
    const start = restDateTimeToIso(slotStart?.dateTime, slotStart?.timeZone as string | undefined);
    const end = restDateTimeToIso(slotEnd?.dateTime, slotEnd?.timeZone as string | undefined);
    if (!start || !end) return [];
    const availability = Array.isArray(row.attendeeAvailability)
      ? row.attendeeAvailability.flatMap((entry) => {
          const item = camelizeKeys<Record<string, unknown>>(entry);
          const address = camelizeKeys<Record<string, unknown>>(item.attendee ?? {});
          const emailAddress = address.emailAddress as Record<string, unknown> | undefined;
          const email = typeof emailAddress?.address === "string" ? emailAddress.address : undefined;
          if (!email) return [];
          return [
            {
              email,
              ...(typeof emailAddress?.name === "string" ? { name: emailAddress.name } : {}),
              availability: statusOf(item.availability),
            },
          ];
        })
      : [];
    return [
      {
        start,
        end,
        ...(typeof row.confidence === "number" ? { confidence: row.confidence } : {}),
        ...(typeof row.suggestionReason === "string" ? { reason: row.suggestionReason } : {}),
        attendeeAvailability: availability,
      },
    ];
  });
}

/** Snapshot data source backed by the Outlook REST calendar. */
export function outlookCalendarSource(client: OutlookClient): CalendarSource {
  return {
    async ownEvents(range) {
      const { events } = await listCalendarView(client, { range });
      return events.flatMap((event): SnapshotEvent[] => {
        if (!event.id || !event.start || !event.end) return [];
        return [
          {
            id: event.id,
            subject: event.subject.slice(0, 200),
            start: event.start,
            end: event.end,
            ...(event.allDay !== undefined ? { all_day: event.allDay } : {}),
            ...(event.showAs ? { show_as: statusOf(event.showAs) } : {}),
            ...(event.location ? { location: event.location.slice(0, 120) } : {}),
            ...(event.isOnlineMeeting !== undefined ? { is_online: event.isOnlineMeeting } : {}),
          },
        ];
      });
    },

    async schedules(attendees, range, timezone, intervalMinutes) {
      return getSchedule(client, { attendees, range, timezone, intervalMinutes });
    },
  };
}

/* ------------------------------------------------------------------ */
/* Writes                                                              */
/* ------------------------------------------------------------------ */

export type EventShowAs = "Free" | "Tentative" | "Busy" | "Oof" | "WorkingElsewhere";
export type EventSensitivity = "Normal" | "Personal" | "Private" | "Confidential";
export type EventResponse = "accept" | "decline" | "tentative";

const RESPONSE_ACTION: Record<EventResponse, string> = {
  accept: "accept",
  decline: "decline",
  tentative: "tentativelyaccept",
};

/** `{DateTime, TimeZone}` in the caller's zone, with .NET-style 7-digit fraction. */
export function eventTime(instant: string, timezone: string): { DateTime: string; TimeZone: string } {
  const wall = wallClock(instant, timezone);
  return { DateTime: `${wall}.0000000`, TimeZone: timezone };
}

/** Midnight of a local date, as an Outlook wall-clock literal. */
function allDayTime(date: string, timezone: string): { DateTime: string; TimeZone: string } {
  return { DateTime: `${date}T00:00:00.0000000`, TimeZone: timezone };
}

function addDays(date: string, days: number): string {
  const parsed = Date.parse(`${date}T00:00:00Z`);
  return new Date(parsed + days * 86_400_000).toISOString().slice(0, 10);
}

function attendeePayload(addresses: string[]): Array<Record<string, unknown>> {
  return addresses.map((address) => ({
    EmailAddress: { Address: address },
    Type: "Required",
  }));
}

export interface EventDraft {
  subject: string;
  /** ISO 8601 instants; for all-day events only the local dates matter. */
  start: string;
  end: string;
  timezone: string;
  body?: string;
  location?: string;
  attendees?: string[];
  allDay?: boolean;
  showAs?: EventShowAs;
  sensitivity?: EventSensitivity;
}

/**
 * PascalCase `/me/events` body. `IsReminderSet` is deliberately absent: it does
 * not exist on the v2 Event entity and makes the POST fail with
 * `400 UnableToDeserializePostBody`.
 */
export function buildEventBody(draft: EventDraft): Record<string, unknown> {
  const times = draft.allDay ? allDayTimes(draft) : { start: eventTime(draft.start, draft.timezone), end: eventTime(draft.end, draft.timezone) };
  return {
    Subject: draft.subject,
    ...(draft.body !== undefined ? { Body: { ContentType: "Text", Content: draft.body } } : {}),
    Start: times.start,
    End: times.end,
    ...(draft.location ? { Location: { DisplayName: draft.location } } : {}),
    ...(draft.attendees && draft.attendees.length > 0 ? { Attendees: attendeePayload(draft.attendees) } : {}),
    ...(draft.allDay !== undefined ? { IsAllDay: draft.allDay } : {}),
    ...(draft.showAs ? { ShowAs: draft.showAs } : {}),
    ...(draft.sensitivity ? { Sensitivity: draft.sensitivity } : {}),
  };
}

function allDayTimes(draft: Pick<EventDraft, "start" | "end" | "timezone">): {
  start: { DateTime: string; TimeZone: string };
  end: { DateTime: string; TimeZone: string };
} {
  const startDate = localDate(draft.start, draft.timezone);
  const endDateRaw = localDate(draft.end, draft.timezone);
  const endDate = endDateRaw > startDate ? endDateRaw : addDays(startDate, 1);
  return {
    start: allDayTime(startDate, draft.timezone),
    end: allDayTime(endDate, draft.timezone),
  };
}

export async function createEvent(client: OutlookClient, draft: EventDraft): Promise<CalendarEvent> {
  const payload = await client.post<Record<string, unknown>>("me/events", buildEventBody(draft));
  return normalizeEvent(camelizeKeys<Record<string, unknown>>(payload ?? {}));
}

export interface EventUpdate {
  timezone: string;
  subject?: string;
  body?: string;
  start?: string;
  end?: string;
  allDay?: boolean;
  location?: string;
  attendees?: string[];
  showAs?: EventShowAs;
  sensitivity?: EventSensitivity;
}

/**
 * Build a PATCH body containing only the supplied fields. Start/End are sent
 * together (the API requires a consistent pair); a missing half is taken from
 * the event as read.
 */
export function buildEventPatch(update: EventUpdate, existing: CalendarEvent): Record<string, unknown> {
  const patch: Record<string, unknown> = {};
  if (update.subject !== undefined) patch.Subject = update.subject;
  if (update.body !== undefined) patch.Body = { ContentType: "Text", Content: update.body };
  if (update.location !== undefined) patch.Location = { DisplayName: update.location };
  if (update.attendees !== undefined) patch.Attendees = attendeePayload(update.attendees);
  if (update.showAs !== undefined) patch.ShowAs = update.showAs;
  if (update.sensitivity !== undefined) patch.Sensitivity = update.sensitivity;
  if (update.allDay !== undefined) patch.IsAllDay = update.allDay;

  if (update.start !== undefined || update.end !== undefined) {
    const start = update.start ?? existing.start;
    const end = update.end ?? existing.end;
    const allDay = update.allDay ?? existing.allDay ?? false;
    const times = allDay
      ? allDayTimes({ start, end, timezone: update.timezone })
      : { start: eventTime(start, update.timezone), end: eventTime(end, update.timezone) };
    patch.Start = times.start;
    patch.End = times.end;
    if (update.allDay === undefined) patch.IsAllDay = allDay;
  }
  return patch;
}

export async function patchEvent(
  client: OutlookClient,
  eventId: string,
  patch: Record<string, unknown>,
): Promise<CalendarEvent> {
  const payload = await client.patch<Record<string, unknown>>(
    `me/events/${encodeURIComponent(eventId)}`,
    patch,
  );
  return normalizeEvent(camelizeKeys<Record<string, unknown>>(payload ?? {}));
}

/** `DELETE /me/events/{id}` → 204. A non-2xx response throws. */
export async function deleteEvent(client: OutlookClient, eventId: string): Promise<void> {
  await client.delete<void>(`me/events/${encodeURIComponent(eventId)}`);
}

export async function respondToEvent(
  client: OutlookClient,
  eventId: string,
  response: EventResponse,
  comment?: string,
): Promise<void> {
  const action = RESPONSE_ACTION[response];
  if (!action) throw new Error(`unsupported response "${response}"`);
  await client.post<void>(`me/events/${encodeURIComponent(eventId)}/${action}`, {
    ...(comment ? { Comment: comment } : {}),
    SendResponse: true,
  });
}

/** Occurrences of a recurring event in a window, one page at a time. */
export interface InstancesPage {
  instances: CalendarEvent[];
  /** `@odata.nextLink` when the window holds more occurrences than one page. */
  nextLink?: string;
}

export async function listInstances(
  client: OutlookClient,
  eventId: string,
  range: Range,
  top = 50,
  cursor?: string,
): Promise<InstancesPage> {
  const params = new URLSearchParams({
    startDateTime: range.start,
    endDateTime: range.end,
    $select: "Id,Subject,Start,End,IsAllDay,ShowAs,Location,Organizer,Attendees,IsOnlineMeeting,WebLink",
    $top: String(top),
  });
  const payload = await client.get<unknown>(
    cursor ?? `me/events/${encodeURIComponent(eventId)}/instances?${params.toString()}`,
  );
  const collection = readCollection(payload, normalizeEvent);
  // The instances endpoint returns occurrences unordered; order each page so a
  // caller merging pages still sees a sane sequence.
  const instances = [...collection.items].sort((a, b) => a.start.localeCompare(b.start));
  return collection.nextLink ? { instances, nextLink: collection.nextLink } : { instances };
}

/** Normalise a subject for guard comparisons (case/whitespace insensitive). */
export function normalizeSubject(value: string | undefined): string {
  return (value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}
