/**
 * `availability` tool: freebusy | suggest | snapshot.
 *
 * Busy blocks come from Outlook REST `POST /me/calendar/getschedule` (via the
 * calendar snapshot). `suggest` asks the service for candidates with
 * `POST /me/findmeetingtimes` and falls back to local slot search over the
 * snapshot's busy blocks when that endpoint is unavailable.
 *
 * A refresh hits the network only when the snapshot is stale, does not cover the
 * requested range, or `refresh` is set.
 */
import { z } from "zod";
import { config } from "../config.js";
import { AuthRequiredError, readTokenCache } from "../auth/store.js";
import { ApiError } from "../http/client.js";
import {
  outlookCalendarSource,
  findMeetingTimes,
  type MeetingSuggestion,
} from "../outlook/calendar.js";
import { outlookClient } from "../outlook/client.js";
import {
  buildAvailabilityView,
  busyFor,
  busyStatusIn,
  freeWindows,
  overlappingBusy,
  summarizeBusyBlocks,
  suggestSlots,
  type SlotCandidate,
} from "../availability/slots.js";
import {
  defaultRange,
  ensureSnapshot,
  hasScheduleData,
  isUnknownAttendee,
  SnapshotError,
  summarizeSnapshot,
  type BusyBlock,
} from "../snapshot/calendar.js";
import { formatLocal, parseInstantDetailed, withinBlockedTime, withinWorkingHours } from "../time.js";
import {
  arrayArg,
  booleanArg,
  errorResult,
  numberArg,
  opOf,
  stringArg,
  textResult,
  timezoneOf,
  toolError,
  unknownOp,
  type ToolModule,
} from "./shared.js";
import { resolveRecipients } from "./recipients.js";

const OPS = ["freebusy", "suggest", "snapshot", "check", "free"] as const;

/**
 * Local slot search and availability-view construction are O(window/interval),
 * so an unbounded date range (e.g. year 1 to 9999) would spin the event loop and
 * allocate gigabytes of digits. A year is far beyond any scheduling question.
 */
const MAX_WINDOW_DAYS = 366;

/** Intersect busy blocks with [start, end]; blocks outside the window vanish. */
function clipBusy(blocks: BusyBlock[], startIso: string, endIso: string): BusyBlock[] {
  const windowStart = Date.parse(startIso);
  const windowEnd = Date.parse(endIso);
  return blocks.flatMap((block): BusyBlock[] => {
    const blockStart = Date.parse(block.start);
    const blockEnd = Date.parse(block.end);
    if ([windowStart, windowEnd, blockStart, blockEnd].some((value) => Number.isNaN(value))) return [];
    const start = Math.max(blockStart, windowStart);
    const end = Math.min(blockEnd, windowEnd);
    if (end <= start) return [];
    return [{ ...block, start: new Date(start).toISOString(), end: new Date(end).toISOString() }];
  });
}

/** The signed-in user's address, from the token cache when available. */
async function selfAddress(): Promise<string | undefined> {
  const { cache } = await readTokenCache();
  return cache?.account?.upn;
}

export const availabilityTool: ToolModule = {
  name: "availability",
  title: "Availability",
  description:
    "Scheduling (read-only). op=check (start+end; attendees[] default self) -> one slot: available, per-attendee status/data_available; an unknown mailbox is never available | op=free (busy gaps from snapshot; duration_min from config, max_slots<=20) -> explicit free windows | op=freebusy (attendees[], start, end) -> busy blocks + availabilityView digits (0 free 1 tentative 2 busy 3 oof 4 elsewhere); data_available=false means no schedule data, not free | op=suggest (duration_min 30, interval_min 15, buffer_min 0, working_hours_only, max_candidates<=20) -> ranked slots inside [start,end]; a subset, not every free window - use freebusy for the full view | op=snapshot writes the YAML resource. Names resolve through the directory, never guessed.",
  shape: {
    op: z.enum(OPS),
    attendees: z
      .array(z.string())
      .optional()
      .describe("names, aliases, or addresses; default self"),
    start: z.string().optional().describe("ISO 8601 window start; default today (check: required)"),
    end: z.string().optional().describe("ISO 8601 window end; default +7d (check: required)"),
    duration_min: z.number().int().min(5).max(480).optional().describe("suggest/free: slot length, default suggest 30 / config for free"),
    interval_min: z.number().int().min(5).max(120).optional().describe("suggest: slot grid, default 15"),
    buffer_min: z.number().int().min(0).max(120).optional().describe("suggest/free: padding around meetings, default 0"),
    working_hours_only: z.boolean().optional().describe("suggest/free: Mon-Fri 09:00-17:00 local by default"),
    max_candidates: z
      .number()
      .int()
      .min(1)
      .optional()
      .describe("suggest: results, default 5, max 20"),
    max_slots: z
      .number()
      .int()
      .min(1)
      .max(20)
      .optional()
      .describe("free: windows returned, default max_results.slots, clamped to 20"),
    timezone: z.string().optional().describe("IANA zone, default from config"),
    refresh: z.boolean().optional().describe("force a snapshot refresh"),
  },
  async handle(args) {
    const op = opOf(args, OPS);
    if (!op) return unknownOp("availability", String(args.op ?? ""), OPS);
    if (op === "check" && (stringArg(args, "start") === undefined || stringArg(args, "end") === undefined)) {
      return errorResult(
        "invalid_argument",
        "check requires both start and end for the exact proposed slot",
        "pass the slot boundaries, e.g. start=2026-11-18T15:00 end=2026-11-18T15:30",
      );
    }

    const cfg = config();
    const timezoneLookup = timezoneOf(args, cfg.timezone);
    if ("result" in timezoneLookup) return timezoneLookup.result;
    const timezone = timezoneLookup.timezone;
    const range = defaultRange(timezone);
    const startInput = typeof args.start === "string" && args.start ? args.start : range.start;
    const endInput = typeof args.end === "string" && args.end ? args.end : range.end;
    const startParsed = parseInstantDetailed(startInput, timezone);
    if ("error" in startParsed) {
      return errorResult(
        "invalid_argument",
        startParsed.error === "nonexistent"
          ? `start "${startInput}" does not exist in ${timezone} (DST spring-forward gap)`
          : `start "${startInput}" is not an ISO 8601 datetime`,
        "use ISO 8601 (with or without an offset) or YYYY-MM-DD; a bare date means local midnight",
      );
    }
    const endParsed = parseInstantDetailed(endInput, timezone);
    if ("error" in endParsed) {
      return errorResult(
        "invalid_argument",
        endParsed.error === "nonexistent"
          ? `end "${endInput}" does not exist in ${timezone} (DST spring-forward gap)`
          : `end "${endInput}" is not an ISO 8601 datetime`,
        "use ISO 8601 (with or without an offset) or YYYY-MM-DD; a bare date means local midnight",
      );
    }
    const windowStart = new Date(startParsed.instant).toISOString();
    const windowEnd = new Date(endParsed.instant).toISOString();
    if (Date.parse(windowEnd) <= Date.parse(windowStart)) {
      return errorResult("invalid_argument", "end must be after start");
    }
    const windowDays = (Date.parse(windowEnd) - Date.parse(windowStart)) / 86_400_000;
    if (windowDays > MAX_WINDOW_DAYS) {
      return errorResult(
        "invalid_argument",
        `window is ${Math.ceil(windowDays)} days; availability is limited to ${MAX_WINDOW_DAYS} days`,
        "narrow start/end (the default window is 7 days), or take a snapshot per month",
      );
    }

    let requested: string[] = [];
    if (arrayArg(args, "attendees")) {
      try {
        const resolved = await resolveRecipients(arrayArg(args, "attendees") ?? []);
        if (!resolved.ok) return resolved.result;
        requested = resolved.addresses;
      } catch (error) {
        const mapped = toolError(error);
        if (mapped) return mapped;
        throw error;
      }
    }
    const self = await selfAddress();
    const attendees = requested.length > 0 ? requested : self ? [self] : [];
    const selfOnly =
      self !== undefined &&
      attendees.length === 1 &&
      attendees[0]!.trim().toLowerCase() === self.trim().toLowerCase();
    const refresh = booleanArg(args, "refresh") ?? false;

    if (attendees.length === 0) {
      return errorResult(
        "auth_required",
        "no attendees: pass attendees[] or sign in so the signed-in user is known",
        "run auth login",
      );
    }

    try {
      if (op === "snapshot") {
        const result = await ensureSnapshot({
          source: outlookCalendarSource(outlookClient()),
          attendees,
          start: windowStart,
          end: windowEnd,
          timezone,
          refresh,
        });
        return textResult({
          ...summarizeSnapshot(result.snapshot, result.path),
          refreshed: result.refreshed,
          url: "teams-mcp://calendar/snapshot",
        });
      }

      const snapshotData = await ensureSnapshot({
        source: outlookCalendarSource(outlookClient()),
        attendees,
        start: windowStart,
        end: windowEnd,
        timezone,
        refresh,
      });
      const snapshot = snapshotData.snapshot;

      if (op === "check") {
        const unknown = attendees.filter((email) => isUnknownAttendee(snapshot, email));
        const known = attendees.filter((email) => !unknown.includes(email));
        const startMs = Date.parse(windowStart);
        const endMs = Date.parse(windowEnd);
        const rows = attendees.map((email) => {
          const isUnknown = unknown.includes(email);
          const status = isUnknown
            ? "unknown"
            : (busyStatusIn(busyFor(snapshot.busy, email), startMs, endMs) ?? "free");
          return {
            email,
            data_available: hasScheduleData(snapshot, email),
            unknown: isUnknown,
            status,
          };
        });
        const conflicts = known
          .flatMap((email) =>
            overlappingBusy(busyFor(snapshot.busy, email), startMs, endMs).map((block) => {
              const start = Math.max(Date.parse(block.start), startMs);
              const end = Math.min(Date.parse(block.end), endMs);
              return {
                email,
                status: block.status,
                start: new Date(start).toISOString(),
                end: new Date(end).toISOString(),
                local: `${formatLocal(start, timezone)} → ${formatLocal(end, timezone)}`,
              };
            }),
          )
          .sort((a, b) => a.start.localeCompare(b.start));
        // No data for a mailbox is not "free": an unresolved attendee makes the
        // slot unverifiable, never available.
        const available =
          attendees.length > 0 &&
          unknown.length === 0 &&
          rows.every((row) => row.status === "free");
        return textResult({
          available,
          range: {
            start: formatLocal(windowStart, timezone),
            end: formatLocal(windowEnd, timezone),
            timezone,
          },
          attendees: rows,
          conflicts: conflicts.slice(0, 5),
          conflict_count: conflicts.length,
          unknown,
        });
      }

      if (op === "free") {
        const unknown = attendees.filter((email) => isUnknownAttendee(snapshot, email));
        const known = attendees.filter((email) => !unknown.includes(email));
        const durationMin = Math.min(
          480,
          Math.max(5, numberArg(args, "duration_min") ?? cfg.defaultDurationMin),
        );
        const maxSlots = Math.min(20, Math.max(1, numberArg(args, "max_slots") ?? cfg.maxResults.slots));
        const workingHoursOnly = booleanArg(args, "working_hours_only") ?? true;
        const result =
          known.length === 0
            ? { windows: [], total: 0, truncated: false }
            : freeWindows({
                busy: snapshot.busy,
                attendees: known,
                start: windowStart,
                end: windowEnd,
                durationMin,
                bufferMin: Math.min(120, Math.max(0, numberArg(args, "buffer_min") ?? 0)),
                workingHoursOnly,
                timezone,
                workHours: cfg.workHours,
                blockedTimes: cfg.blockedTimes,
                maxSlots,
              });
        return textResult({
          slots: result.windows.map((window) => ({
            start: window.start,
            end: window.end,
            local: `${window.local_start} → ${window.local_end}`,
          })),
          count: result.windows.length,
          truncated: result.truncated,
          data_available: attendees.length > 0 && unknown.length === 0,
          unknown,
        });
      }

      if (op === "freebusy") {
        // A cached snapshot may cover a much wider range than this call asked
        // for; report only the requested window.
        const unknown = attendees.filter((email) => isUnknownAttendee(snapshot, email));
        const known = attendees.filter((email) => !unknown.includes(email));
        return textResult({
          range: {
            start: formatLocal(windowStart, timezone),
            end: formatLocal(windowEnd, timezone),
            timezone,
          },
          fetched_at: snapshot.metadata.fetched_at,
          stale: snapshotData.stale,
          refreshed: snapshotData.refreshed,
          // Digits for unknown mailboxes would read as "free all day".
          availability_view: buildAvailabilityView(snapshot.busy, known, {
            start: windowStart,
            end: windowEnd,
            intervalMinutes: 30,
          }),
          ...(unknown.length > 0 ? { unknown } : {}),
          attendees: attendees.map((email) => {
            const isUnknown = unknown.includes(email);
            return {
              email,
              // An unresolved mailbox is NOT a free person.
              ...(isUnknown ? { unknown: true } : {}),
              // False means "no data for this mailbox", not "free all day".
              data_available: hasScheduleData(snapshot, email),
              busy: summarizeBusyBlocks(
                clipBusy(busyFor(snapshot.busy, email), windowStart, windowEnd),
                timezone,
                20,
              ),
            };
          }),
          ...(unknown.length > 0
            ? {
                warning:
                  "unknown_mailbox: the service returned no schedule data for " +
                  `${unknown.join(", ")}; an empty busy list there is not evidence of availability`,
              }
            : {}),
        });
      }

      const durationMin = Math.min(480, Math.max(5, numberArg(args, "duration_min") ?? cfg.defaultDurationMin));
      const maxCandidates = Math.min(20, Math.max(1, numberArg(args, "max_candidates") ?? cfg.maxResults.slots));
      const workingHoursOnly = booleanArg(args, "working_hours_only") ?? true;
      const windowStartMs = Date.parse(windowStart);
      const windowEndMs = Date.parse(windowEnd);
      const nowMs = Date.now();

      // Never rank a slot by treating an unresolvable mailbox as free.
      const unknown = attendees.filter((email) => isUnknownAttendee(snapshot, email));
      const known = attendees.filter((email) => !unknown.includes(email));
      if (known.length === 0) {
        return errorResult(
          "unknown_attendees",
          `no schedule data for ${unknown.join(", ")}; the service could not resolve the mailbox`,
          "check the addresses for typos, or use people.search to resolve a name to a real address",
        );
      }

      // Service candidates first; the local search is the fallback and the
      // ranking authority when the service cannot answer. Both are constrained
      // to the caller's window, not the snapshot's (possibly much wider) range.
      let serviceCandidates: MeetingSuggestion[] = [];
      let serviceError: string | undefined;
      try {
        serviceCandidates = await findMeetingTimes(outlookClient(), {
          attendees: known,
          range: { start: windowStart, end: windowEnd },
          durationMin,
          maxCandidates,
        });
      } catch (error) {
        if (error instanceof ApiError) serviceError = error.toLine();
        else throw error;
      }

      const eligibleService = serviceCandidates.filter((suggestion) => {
        const slotStart = Date.parse(suggestion.start);
        const slotEnd = Date.parse(suggestion.end);
        if (!(slotStart >= windowStartMs && slotEnd <= windowEndMs)) return false;
        if (slotStart < nowMs) return false;
        if (workingHoursOnly && !withinWorkingHours(slotStart, slotEnd, timezone, cfg.workHours)) return false;
        return !withinBlockedTime(slotStart, slotEnd, timezone, cfg.blockedTimes);
      });

      const localAll = suggestSlots({
        busy: snapshot.busy,
        attendees: known,
        start: windowStart,
        end: windowEnd,
        durationMin,
        intervalMin: Math.min(120, Math.max(5, numberArg(args, "interval_min") ?? 15)),
        bufferMin: Math.min(120, Math.max(0, numberArg(args, "buffer_min") ?? 0)),
        workingHoursOnly,
        timezone,
        workHours: cfg.workHours,
        blockedTimes: cfg.blockedTimes,
        maxCandidates: 10_000,
      }).filter(
        (candidate) =>
          !workingHoursOnly ||
          !candidate.conflicts.some((conflict) => conflict.reason === "working_hours"),
      );
      const localTotal = localAll.length;
      const localCandidates = localAll.slice(0, maxCandidates);

      // Candidate honesty: a slot is verified only when every requested attendee
      // actually returned schedule data AND the service covered this candidate.
      const missingData = attendees.filter((email) => !hasScheduleData(snapshot, email));
      const allAttendeesHaveData = missingData.length === 0;
      const fromService = eligibleService.slice(0, maxCandidates).map((suggestion) => {
        const covered = known.every((email) =>
          suggestion.attendeeAvailability.some(
            (entry) => (entry.email ?? "").toLowerCase() === email.toLowerCase(),
          ),
        );
        const anyAvailability = suggestion.attendeeAvailability.length > 0;
        const verified = allAttendeesHaveData && covered;
        return {
          start: suggestion.start,
          end: suggestion.end,
          local: `${formatLocal(suggestion.start, timezone)} → ${formatLocal(suggestion.end, timezone)}`,
          // A count with no attendee data would be a fabricated "free" claim.
          free: anyAvailability ? availabilityLabel(suggestion, known) : ("unknown" as const),
          conflicts: anyAvailability
            ? conflictsOf(suggestion)
            : "findmeetingtimes returned no attendee availability for this candidate",
          data_available: verified,
          verified,
          source: "findmeetingtimes" as const,
        };
      });
      // A self-only request is the common "when am I free" shape. Outlook's
      // findMeetingTimes routinely returns candidates with no AttendeeAvailability
      // for the organizer alone, which used to leave every candidate
      // free:"unknown". When the service cannot vouch for the only attendee, the
      // local snapshot search is authoritative and still returns real slots.
      const serviceVouched = fromService.some((candidate) => candidate.free !== "unknown");
      const useLocal = fromService.length === 0 || (selfOnly && !serviceVouched);
      const candidates = useLocal
        ? localCandidates.map((candidate) => formatLocalCandidate(candidate, allAttendeesHaveData))
        : fromService;
      const cappedService = eligibleService.length > fromService.length;
      const dataWarnings: string[] = [];
      if (!allAttendeesHaveData) {
        dataWarnings.push(
          `no schedule data for ${missingData.join(", ")}; candidate availability is unverified`,
        );
      }
      if (!useLocal && fromService.some((candidate) => candidate.free === "unknown")) {
        dataWarnings.push(
          "findmeetingtimes returned candidates without per-attendee availability; free counts are not claimed",
        );
      }

      return textResult({
        range: {
          start: formatLocal(windowStart, timezone),
          end: formatLocal(windowEnd, timezone),
          timezone,
        },
        duration_min: durationMin,
        working_hours_only: workingHoursOnly,
        max_candidates: maxCandidates,
        returned: candidates.length,
        truncated: candidates.length > 0 && (useLocal ? localTotal > candidates.length : cappedService),
        ...(useLocal && candidates.length > 0 ? { total_available: localTotal } : {}),
        fetched_at: snapshot.metadata.fetched_at,
        stale: snapshotData.stale,
        ...(unknown.length > 0
          ? {
              excluded_unknown_attendees: unknown,
              warning:
                "unknown_mailbox: no schedule data for " +
                `${unknown.join(", ")}; they were excluded from ranking (not assumed free)`,
            }
          : {}),
        candidates,
        ...(dataWarnings.length > 0 ? { data_warning: dataWarnings.join("; ") } : {}),
        ...(candidates.length === 0
          ? {
              no_candidates: `no slot in ${formatLocal(windowStart, timezone)}..${formatLocal(
                windowEnd,
                timezone,
              )}${workingHoursOnly ? " within working hours" : ""}`,
            }
          : {}),
        ...(serviceError ? { service_error: serviceError } : {}),
        ...(useLocal && !serviceError
          ? {
              note:
                fromService.length === 0
                  ? "findmeetingtimes returned no candidates; local search used"
                  : "findmeetingtimes returned no per-attendee availability for the self-only request; local search used",
            }
          : {}),
        source: useLocal ? "local" : "findmeetingtimes",
      });
    } catch (error) {
      if (error instanceof ApiError) return errorResult(error.code, error.message, error.hint);
      if (error instanceof AuthRequiredError) return errorResult(error.code, error.message, error.hint);
      if (error instanceof SnapshotError) return errorResult(error.code, error.message, error.hint);
      throw error;
    }
  },
};

function formatLocalCandidate(candidate: SlotCandidate, dataAvailable: boolean) {
  return {
    start: candidate.start,
    end: candidate.end,
    local: `${candidate.local_start} → ${candidate.local_end}`,
    free: `${candidate.free_count}/${candidate.attendee_count}`,
    conflicts:
      candidate.conflicts.length === 0
        ? undefined
        : candidate.conflicts.map((conflict) => `${conflict.attendee}: ${conflict.reason}`).join("; "),
    data_available: dataAvailable,
    verified: dataAvailable,
    source: "local" as const,
  };
}

function availabilityLabel(suggestion: MeetingSuggestion, attendees: string[]): string {
  const free = suggestion.attendeeAvailability.filter((entry) => entry.availability === "free").length;
  return `${free}/${attendees.length}`;
}

function conflictsOf(suggestion: MeetingSuggestion): string | undefined {
  const busy = suggestion.attendeeAvailability.filter((entry) => entry.availability !== "free");
  if (busy.length === 0) return undefined;
  return busy.map((entry) => `${entry.email}: ${entry.availability}`).join("; ");
}
