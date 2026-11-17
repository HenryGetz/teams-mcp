import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { GraphClient, setGraphClient } from "../src/graph/client.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { clearSnapshotCache, SNAPSHOT_VERSION, writeSnapshot } from "../src/snapshot/calendar.js";
import { availabilityTool } from "../src/tools/availability.js";
import { formatLocal } from "../src/time.js";

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

const errorText = (result: CallToolResult): string =>
  result.content[0]?.type === "text" ? result.content[0].text : "";

interface Call {
  api: "outlook" | "graph";
  method: string;
  url: string;
  body: Record<string, unknown> | undefined;
}

function harness(responses: { outlook?: Array<() => Response>; graph?: Array<() => Response> }) {
  const calls: Call[] = [];
  let outlookIndex = 0;
  let graphIndex = 0;
  const capture = (api: "outlook" | "graph") => async (url: string, init?: RequestInit) => {
    calls.push({
      api,
      method: init?.method ?? "GET",
      url,
      body: typeof init?.body === "string" ? (JSON.parse(init.body) as Record<string, unknown>) : undefined,
    });
    const queue = responses[api] ?? [];
    const index = api === "outlook" ? outlookIndex++ : graphIndex++;
    const factory = queue[index];
    if (!factory) throw new Error(`unexpected ${api} call ${url}`);
    return factory();
  };
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl: capture("outlook"), maxAttempts: 1 }),
  );
  setGraphClient(
    new GraphClient({ tokenProvider: async () => ({ token: "g" }), fetchImpl: capture("graph"), maxAttempts: 1 }),
  );
  return calls;
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

const ATTENDEE = "a@example.com";
const TZ = "America/Chicago";

function suggestion(startIso: string, endIso: string) {
  return {
    Confidence: 100,
    MeetingTimeSlot: {
      Start: { DateTime: startIso.replace(/\.\d+Z$/, "").replace(/Z$/, ""), TimeZone: "UTC" },
      End: { DateTime: endIso.replace(/\.\d+Z$/, "").replace(/Z$/, ""), TimeZone: "UTC" },
    },
    AttendeeAvailability: [
      { Attendee: { EmailAddress: { Address: ATTENDEE } }, Availability: "Free" },
    ],
  };
}

/** Tuesday 2026-11-17 09:00–12:00 and Wednesday 10:00, in America/Chicago. */
const DAY1_START = "2026-11-17T15:00:00.000Z"; // 09:00 CST
const DAY1_END = "2026-11-17T18:00:00.000Z"; // 12:00 CST
const DAY2_START = "2026-11-18T16:00:00.000Z"; // 10:00 CST
const DAY2_END = "2026-11-18T19:00:00.000Z";
/**
 * Fixtures are pinned to a specific calendar day, so the clock is pinned before
 * that day starts: `suggest` correctly drops slots in the past, and without this
 * the expected candidate counts would shrink as the real day advances.
 */
const FIXED_NOW = "2026-11-17T14:00:00.000Z";
const IN_DAY1 = suggestion("2026-11-17T15:30:00Z", "2026-11-17T16:00:00Z");
const IN_DAY2 = suggestion("2026-11-18T16:00:00Z", "2026-11-18T16:30:00Z");
const OUTSIDE = suggestion("2026-11-21T15:00:00Z", "2026-11-21T15:30:00Z");
const BEFORE_HOURS = suggestion("2026-11-17T13:00:00Z", "2026-11-17T13:30:00Z"); // 07:00 CST

let dir: string;
const saved = { ...process.env };

async function wideSnapshot(attendees: string[] = [ATTENDEE]): Promise<void> {
  const now = Date.now();
  await writeSnapshot({
    version: SNAPSHOT_VERSION,
    metadata: {
      fetched_at: new Date(now).toISOString(),
      range_start: new Date(now - 3_600_000).toISOString(),
      range_end: new Date(now + 7 * 86_400_000).toISOString(),
      timezone: TZ,
      ttl_seconds: 1800,
      stale_after: new Date(now + 1_800_000).toISOString(),
      source: "outlook",
      attendees,
      unknown: [],
    },
    events: [],
    busy: Object.fromEntries(attendees.map((email) => [email, []])),
  });
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-availability-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_SNAPSHOT_DIR = join(dir, "snapshots");
  resetConfigCache();
  clearSnapshotCache();
  // Only Date is faked; timers stay real for the async handler.
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date(FIXED_NOW));
});

afterEach(async () => {
  vi.useRealTimers();
  setOutlookClient(undefined);
  setGraphClient(undefined);
  clearSnapshotCache();
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("availability.freebusy window", () => {
  it("REG: clips a cached wide snapshot to the requested window", async () => {
    await wideSnapshot();
    const result = await availabilityTool.handle({
      op: "freebusy",
      attendees: [ATTENDEE],
      start: DAY1_START,
      end: DAY1_END,
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.range).toEqual({
      start: formatLocal(DAY1_START, TZ),
      end: formatLocal(DAY1_END, TZ),
      timezone: TZ,
    });
    expect(out.availability_view).toEqual({ [ATTENDEE]: "000000" });
  });
});

describe("availability.suggest window", () => {
  it("REG: drives findMeetingTimes and local search from the requested window", async () => {
    await wideSnapshot();
    const calls = harness({ outlook: [json({ MeetingTimeSuggestions: [IN_DAY2, OUTSIDE, IN_DAY1, BEFORE_HOURS] })] });
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: [ATTENDEE],
      start: DAY1_START,
      end: DAY1_END,
      working_hours_only: true,
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.range).toEqual({
      start: formatLocal(DAY1_START, TZ),
      end: formatLocal(DAY1_END, TZ),
      timezone: TZ,
    });
    const candidates = out.candidates as Array<Record<string, unknown>>;
    expect(candidates).toHaveLength(1);
    expect(candidates[0]?.start).toBe("2026-11-17T15:30:00.000Z");
    expect(out).toMatchObject({ max_candidates: 5, returned: 1, truncated: false });
    expect(out.total_available).toBeUndefined();

    const request = calls.find((call) => call.url.includes("findmeetingtimes"));
    const timeslot = (request?.body?.TimeConstraint as Record<string, unknown>)?.Timeslots as Array<
      Record<string, unknown>
    >;
    expect((timeslot?.[0]?.Start as Record<string, unknown>)?.DateTime).toBe("2026-11-17T15:00:00");
    expect((timeslot?.[0]?.End as Record<string, unknown>)?.DateTime).toBe("2026-11-17T18:00:00");
    expect(request?.body?.MaxCandidates).toBe("5");
  });

  it("REG: per-day calls return only that day, never the whole cached week", async () => {
    await wideSnapshot();
    harness({
      outlook: [
        json({ MeetingTimeSuggestions: [IN_DAY1, IN_DAY2] }),
        json({ MeetingTimeSuggestions: [IN_DAY1, IN_DAY2] }),
      ],
    });
    const day1 = payload(
      await availabilityTool.handle({ op: "suggest", attendees: [ATTENDEE], start: DAY1_START, end: DAY1_END }),
    );
    const day2 = payload(
      await availabilityTool.handle({ op: "suggest", attendees: [ATTENDEE], start: DAY2_START, end: DAY2_END }),
    );
    const rows = (value: Record<string, unknown>) => value.candidates as Array<Record<string, unknown>>;
    expect(rows(day1).map((row) => row.start)).toEqual(["2026-11-17T15:30:00.000Z"]);
    expect(rows(day2).map((row) => row.start)).toEqual(["2026-11-18T16:00:00.000Z"]);
    expect(day1).toMatchObject({ returned: 1, truncated: false, max_candidates: 5 });
    expect(day2).toMatchObject({ returned: 1, truncated: false, max_candidates: 5 });
    expect((day1.range as Record<string, unknown>).start).not.toBe((day2.range as Record<string, unknown>).start);
  });

  it("falls back to local search inside the window when findMeetingTimes fails", async () => {
    // The fixture day is 2026-11-17 09:00 CST. Pin the clock before that so the
    // past-slot filter cannot change the expected candidate count as the day
    // advances. Only Date is faked; timers stay real for the async handler.
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(new Date("2026-11-17T14:00:00.000Z"));
    try {
      await wideSnapshot();
    harness({
      outlook: [
        json({ error: { code: "MockFailure", message: "no service" } }, 400),
        json({ error: { code: "MockFailure", message: "no service" } }, 400),
      ],
    });
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: [ATTENDEE],
      start: DAY1_START,
      end: DAY1_END,
      duration_min: 30,
      interval_min: 30,
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.source).toBe("local");
    expect(out.service_error).toContain("MockFailure");
    expect(out).toMatchObject({
      max_candidates: 5,
      returned: 5,
      total_available: 6,
      truncated: true,
    });
    const candidates = out.candidates as Array<Record<string, unknown>>;
    expect(candidates.length).toBe(5);
    for (const candidate of candidates) {
      expect(Date.parse(candidate.start as string)).toBeGreaterThanOrEqual(Date.parse(DAY1_START));
      expect(Date.parse(candidate.end as string)).toBeLessThanOrEqual(Date.parse(DAY1_END));
      expect(formatLocal(candidate.start as string, TZ)).toMatch(/ 09:| 1[01]:/);
    }

    const wider = payload(
      await availabilityTool.handle({
        op: "suggest",
        attendees: [ATTENDEE],
        start: DAY1_START,
        end: DAY1_END,
        duration_min: 30,
        interval_min: 30,
        max_candidates: 20,
      }),
    );
    expect(wider).toMatchObject({
      max_candidates: 20,
      returned: 6,
      total_available: 6,
      truncated: false,
    });
    } finally {
      vi.useRealTimers();
    }
  });

  it("clamps oversized max_candidates to 20 instead of rejecting", async () => {
    expect(availabilityTool.shape.max_candidates?.safeParse(200).success).toBe(true);
    const many = Array.from({ length: 25 }, (_, index) =>
      suggestion(
        new Date(Date.parse(DAY1_START) + index * 5 * 60_000).toISOString(),
        new Date(Date.parse(DAY1_START) + (index * 5 + 5) * 60_000).toISOString(),
      ),
    );
    await wideSnapshot();
    const calls = harness({ outlook: [json({ MeetingTimeSuggestions: many })] });
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: [ATTENDEE],
      start: DAY1_START,
      end: DAY1_END,
      duration_min: 5,
      interval_min: 5,
      max_candidates: 200,
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect((out.candidates as unknown[]).length).toBe(20);
    expect(out).toMatchObject({ max_candidates: 20, returned: 20, truncated: true });
    expect(out.total_available).toBeUndefined();
    const request = calls.find((call) => call.url.includes("findmeetingtimes"));
    expect(request?.body?.MaxCandidates).toBe("20");
  });

  it("REGRESSION: resolves attendee names and refuses ambiguity before any API call", async () => {
    const joel = {
      id: "id-1",
      displayName: "Student, Test",
      mail: "joel@example.com",
      userPrincipalName: "joel@example.com",
    };
    const twin = { ...joel, id: "id-2", displayName: "Student, Test (SW)", mail: "joel2@example.com" };
    const calls = harness({ graph: [json({ value: [joel, twin] })] });
    const ambiguous = await availabilityTool.handle({
      op: "suggest",
      attendees: ["Test Student"],
    });
    expect(ambiguous.isError).toBe(true);
    expect(JSON.parse(errorText(ambiguous))).toMatchObject({
      error: "ambiguous_recipient",
      needs_disambiguation: true,
    });
    expect(calls.some((call) => call.api === "outlook")).toBe(false);
  });

  it("applies the same resolver to freebusy and snapshot before any REST call", async () => {
    const joel = {
      id: "id-1",
      displayName: "Student, Test",
      mail: "joel@example.com",
      userPrincipalName: "joel@example.com",
    };
    const twin = { ...joel, id: "id-2", displayName: "Student, Test (SW)", mail: "joel2@example.com" };
    for (const op of ["freebusy", "snapshot"] as const) {
      const calls = harness({ graph: [json({ value: [joel, twin] })] });
      const result = await availabilityTool.handle({ op, attendees: ["Test Student"] });
      expect(result.isError).toBe(true);
      expect(JSON.parse(errorText(result))).toMatchObject({
        error: "ambiguous_recipient",
        needs_disambiguation: true,
      });
      expect(calls.some((call) => call.api === "outlook")).toBe(false);
    }
  });

  it("uses the single directory match for the scheduling request", async () => {
    await wideSnapshot(["joel@example.com"]);
    harness({
      graph: [
        json({
          value: [
            {
              id: "id-1",
              displayName: "Student, Test",
              mail: "joel@example.com",
              userPrincipalName: "joel@example.com",
            },
          ],
        }),
      ],
      outlook: [json({ MeetingTimeSuggestions: [IN_DAY1] })],
    });
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: ["Test Student"],
      start: DAY1_START,
      end: DAY1_END,
    });
    expect(result.isError).toBeUndefined();
    expect((payload(result).range as Record<string, unknown>).start).toBe(formatLocal(DAY1_START, TZ));
  });
});
