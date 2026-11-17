/**
 * Availability honesty: a mailbox with no returned data is never reported as
 * free, and service candidates that carry no per-attendee availability do not
 * get a fabricated "n/n free" count.
 */
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { availabilityTool } from "../src/tools/availability.js";
import {
  buildSnapshot,
  clearSnapshotCache,
  SNAPSHOT_VERSION,
  writeSnapshot,
} from "../src/snapshot/calendar.js";
import type { FetchLike } from "../src/http/client.js";

const START = "2026-11-17T15:00:00.000Z";
const END = "2026-11-17T18:00:00.000Z";
/**
 * Fixtures are pinned to 2026-11-17 09:00 CST, so freeze the clock before that
 * day starts; otherwise `suggest` drops the fixture slots as "past" and the
 * candidate assertions change with the wall clock.
 */
const FIXED_NOW = "2026-11-17T14:00:00.000Z";
const KNOWN = "known@example.com";
const TYPO = "typo@example.com";
const TZ = "America/Chicago";

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

function harness(responses: Array<() => Response>) {
  let index = 0;
  const fetchImpl: FetchLike = async () => {
    const factory = responses[Math.min(index, responses.length - 1)];
    index += 1;
    if (!factory) throw new Error("no response queued");
    return factory();
  };
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

function suggestion(startIso: string, endIso: string, attendeeAvailability: unknown[]) {
  return {
    Confidence: 100,
    MeetingTimeSlot: {
      Start: { DateTime: startIso.replace(/\.\d+Z$/, "").replace(/Z$/, ""), TimeZone: "UTC" },
      End: { DateTime: endIso.replace(/\.\d+Z$/, "").replace(/Z$/, ""), TimeZone: "UTC" },
    },
    AttendeeAvailability: attendeeAvailability,
  };
}

const SLOT = suggestion("2026-11-17T15:30:00Z", "2026-11-17T16:00:00Z", [
  { Attendee: { EmailAddress: { Address: KNOWN } }, Availability: "Free" },
]);

let dir: string;
const saved = { ...process.env };

async function snapshot(attendees: string[], unknown: string[], dataAvailable: string[]): Promise<void> {
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
      unknown,
      data_available: dataAvailable,
    },
    events: [],
    busy: Object.fromEntries(attendees.map((email) => [email, []])),
  });
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-availability-data-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_SNAPSHOT_DIR = join(dir, "snapshots");
  process.env.TEAMS_MCP_TIMEZONE = TZ;
  resetConfigCache();
  clearSnapshotCache();
  // Only Date is faked; timers stay real for the async handler.
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date(FIXED_NOW));
});

afterEach(async () => {
  vi.useRealTimers();
  setOutlookClient(undefined);
  clearSnapshotCache();
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("freebusy data_available", () => {
  it("marks a mailbox with data true and an unresolved mailbox false", async () => {
    await snapshot([KNOWN, TYPO], [TYPO], [KNOWN]);
    const out = payload(
      await availabilityTool.handle({ op: "freebusy", attendees: [KNOWN, TYPO], start: START, end: END }),
    );
    const rows = out.attendees as Array<Record<string, unknown>>;
    expect(rows.find((row) => row.email === KNOWN)).toMatchObject({ data_available: true });
    expect(rows.find((row) => row.email === TYPO)).toMatchObject({
      unknown: true,
      data_available: false,
      busy: [],
    });
  });
});

describe("suggest verified candidates", () => {
  it("marks a local candidate verified when every attendee has data", async () => {
    await snapshot([KNOWN], [], [KNOWN]);
    harness([json({ error: { code: "MockFailure", message: "no service" } }, 400)]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        attendees: [KNOWN],
        start: START,
        end: END,
        duration_min: 30,
        interval_min: 30,
        working_hours_only: false,
      }),
    );
    expect(out.source).toBe("local");
    const candidate = (out.candidates as Array<Record<string, unknown>>)[0];
    expect(candidate).toMatchObject({ data_available: true, verified: true });
    expect(out.data_warning).toBeUndefined();
  });

  it("leaves candidates unverified when an attendee has no data", async () => {
    await snapshot([KNOWN, TYPO], [TYPO], [KNOWN]);
    harness([json({ error: { code: "MockFailure", message: "no service" } }, 400)]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        attendees: [KNOWN, TYPO],
        start: START,
        end: END,
        duration_min: 30,
        interval_min: 30,
        working_hours_only: false,
      }),
    );
    expect(out.excluded_unknown_attendees).toEqual([TYPO]);
    expect(String(out.data_warning)).toContain("no schedule data");
    const candidate = (out.candidates as Array<Record<string, unknown>>)[0];
    expect(candidate).toMatchObject({ data_available: false, verified: false });
  });

  it("does not claim a free count when the service returned no attendee data", async () => {
    await snapshot([KNOWN], [], [KNOWN]);
    harness([json({ MeetingTimeSuggestions: [suggestion("2026-11-17T15:30:00Z", "2026-11-17T16:00:00Z", [])] })]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        attendees: [KNOWN],
        start: START,
        end: END,
        duration_min: 30,
        working_hours_only: false,
      }),
    );
    const candidate = (out.candidates as Array<Record<string, unknown>>)[0];
    expect(candidate?.free).toBe("unknown");
    expect(candidate?.free).not.toMatch(/^\d+\/\d+$/);
    expect(candidate).toMatchObject({ data_available: false, verified: false });
    expect(String(out.data_warning)).toContain("without per-attendee availability");
  });

  it("marks a service candidate verified when attendee availability covers everyone", async () => {
    await snapshot([KNOWN], [], [KNOWN]);
    harness([json({ MeetingTimeSuggestions: [SLOT] })]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        attendees: [KNOWN],
        start: START,
        end: END,
        duration_min: 30,
        working_hours_only: false,
      }),
    );
    expect(out.source).toBe("findmeetingtimes");
    expect((out.candidates as Array<Record<string, unknown>>)[0]).toMatchObject({
      free: "1/1",
      data_available: true,
      verified: true,
    });
    expect(out.data_warning).toBeUndefined();
  });
});

describe("snapshot data_available metadata", () => {
  it("records only mailboxes that returned real schedule data", async () => {
    const built = await buildSnapshot({
      source: {
        async ownEvents() {
          return [];
        },
        async schedules(attendees) {
          return attendees.map((email) => ({
            email,
            blocks: [],
            ...(email === TYPO ? { data_available: false, unknown: true as const } : { data_available: true }),
          }));
        },
      },
      attendees: [KNOWN, TYPO],
      range: { start: START, end: END },
    });
    expect(built.metadata.data_available).toEqual([KNOWN]);
    expect(built.metadata.unknown).toEqual([TYPO]);
  });
});
