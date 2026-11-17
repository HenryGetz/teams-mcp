/**
 * `availability.free`: explicit free windows over snapshot busy gaps.
 *
 * Covers working-hours clipping, busy subtraction, buffer padding, blocked
 * times, the returned cap/truncated flag, and the unknown-mailbox rule.
 */
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { availabilityTool } from "../src/tools/availability.js";
import { freeWindows } from "../src/availability/slots.js";
import type { BusyBlock } from "../src/snapshot/calendar.js";
import {
  clearSnapshotCache,
  SNAPSHOT_VERSION,
  writeSnapshot,
} from "../src/snapshot/calendar.js";
import type { FetchLike } from "../src/http/client.js";

const TZ = "America/Chicago";
const SELF = "self@example.com";
const OTHER = "other@example.com";
const TYPO = "typo@example.com";
const WORK_HOURS = { days: ["mon", "tue", "wed", "thu", "fri"] as const, start: "09:00", end: "17:00" };

/** Friday 2026-11-20: 09:00–17:00 CST == 15:00–23:00Z. */
const DAY_START = "2026-11-20T15:00:00.000Z";
const DAY_END = "2026-11-20T23:00:00.000Z";

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

/** Any network call is a test failure: free/check must be snapshot-only. */
function noNetwork() {
  const calls: string[] = [];
  const fetchImpl: FetchLike = async (url) => {
    calls.push(url);
    throw new Error(`unexpected network call: ${url}`);
  };
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
  return calls;
}

function busy(start: string, end: string, status: BusyBlock["status"] = "busy"): BusyBlock {
  return { start, end, status };
}

async function snapshot(busyMap: Record<string, BusyBlock[]>, unknown: string[] = [], attendees = [SELF]) {
  const fetchedAt = new Date().toISOString();
  await writeSnapshot({
    version: SNAPSHOT_VERSION,
    metadata: {
      fetched_at: fetchedAt,
      range_start: "2026-11-19T00:00:00.000Z",
      range_end: "2026-11-22T00:00:00.000Z",
      timezone: TZ,
      ttl_seconds: 3600,
      stale_after: new Date(Date.parse(fetchedAt) + 3_600_000).toISOString(),
      source: "outlook",
      attendees: [...new Set([...attendees, ...unknown])],
      unknown,
      data_available: attendees,
    },
    events: [],
    busy: Object.fromEntries(
      [...new Set([...attendees, ...unknown])].map((email) => [email, busyMap[email] ?? []]),
    ),
  });
}

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-free-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_SNAPSHOT_DIR = join(dir, "snapshots");
  process.env.TEAMS_MCP_TIMEZONE = TZ;
  resetConfigCache();
  clearSnapshotCache();
});

afterEach(async () => {
  setOutlookClient(undefined);
  clearSnapshotCache();
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("freeWindows", () => {
  const base = {
    busy: {
      [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z")],
    },
    attendees: [SELF],
    start: DAY_START,
    end: DAY_END,
    durationMin: 30,
    bufferMin: 0,
    workingHoursOnly: true,
    timezone: TZ,
    workHours: { days: [...WORK_HOURS.days], start: WORK_HOURS.start, end: WORK_HOURS.end },
    blockedTimes: [],
    maxSlots: 20,
  };

  it("inverts busy blocks into free windows inside working hours", () => {
    const result = freeWindows(base);
    expect(result.windows.map((window) => [window.local_start, window.local_end])).toEqual([
      ["2026-11-20 09:00", "2026-11-20 10:00"],
      ["2026-11-20 10:30", "2026-11-20 17:00"],
    ]);
    expect(result.truncated).toBe(false);
    expect(result.total).toBe(2);
  });

  it("pads busy blocks by buffer_min on both sides", () => {
    const result = freeWindows({ ...base, bufferMin: 15 });
    expect(result.windows.map((window) => [window.local_start, window.local_end])).toEqual([
      ["2026-11-20 09:00", "2026-11-20 09:45"],
      ["2026-11-20 10:45", "2026-11-20 17:00"],
    ]);
  });

  it("drops gaps shorter than duration_min", () => {
    const result = freeWindows({
      ...base,
      durationMin: 61,
      busy: {
        [SELF]: [busy("2026-11-20T15:30:00.000Z", "2026-11-20T16:00:00.000Z")],
      },
    });
    // 09:00–09:30 is only 30 min; 10:00–17:00 qualifies.
    expect(result.windows.map((window) => window.local_start)).toEqual(["2026-11-20 10:00"]);
  });

  it("subtracts recurring blocked times", () => {
    const result = freeWindows({
      ...base,
      blockedTimes: [{ start: "12:00", end: "13:00", reason: "lunch" }],
    });
    expect(result.windows.map((window) => [window.local_start, window.local_end])).toEqual([
      ["2026-11-20 09:00", "2026-11-20 10:00"],
      ["2026-11-20 10:30", "2026-11-20 12:00"],
      ["2026-11-20 13:00", "2026-11-20 17:00"],
    ]);
  });

  it("honours the slot cap and reports truncation", () => {
    const result = freeWindows({
      ...base,
      maxSlots: 1,
      blockedTimes: [{ start: "12:00", end: "13:00" }],
    });
    expect(result.windows).toHaveLength(1);
    expect(result.total).toBe(3);
    expect(result.truncated).toBe(true);
  });

  it("without working hours the whole range is considered", () => {
    const result = freeWindows({
      ...base,
      workingHoursOnly: false,
      start: "2026-11-20T13:00:00.000Z",
      end: "2026-11-20T23:00:00.000Z",
      busy: {
        [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z")],
      },
    });
    expect(result.windows.map((window) => window.local_start)).toEqual([
      "2026-11-20 07:00",
      "2026-11-20 10:30",
    ]);
  });
});

describe("availability.free op", () => {
  it("returns compact free slots for the signed-in user", async () => {
    await snapshot({ [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z")] });
    const calls = noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "free",
        attendees: [SELF],
        start: DAY_START,
        end: DAY_END,
        duration_min: 60,
      }),
    );
    expect(out.count).toBe(2);
    expect(out.truncated).toBe(false);
    expect(out.data_available).toBe(true);
    expect(out.unknown).toEqual([]);
    expect((out.slots as Array<Record<string, unknown>>)[0]).toEqual({
      start: "2026-11-20T15:00:00.000Z",
      end: "2026-11-20T16:00:00.000Z",
      local: "2026-11-20 09:00 → 2026-11-20 10:00",
    });
    expect(calls).toEqual([]);
  });

  it("intersects several attendees' busy blocks", async () => {
    await snapshot(
      {
        [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z")],
        [OTHER]: [busy("2026-11-20T18:00:00.000Z", "2026-11-20T19:00:00.000Z")],
      },
      [],
      [SELF, OTHER],
    );
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "free",
        attendees: [SELF, OTHER],
        start: DAY_START,
        end: DAY_END,
        duration_min: 30,
      }),
    );
    expect((out.slots as Array<Record<string, unknown>>).map((slot) => slot.local)).toEqual([
      "2026-11-20 09:00 → 2026-11-20 10:00",
      "2026-11-20 10:30 → 2026-11-20 12:00",
      "2026-11-20 13:00 → 2026-11-20 17:00",
    ]);
  });

  it("never treats an unknown attendee as free", async () => {
    await snapshot({ [SELF]: [] }, [TYPO], [SELF]);
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "free",
        attendees: [SELF, TYPO],
        start: DAY_START,
        end: DAY_END,
      }),
    );
    expect(out.data_available).toBe(false);
    expect(out.unknown).toEqual([TYPO]);
    // The signed-in user still has verified windows; the flag says they are
    // not a complete answer for the whole attendee set.
    expect((out.slots as unknown[]).length).toBeGreaterThan(0);
  });

  it("returns no windows when every attendee is unknown", async () => {
    await snapshot({}, [TYPO], []);
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "free",
        attendees: [TYPO],
        start: DAY_START,
        end: DAY_END,
      }),
    );
    expect(out).toMatchObject({ count: 0, truncated: false, data_available: false, unknown: [TYPO] });
  });
});
