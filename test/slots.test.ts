import { describe, expect, it } from "vitest";
import type { WorkHours } from "../src/config.js";
import { buildAvailabilityView, suggestSlots, summarizeBusyBlocks } from "../src/availability/slots.js";
import type { BusyBlock } from "../src/snapshot/calendar.js";

const WORK_HOURS: WorkHours = { days: ["mon", "tue", "wed", "thu", "fri"], start: "09:00", end: "17:00" };
const ZONE = "America/Chicago";
// 2026-11-16 is a Monday. 15:00Z = 09:00 CST.
const MONDAY_09 = "2026-11-16T15:00:00.000Z";

const base = {
  intervalMin: 15,
  bufferMin: 0,
  workingHoursOnly: true,
  workHours: WORK_HOURS,
  blockedTimes: [],
  maxCandidates: 5,
  timezone: ZONE,
  start: "2026-11-16T15:00:00.000Z",
  end: "2026-11-16T22:00:00.000Z",
  // Past is irrelevant for these fixtures.
  now: Date.parse("2026-11-01T00:00:00.000Z"),
};

function busy(blocks: BusyBlock[]): Record<string, BusyBlock[]> {
  return { "alice@example.com": blocks };
}

describe("suggestSlots", () => {
  it("returns the earliest free working-hours slot first", () => {
    const slots = suggestSlots({ ...base, busy: {}, attendees: ["alice@example.com"], durationMin: 30 });
    expect(slots[0]?.start).toBe(MONDAY_09);
    expect(slots[0]?.end).toBe("2026-11-16T15:30:00.000Z");
    expect(slots[0]?.local_start).toBe("2026-11-16 09:00");
    expect(slots[0]?.conflicts).toEqual([]);
    expect(slots[0]?.free_count).toBe(1);
  });

  it("skips busy windows and reports the conflict", () => {
    const slots = suggestSlots({
      ...base,
      busy: busy([{ start: MONDAY_09, end: "2026-11-16T16:00:00.000Z", status: "busy" }]),
      attendees: ["alice@example.com"],
      durationMin: 60,
      maxCandidates: 2,
    });
    expect(slots[0]?.start).toBe("2026-11-16T16:00:00.000Z");
    const conflicted = slots.find((slot) => slot.start === MONDAY_09);
    expect(conflicted).toBeUndefined();
    expect(slots.length).toBe(2);
  });

  it("ranks fully free slots above slots with conflicts", () => {
    const slots = suggestSlots({
      ...base,
      busy: busy([{ start: "2026-11-16T15:00:00.000Z", end: "2026-11-16T16:00:00.000Z", status: "busy" }]),
      attendees: ["alice@example.com"],
      durationMin: 30,
      maxCandidates: 3,
    });
    expect(slots[0]?.conflicts).toEqual([]);
    expect(slots[0]?.start).toBe("2026-11-16T16:00:00.000Z");
  });

  it("honours working hours", () => {
    const slots = suggestSlots({
      ...base,
      busy: {},
      attendees: ["alice@example.com"],
      durationMin: 30,
      start: "2026-11-16T13:00:00.000Z",
      end: "2026-11-16T22:00:00.000Z",
    });
    expect(slots[0]?.start).toBe(MONDAY_09);
    expect(slots.every((slot) => slot.conflicts.every((c) => c.reason !== "working_hours"))).toBe(true);
  });

  it("treats blocked times as conflicts", () => {
    const slots = suggestSlots({
      ...base,
      busy: {},
      attendees: ["alice@example.com"],
      durationMin: 30,
      blockedTimes: [{ start: "09:00", end: "10:00", reason: "lunch" }],
      maxCandidates: 20,
    });
    const first = slots.find((slot) => slot.local_start === "2026-11-16 09:00");
    expect(first?.conflicts).toContainEqual({ attendee: "(blocked)", reason: "blocked" });
    expect(slots[0]?.conflicts).toEqual([]);
  });

  it("applies the buffer around candidates", () => {
    const slots = suggestSlots({
      ...base,
      busy: busy([{ start: "2026-11-16T15:40:00.000Z", end: "2026-11-16T16:00:00.000Z", status: "busy" }]),
      attendees: ["alice@example.com"],
      durationMin: 30,
      bufferMin: 15,
    });
    expect(slots[0]?.conflicts).toEqual([]);
    expect(slots[0]?.start).toBe("2026-11-16T16:15:00.000Z");
  });

  it("never returns overlapping candidates", () => {
    const slots = suggestSlots({
      ...base,
      busy: {},
      attendees: ["alice@example.com"],
      durationMin: 60,
      maxCandidates: 4,
    });
    for (let index = 1; index < slots.length; index += 1) {
      expect(Date.parse(slots[index]!.start)).toBeGreaterThanOrEqual(Date.parse(slots[index - 1]!.end));
    }
  });

  it("deduplicates repeated attendees", () => {
    const slots = suggestSlots({
      ...base,
      busy: busy([{ start: MONDAY_09, end: "2026-11-16T16:00:00.000Z", status: "busy" }]),
      attendees: ["alice@example.com"],
      durationMin: 30,
    });
    expect(slots[0]?.attendee_count).toBe(1);
  });
});

describe("buildAvailabilityView", () => {
  it("emits one digit per interval with the Graph legend", () => {
    const view = buildAvailabilityView(
      busy([{ start: MONDAY_09, end: "2026-11-16T15:30:00.000Z", status: "busy" }]),
      ["alice@example.com"],
      { start: MONDAY_09, end: "2026-11-16T16:00:00.000Z", intervalMinutes: 30 },
    );
    expect(view["alice@example.com"]).toBe("20");
  });

  it("maps oof and tentative", () => {
    const view = buildAvailabilityView(
      busy([
        { start: MONDAY_09, end: "2026-11-16T15:30:00.000Z", status: "tentative" },
        { start: "2026-11-16T15:30:00.000Z", end: "2026-11-16T16:00:00.000Z", status: "oof" },
      ]),
      ["alice@example.com"],
      { start: MONDAY_09, end: "2026-11-16T16:00:00.000Z", intervalMinutes: 30 },
    );
    expect(view["alice@example.com"]).toBe("13");
  });
});

describe("summarizeBusyBlocks", () => {
  it("renders local times and marks truncation", () => {
    const blocks: BusyBlock[] = Array.from({ length: 3 }, (_, index) => ({
      start: new Date(Date.parse(MONDAY_09) + index * 2 * 3_600_000).toISOString(),
      end: new Date(Date.parse(MONDAY_09) + (index * 2 + 0.5) * 3_600_000).toISOString(),
      status: "busy" as const,
    }));
    const rows = summarizeBusyBlocks(blocks, ZONE, 2);
    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual({ start: "2026-11-16 09:00", end: "2026-11-16 09:30", status: "busy" });
    expect(rows[2]).toEqual({ omitted: 1 });
  });
});
