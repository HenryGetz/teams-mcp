/**
 * DST regressions (America/Chicago, 2027): 03-14 spring forward (02:00 -> 03:00)
 * and 11-07 fall back (02:00 -> 01:00).
 *
 * The bugs these lock down: a wall clock just after the spring-forward was
 * resolved with the pre-gap offset (one hour late), local midnight was derived
 * from a noon probe (one hour off on both transition days), `defaultRange`
 * added nine absolute hours instead of landing on 09:00 local, and
 * `parseInstant` silently slid a nonexistent 02:30 to 03:30.
 */
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { buildEventBody, eventTime } from "../src/outlook/calendar.js";
import { calendarTool } from "../src/tools/calendar.js";
import { suggestSlots, type SuggestOptions } from "../src/availability/slots.js";
import { defaultRange } from "../src/snapshot/calendar.js";
import {
  formatLocal,
  localMidnightOfDate,
  parseInstant,
  parseInstantDetailed,
  startOfLocalDay,
  wallClock,
  zoneOffsetMinutes,
} from "../src/time.js";
import type { FetchLike } from "../src/http/client.js";

const TZ = "America/Chicago";

const iso = (ms: number | undefined): string | undefined =>
  ms === undefined ? undefined : new Date(ms).toISOString();

describe("DST — spring forward (2027-03-14)", () => {
  it("resolves wall clocks after the gap with the post-transition offset", () => {
    // 03:00 CDT == 08:00Z. The old one-pass probe produced 09:00Z (an hour late).
    expect(iso(parseInstant("2027-03-14T03:00", TZ))).toBe("2027-03-14T08:00:00.000Z");
    expect(iso(parseInstant("2027-03-14T03:30", TZ))).toBe("2027-03-14T08:30:00.000Z");
    expect(iso(parseInstant("2027-03-14T03:59", TZ))).toBe("2027-03-14T08:59:00.000Z");
    // Still-CST side of the boundary.
    expect(iso(parseInstant("2027-03-14T01:59", TZ))).toBe("2027-03-14T07:59:00.000Z");
    // Well clear of the transition.
    expect(iso(parseInstant("2027-03-14T10:00", TZ))).toBe("2027-03-14T15:00:00.000Z");
  });

  it("reports a nonexistent wall clock instead of sliding it an hour", () => {
    expect(parseInstantDetailed("2027-03-14T02:30", TZ)).toEqual({ error: "nonexistent" });
    expect(parseInstant("2027-03-14T02:30", TZ)).toBeUndefined();
    expect(parseInstantDetailed("2027-03-14T02:00", TZ)).toEqual({ error: "nonexistent" });
    expect(parseInstantDetailed("2027-03-14T03:00", TZ)).toEqual({
      instant: Date.parse("2027-03-14T08:00:00Z"),
    });
  });

  it("puts midnight at CST on the transition day", () => {
    expect(iso(localMidnightOfDate("2027-03-14", TZ))).toBe("2027-03-14T06:00:00.000Z");
    expect(formatLocal(localMidnightOfDate("2027-03-14", TZ), TZ)).toBe("2027-03-14 00:00");
    expect(iso(localMidnightOfDate("2027-03-15", TZ))).toBe("2027-03-15T05:00:00.000Z");
    // The old noon-probe returned 05:00Z, i.e. 23:00 the previous day.
    expect(startOfLocalDay(Date.parse("2027-03-14T12:00:00Z"), TZ)).toBe(
      Date.parse("2027-03-14T06:00:00Z"),
    );
  });

  it("keeps the default scheduling window at 09:00 local", () => {
    const range = defaultRange(TZ, Date.parse("2027-03-14T18:00:00Z"));
    expect(formatLocal(range.start, TZ)).toBe("2027-03-14 09:00");
    expect(range.start).toBe("2027-03-14T14:00:00.000Z");
    expect(formatLocal(range.end, TZ)).toBe("2027-03-21 17:00");
  });

  it("offsets reflect the transition instant", () => {
    expect(zoneOffsetMinutes(Date.parse("2027-03-14T07:59:00Z"), TZ)).toBe(-360);
    expect(zoneOffsetMinutes(Date.parse("2027-03-14T08:00:00Z"), TZ)).toBe(-300);
  });
});

describe("DST — fall back (2027-11-07)", () => {
  it("puts midnight at CDT on the transition day", () => {
    expect(iso(localMidnightOfDate("2027-11-07", TZ))).toBe("2027-11-07T05:00:00.000Z");
    expect(formatLocal(localMidnightOfDate("2027-11-07", TZ), TZ)).toBe("2027-11-07 00:00");
    expect(iso(localMidnightOfDate("2027-11-08", TZ))).toBe("2027-11-08T06:00:00.000Z");
  });

  it("resolves an ambiguous wall clock to its first occurrence and round-trips", () => {
    const first = parseInstant("2027-11-07T01:30", TZ);
    expect(iso(first)).toBe("2027-11-07T06:30:00.000Z");
    expect(wallClock(first ?? 0, TZ)).toBe("2027-11-07T01:30:00");
    // The second 01:30 (CST) is reachable with an explicit offset.
    expect(iso(parseInstant("2027-11-07T01:30:00-06:00", TZ))).toBe("2027-11-07T07:30:00.000Z");
    expect(iso(parseInstant("2027-11-07T02:30", TZ))).toBe("2027-11-07T08:30:00.000Z");
  });

  it("keeps the default scheduling window at 09:00/17:00 local", () => {
    const range = defaultRange(TZ, Date.parse("2027-11-07T12:00:00Z"));
    expect(formatLocal(range.start, TZ)).toBe("2027-11-07 09:00");
    expect(range.start).toBe("2027-11-07T15:00:00.000Z");
    expect(formatLocal(range.end, TZ)).toBe("2027-11-14 17:00");
    expect(range.end).toBe("2027-11-14T23:00:00.000Z");
  });
});

describe("DST — slot search spanning the boundary", () => {
  const base: SuggestOptions = {
    busy: {},
    attendees: ["me@example.com"],
    start: "",
    end: "",
    durationMin: 30,
    intervalMin: 15,
    bufferMin: 0,
    workingHoursOnly: false,
    timezone: TZ,
    workHours: { days: ["sun", "mon", "tue", "wed", "thu", "fri", "sat"], start: "09:00", end: "17:00" },
    blockedTimes: [],
    maxCandidates: 100,
    now: 0,
  };

  it("steps over the spring-forward gap without producing a 02:xx slot", () => {
    const slots = suggestSlots({
      ...base,
      start: "2027-03-14T07:00:00.000Z", // 01:00 CST
      end: "2027-03-14T10:00:00.000Z", // 05:00 CDT
    });
    const locals = slots.map((slot) => slot.local_start);
    // No candidate can ever be labelled 02:xx — that wall clock does not exist.
    expect(locals.some((local) => local.includes(" 02:"))).toBe(false);
    // The last pre-gap slot is 01:30 CST (07:30Z); the next is 03:00 CDT (08:00Z).
    expect(slots.find((slot) => slot.local_start === "2027-03-14 01:30")?.start).toBe(
      "2027-03-14T07:30:00.000Z",
    );
    expect(slots.find((slot) => slot.local_start === "2027-03-14 03:00")?.start).toBe(
      "2027-03-14T08:00:00.000Z",
    );
    expect(locals.slice(0, 3)).toEqual([
      "2027-03-14 01:00",
      "2027-03-14 01:30",
      "2027-03-14 03:00",
    ]);
  });

  it("offers the repeated fall-back hour twice, as two real instants", () => {
    const slots = suggestSlots({
      ...base,
      start: "2027-11-07T05:00:00.000Z", // 00:00 CDT
      end: "2027-11-07T09:00:00.000Z", // 03:00 CST
    });
    const repeated = slots.filter((slot) => slot.local_start === "2027-11-07 01:30");
    expect(repeated.map((slot) => slot.start)).toEqual([
      "2027-11-07T06:30:00.000Z",
      "2027-11-07T07:30:00.000Z",
    ]);
    // Four real hours on a 30-minute duration / 15-minute grid: 8 non-overlapping.
    expect(slots).toHaveLength(8);
  });
});

describe("DST — event bodies", () => {
  it("keeps an all-day event on its local dates across the spring-forward day", () => {
    const body = buildEventBody({
      subject: "AI TEST — teams-mcp — dst all-day",
      start: "2027-03-14T06:00:00.000Z", // local midnight, CST
      end: "2027-03-15T05:00:00.000Z", // local midnight, CDT
      timezone: TZ,
      allDay: true,
    });
    expect(body).toMatchObject({
      IsAllDay: true,
      Start: { DateTime: "2027-03-14T00:00:00.0000000", TimeZone: TZ },
      End: { DateTime: "2027-03-15T00:00:00.0000000", TimeZone: TZ },
    });
  });

  it("writes timed bounds as wall clocks with the zone, never a shifted offset", () => {
    // 01:30 CST -> 04:30 CDT spans the gap: 90 wall-clock minutes, 30 real ones.
    expect(eventTime("2027-03-14T07:30:00.000Z", TZ)).toEqual({
      DateTime: "2027-03-14T01:30:00.0000000",
      TimeZone: TZ,
    });
    expect(eventTime("2027-03-14T09:30:00.000Z", TZ)).toEqual({
      DateTime: "2027-03-14T04:30:00.0000000",
      TimeZone: TZ,
    });
  });
});

describe("all-day events in list rows", () => {
  let dir: string;
  const saved = { ...process.env };

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), "teams-mcp-allday-"));
    process.env.TEAMS_MCP_CONFIG_DIR = dir;
    process.env.TEAMS_MCP_TIMEZONE = TZ;
    resetConfigCache();
  });

  afterEach(async () => {
    setOutlookClient(undefined);
    process.env = { ...saved };
    resetConfigCache();
    await rm(dir, { recursive: true, force: true });
  });

  it("labels a UTC-midnight all-day event by its date, not the previous evening", async () => {
    const fetchImpl: FetchLike = async () =>
      new Response(
        JSON.stringify({
          value: [
            {
              Id: "allday-1",
              Subject: "Holiday",
              IsAllDay: true,
              Start: { DateTime: "2027-03-14T00:00:00.0000000", TimeZone: "UTC" },
              End: { DateTime: "2027-03-15T00:00:00.0000000", TimeZone: "UTC" },
            },
          ],
        }),
        { status: 200 },
      );
    setOutlookClient(
      new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
    );
    const result = await calendarTool.handle({
      op: "list",
      start: "2027-03-14",
      end: "2027-03-16",
    });
    const text = result.content[0];
    const parsed = JSON.parse(text && text.type === "text" ? text.text : "{}") as {
      events?: Array<Record<string, unknown>>;
    };
    expect(parsed.events?.[0]?.local).toBe("2027-03-14 (all day)");
    expect(parsed.events?.[0]?.all_day).toBe(true);
  });
});

describe("DST — calendar.create", () => {
  let dir: string;
  const saved = { ...process.env };

  interface Call {
    url: string;
    body: unknown;
  }

  function harness(responses: Array<() => Response>): Call[] {
    const calls: Call[] = [];
    let index = 0;
    const fetchImpl: FetchLike = async (url, init) => {
      calls.push({ url, body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined });
      const factory = responses[Math.min(index, responses.length - 1)];
      index += 1;
      if (!factory) throw new Error("no response queued");
      return factory();
    };
    setOutlookClient(
      new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
    );
    return calls;
  }

  function textOf(result: { content: unknown[] }): string {
    const first = result.content[0];
    if (!first || typeof first !== "object" || !("text" in first) || typeof first.text !== "string") {
      throw new Error("expected a text result");
    }
    return first.text;
  }

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), "teams-mcp-dst-"));
    process.env.TEAMS_MCP_CONFIG_DIR = dir;
    process.env.TEAMS_MCP_TIMEZONE = TZ;
    resetConfigCache();
  });

  afterEach(async () => {
    setOutlookClient(undefined);
    process.env = { ...saved };
    resetConfigCache();
    await rm(dir, { recursive: true, force: true });
  });

  it("fails with the gap named and issues no request", async () => {
    const calls = harness([]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "AI TEST — teams-mcp — gap",
      start: "2027-03-14T02:30",
      end: "2027-03-14T03:30",
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("does not exist");
    expect(textOf(result)).toContain("America/Chicago");
    expect(calls).toHaveLength(0);
  });

  it("creates a post-gap event at the right instant", async () => {
    const calls = harness([
      () => new Response(JSON.stringify({ value: [] }), { status: 200 }),
      () =>
        new Response(
          JSON.stringify({
            Id: "evt-dst",
            Subject: "AI TEST — teams-mcp — post gap",
            Start: { DateTime: "2027-03-14T03:30:00.0000000", TimeZone: TZ },
            End: { DateTime: "2027-03-14T04:00:00.0000000", TimeZone: TZ },
            IsAllDay: false,
            ShowAs: "Busy",
          }),
          { status: 201 },
        ),
    ]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "AI TEST — teams-mcp — post gap",
      start: "2027-03-14T03:30",
      end: "2027-03-14T04:00",
    });
    expect(result.isError).toBeUndefined();
    // The request carries the post-gap wall clock, with the zone named.
    expect(calls[1]?.body).toMatchObject({
      Start: { DateTime: "2027-03-14T03:30:00.0000000", TimeZone: TZ },
      End: { DateTime: "2027-03-14T04:00:00.0000000", TimeZone: TZ },
    });
    // The echoed zone is converted back: 03:30 CDT == 08:30Z.
    const parsed = JSON.parse(textOf(result)) as Record<string, unknown>;
    expect(parsed.start).toBe("2027-03-14T08:30:00.000Z");
    expect(parsed.end).toBe("2027-03-14T09:00:00.000Z");
  });
});
