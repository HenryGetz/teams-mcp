/**
 * `calendar.agenda`: compact "what's my day" view with no attendees, location,
 * organizer or web link. Read-only and paginated through the existing cursor.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { calendarTool } from "../src/tools/calendar.js";
import type { FetchLike } from "../src/http/client.js";

const TZ = "America/Chicago";

interface Call {
  method: string;
  url: string;
}

function harness(responses: Array<() => Response>) {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({ method: init?.method ?? "GET", url });
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

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

function event(id: string, subject: string, start: string, end: string, extra: Record<string, unknown> = {}) {
  return {
    Id: id,
    Subject: subject,
    Start: { DateTime: start, TimeZone: "UTC" },
    End: { DateTime: end, TimeZone: "UTC" },
    IsAllDay: false,
    ShowAs: "Busy",
    Location: { DisplayName: "Secret Room" },
    Organizer: { EmailAddress: { Address: "boss@example.com" } },
    Attendees: [{ EmailAddress: { Address: "peer@example.com" }, Type: "Required" }],
    WebLink: "https://outlook.office365.com/owa/?itemid=x",
    ...extra,
  };
}

/** Friday 2026-11-20, 09:00 and 10:00 CST. */
const EVENTS = [
  event("evt-1", "Standup", "2026-11-20T15:00:00.0000000", "2026-11-20T15:30:00.0000000"),
  event("evt-2", "Deep work", "2026-11-20T16:00:00.0000000", "2026-11-20T17:00:00.0000000"),
];

const saved = { ...process.env };

beforeEach(() => {
  process.env.TEAMS_MCP_TIMEZONE = TZ;
  resetConfigCache();
});

afterEach(() => {
  setOutlookClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
});

describe("calendar.agenda", () => {
  it("returns compact rows without attendees, location or organizer", async () => {
    const calls = harness([json({ value: EVENTS })]);
    const out = payload(
      await calendarTool.handle({
        op: "agenda",
        start: "2026-11-20T00:00:00-06:00",
        end: "2026-11-21T00:00:00-06:00",
      }),
    );
    expect(out.count).toBe(2);
    expect(out.truncated).toBe(false);
    expect(out.range).toEqual({
      start: "2026-11-20 00:00",
      end: "2026-11-21 00:00",
      timezone: TZ,
    });
    expect(out.events).toEqual([
      {
        date: "2026-11-20",
        start_local: "09:00",
        end_local: "09:30",
        time: "09:00–09:30",
        subject: "Standup",
        id: "evt-1",
      },
      {
        date: "2026-11-20",
        start_local: "10:00",
        end_local: "11:00",
        time: "10:00–11:00",
        subject: "Deep work",
        id: "evt-2",
      },
    ]);
    const row = (out.events as Array<Record<string, unknown>>)[0]!;
    expect(row).not.toHaveProperty("attendees");
    expect(row).not.toHaveProperty("location");
    expect(row).not.toHaveProperty("organizer");
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
  });

  it("omits ids when include_ids is false", async () => {
    harness([json({ value: EVENTS })]);
    const out = payload(
      await calendarTool.handle({
        op: "agenda",
        start: "2026-11-20T00:00:00-06:00",
        end: "2026-11-21T00:00:00-06:00",
        include_ids: false,
      }),
    );
    expect((out.events as Array<Record<string, unknown>>)[0]).not.toHaveProperty("id");
  });

  it("reports truncation and a cursor when the page is cut", async () => {
    harness([
      json({
        value: [EVENTS[0]],
        "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/calendarview?$skip=1",
      }),
    ]);
    const out = payload(
      await calendarTool.handle({
        op: "agenda",
        start: "2026-11-20T00:00:00-06:00",
        end: "2026-11-21T00:00:00-06:00",
        top: 1,
      }),
    );
    expect(out.count).toBe(1);
    expect(out.truncated).toBe(true);
    expect(typeof out.next_cursor).toBe("string");
  });

  it("labels all-day events by date, not a midnight conversion", async () => {
    harness([
      json({
        value: [
          event("evt-ad", "Holiday", "2026-11-20T00:00:00.0000000", "2026-11-21T00:00:00.0000000", {
            IsAllDay: true,
          }),
        ],
      }),
    ]);
    const out = payload(
      await calendarTool.handle({
        op: "agenda",
        start: "2026-11-20",
        end: "2026-11-21",
      }),
    );
    expect((out.events as Array<Record<string, unknown>>)[0]).toEqual({
      date: "2026-11-20",
      start_local: "all day",
      end_local: "all day",
      time: "all day",
      subject: "Holiday",
      id: "evt-ad",
    });
  });

  it("defaults to today when start and end are omitted", async () => {
    harness([json({ value: [] })]);
    const out = payload(await calendarTool.handle({ op: "agenda" }));
    // The window is one local day, whatever the wall-clock date is.
    const range = out.range as Record<string, string>;
    expect(range.timezone).toBe(TZ);
    expect(range.end).not.toBe(range.start);
    expect(Date.parse(range.end!.replace(" ", "T"))).toBeGreaterThan(
      Date.parse(range.start!.replace(" ", "T")),
    );
  });
});
