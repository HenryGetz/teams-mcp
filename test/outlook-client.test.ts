import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  camelizeKeys,
  nextLinkOf,
  normalizeEvent,
  normalizeMessage,
  OutlookClient,
  parseUnsubscribe,
  readCollection,
} from "../src/outlook/client.js";
import {
  findMeetingTimes,
  getSchedule,
  restDateTimeToIso,
} from "../src/outlook/calendar.js";
import type { FetchLike } from "../src/http/client.js";

const ARTIFACTS = "docs/recon-artifacts";
const hasFixtures = existsSync(`${ARTIFACTS}/sample-calendarview.json`);

function fixture(name: string): Record<string, unknown> {
  return JSON.parse(readFileSync(`${ARTIFACTS}/${name}`, "utf8")) as Record<string, unknown>;
}

function stubClient(handler: (url: string, init?: RequestInit) => Response | Promise<Response>) {
  const calls: Array<{ url: string; method: string; body: unknown }> = [];
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({
      url,
      method: init?.method ?? "GET",
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    });
    return handler(url, init);
  };
  const client = new OutlookClient({
    tokenProvider: async () => ({ token: "test-token" }),
    fetchImpl,
    maxAttempts: 1,
  });
  return { client, calls };
}

describe("PascalCase boundary", () => {
  it("camelises nested keys and drops @odata members", () => {
    expect(
      camelizeKeys({
        "@odata.etag": "x",
        Subject: "Hi",
        ReceivedDateTime: "2026-11-16T12:00:00Z",
        From: { EmailAddress: { Name: "A", Address: "a@example.com" } },
        InternetMessageHeaders: [{ Name: "List-Unsubscribe", Value: "<mailto:u@x>" }],
      }),
    ).toEqual({
      subject: "Hi",
      receivedDateTime: "2026-11-16T12:00:00Z",
      from: { emailAddress: { name: "A", address: "a@example.com" } },
      internetMessageHeaders: [{ name: "List-Unsubscribe", value: "<mailto:u@x>" }],
    });
  });

  it("reads the value[] envelope and @odata.nextLink", () => {
    const payload = {
      "@odata.context": "ctx",
      "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/messages?$skip=2",
      value: [{ Id: "1", Subject: "One" }, { Id: "2", Subject: "Two" }],
    };
    const collection = readCollection(payload, (row) => row);
    expect(collection.items).toHaveLength(2);
    expect(collection.nextLink).toContain("$skip=2");
    expect(nextLinkOf(payload)).toContain("$skip=2");
  });

  it("normalizes a REST DateTime with 7 fractional digits and no offset", () => {
    expect(restDateTimeToIso("2026-11-16T21:30:00.0000000")).toBe("2026-11-16T21:30:00.000Z");
    expect(restDateTimeToIso("2026-11-16T21:30:00.1234567Z")).toBe("2026-11-16T21:30:00.123Z");
  });

  it("REG: converts a wall clock through the payload's TimeZone instead of assuming UTC", () => {
    // Write responses echo the timezone sent in the request body.
    expect(restDateTimeToIso("2026-11-17T15:00:00.0000000", "America/Chicago")).toBe(
      "2026-11-17T21:00:00.000Z",
    );
    expect(restDateTimeToIso("2026-11-17T15:00:00.0000000", "UTC")).toBe("2026-11-17T15:00:00.000Z");
    // Calendarview with Prefer: outlook.timezone="America/Chicago" is wall clock too.
    expect(
      normalizeEvent({
        id: "e1",
        subject: "Wall clock",
        start: { dateTime: "2026-11-02T15:30:00.0000000", timeZone: "America/Chicago" },
        end: { dateTime: "2026-11-02T17:30:00.0000000", timeZone: "America/Chicago" },
      }),
    ).toMatchObject({ start: "2026-11-02T21:30:00.000Z", end: "2026-11-02T23:30:00.000Z" });
  });
});

describe("message mapping", () => {
  it("maps a camelised REST message into the compact shape", () => {
    const message = normalizeMessage({
      id: "AAMk",
      subject: "Standup notes",
      receivedDateTime: "2026-11-16T15:00:00Z",
      isRead: false,
      flag: { flagStatus: "flagged" },
      bodyPreview: "Here are the notes",
      hasAttachments: true,
      webLink: "https://outlook.office.com/owa/?ItemID=AAMk",
      from: { emailAddress: { name: "Pat", address: "pat@example.com" } },
      toRecipients: [{ emailAddress: { address: "me@example.com" } }],
      body: { contentType: "html", content: "<p>hi</p>" },
      internetMessageHeaders: [{ name: "List-Unsubscribe", value: "<mailto:u@example.com>" }],
    });
    expect(message).toMatchObject({
      id: "AAMk",
      subject: "Standup notes",
      unread: true,
      flagged: true,
      hasAttachments: true,
      from: { email: "pat@example.com" },
      to: [{ email: "me@example.com" }],
      body: { contentType: "html", content: "<p>hi</p>" },
    });
    expect(message.received?.endsWith("Z")).toBe(true);
  });

  it("parses List-Unsubscribe shapes observed in this mailbox", () => {
    const both = parseUnsubscribe([
      {
        name: "List-Unsubscribe",
        value: "<https://go.example.com/u/abc?e=1>, <mailto:leave@example.com?subject=unsub>",
      },
      { name: "List-Unsubscribe-Post", value: "List-Unsubscribe=One-Click" },
    ]);
    expect(both).toMatchObject({
      url: "https://go.example.com/u/abc?e=1",
      mailto: "mailto:leave@example.com?subject=unsub",
      oneClick: true,
    });

    const onlyMail = parseUnsubscribe([
      { name: "List-Unsubscribe", value: "<mailto:leave@example.com>" },
    ]);
    expect(onlyMail?.mailto).toBe("mailto:leave@example.com");
    expect(onlyMail?.url).toBeUndefined();
    expect(parseUnsubscribe(undefined)).toBeUndefined();
  });
});

describe("event mapping", () => {
  it("maps a camelised REST event", () => {
    const event = normalizeEvent({
      id: "evt-1",
      subject: "Sync",
      start: { dateTime: "2026-11-16T21:30:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-11-16T22:00:00.0000000", timeZone: "UTC" },
      isAllDay: false,
      showAs: "Busy",
      location: { displayName: "Teams" },
      organizer: { emailAddress: { name: "Pat", address: "pat@example.com" } },
      attendees: [{ emailAddress: { address: "a@example.com" }, type: "Required", status: { response: "Accepted" } }],
      isOnlineMeeting: true,
      webLink: "https://outlook.office.com/calendar/item/evt-1",
    });
    expect(event.start).toBe("2026-11-16T21:30:00.000Z");
    expect(event.attendees).toEqual([
      { email: "a@example.com", type: "Required", status: "Accepted" },
    ]);
    expect(event.location).toBe("Teams");
  });
});

describe("REST calls", () => {
  it("posts PascalCase getSchedule and normalises busy blocks", async () => {
    const { client, calls } = stubClient(
      () =>
        new Response(
          JSON.stringify({
            value: [
              {
                ScheduleId: "a@example.com",
                AvailabilityView: "0002",
                ScheduleItems: [
                  {
                    Status: "Busy",
                    Start: { DateTime: "2026-11-16T21:00:00.0000000", TimeZone: "UTC" },
                    End: { DateTime: "2026-11-16T22:00:00.0000000", TimeZone: "UTC" },
                  },
                ],
              },
            ],
          }),
          { status: 200 },
        ),
    );
    const schedules = await getSchedule(client, {
      attendees: ["a@example.com"],
      range: { start: "2026-11-16T00:00:00.000Z", end: "2026-11-17T00:00:00.000Z" },
      intervalMinutes: 30,
    });
    expect(schedules[0]).toMatchObject({
      email: "a@example.com",
      availability_view: "0002",
      blocks: [{ start: "2026-11-16T21:00:00.000Z", end: "2026-11-16T22:00:00.000Z", status: "busy" }],
    });
    const body = calls[0]?.body as Record<string, unknown>;
    expect(body.Schedules).toEqual(["a@example.com"]);
    expect(body.StartTime).toEqual({ DateTime: "2026-11-16T00:00:00", TimeZone: "UTC" });
    expect(body.AvailabilityViewInterval).toBe(30);
  });

  it("posts PascalCase findMeetingTimes and maps suggestions", async () => {
    const { client, calls } = stubClient(
      () =>
        new Response(
          JSON.stringify({
            MeetingTimeSuggestions: [
              {
                Confidence: 100,
                SuggestionReason: "All free",
                MeetingTimeSlot: {
                  Start: { DateTime: "2026-11-16T21:30:00.0000000", TimeZone: "UTC" },
                  End: { DateTime: "2026-11-16T22:00:00.0000000", TimeZone: "UTC" },
                },
                AttendeeAvailability: [
                  {
                    Availability: "Free",
                    Attendee: { EmailAddress: { Address: "a@example.com" } },
                  },
                ],
              },
            ],
          }),
          { status: 200 },
        ),
    );
    const suggestions = await findMeetingTimes(client, {
      attendees: ["a@example.com"],
      range: { start: "2026-11-16T00:00:00.000Z", end: "2026-11-20T00:00:00.000Z" },
      durationMin: 30,
    });
    expect(suggestions[0]).toMatchObject({
      start: "2026-11-16T21:30:00.000Z",
      confidence: 100,
      attendeeAvailability: [{ email: "a@example.com", availability: "free" }],
    });
    const body = calls[0]?.body as Record<string, unknown>;
    expect(body.MeetingDuration).toBe("PT30M");
    expect(Array.isArray(body.Attendees)).toBe(true);
  });
});

describe.skipIf(!hasFixtures)("real recon fixtures", () => {
  it("parses every calendarview event", () => {
    const payload = fixture("sample-calendarview.json");
    const collection = readCollection(payload, normalizeEvent);
    expect(collection.items.length).toBeGreaterThan(0);
    for (const event of collection.items) {
      expect(event.id).not.toBe("");
      expect(Date.parse(event.start)).not.toBeNaN();
      expect(Date.parse(event.end)).not.toBeNaN();
    }
  });

  it("parses every message row and its headers", () => {
    const payload = fixture("sample-message-metadata.json");
    const collection = readCollection(payload, normalizeMessage);
    expect(collection.items.length).toBeGreaterThan(0);
    const first = collection.items[0];
    expect(first?.id).toBeTruthy();
    expect(first?.subject).toBeTruthy();
    expect(Date.parse(first?.received ?? "")).not.toBeNaN();
    const withHeaders = collection.items.find((message) => (message.headers?.length ?? 0) > 0);
    expect(withHeaders).toBeDefined();
  });

  it("maps real getSchedule rows for two users", () => {
    const payload = fixture("sample-getschedule-2users.json");
    const rows = Array.isArray(payload.value) ? (payload.value as Record<string, unknown>[]) : [];
    expect(rows.length).toBeGreaterThan(1);
    const camel = camelizeKeys<Array<Record<string, unknown>>>(rows);
    expect(camel[0]?.scheduleId).toBeTruthy();
    expect(Array.isArray(camel[0]?.scheduleItems)).toBe(true);
  });

  it("maps real findmeetingtimes suggestions", () => {
    const payload = fixture("sample-findmeetingtimes.json");
    const rows = Array.isArray(payload.value) ? (payload.value as Record<string, unknown>[]) : [];
    expect(rows.length).toBeGreaterThan(0);
    const row = camelizeKeys<Record<string, unknown>>(rows[0] ?? {});
    const slot = row.meetingTimeSlot as Record<string, unknown> | undefined;
    expect(restDateTimeToIso((slot?.start as Record<string, unknown> | undefined)?.dateTime)).toMatch(/Z$/);
  });
});
