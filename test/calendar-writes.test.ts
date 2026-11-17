import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { resetWriteLimiter } from "../src/safety.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { calendarTool } from "../src/tools/calendar.js";
import type { FetchLike } from "../src/http/client.js";

const BASE = "https://outlook.office.com/api/v2.0";

interface Call {
  method: string;
  url: string;
  body: unknown;
}

function harness(responses: Array<() => Response>) {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({
      method: init?.method ?? "GET",
      url,
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    });
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
  new Response(JSON.stringify(body), { status });
const empty = (status: number): (() => Response) => () => new Response(null, { status });

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

function event(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    Id: "evt-1",
    Subject: "Real subject",
    Start: { DateTime: "2026-11-18T16:00:00.0000000", TimeZone: "UTC" },
    End: { DateTime: "2026-11-18T16:30:00.0000000", TimeZone: "UTC" },
    IsAllDay: false,
    ShowAs: "Busy",
    Attendees: [],
    WebLink: "https://outlook.office365.com/owa/?itemid=evt-1",
    ...overrides,
  };
}

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-calendar-writes-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  resetConfigCache();
  resetWriteLimiter();
});

afterEach(async () => {
  setOutlookClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("calendar.create", () => {
  it("posts ISO wall-clock times with the configured timezone and never IsReminderSet", async () => {
    const calls = harness([json({ value: [] }), json(event({ Id: "new-1", Subject: "AI TEST — write probe" }), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "AI TEST — write probe",
      start: "2026-11-18T16:00:00Z",
      attendees: [],
    });
    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(2);
    expect(calls[0]?.url).toContain("calendarview");
    expect(calls[1]?.url).toBe(`${BASE}/me/events`);
    expect(calls[1]?.method).toBe("POST");
    const body = calls[1]?.body as Record<string, unknown>;
    expect(body.Subject).toBe("AI TEST — write probe");
    expect(body.Start).toEqual({ DateTime: "2026-11-18T10:00:00.0000000", TimeZone: "America/Chicago" });
    expect(body.End).toEqual({ DateTime: "2026-11-18T10:30:00.0000000", TimeZone: "America/Chicago" });
    expect(Object.keys(body)).not.toContain("IsReminderSet");
    expect(body.Attendees).toBeUndefined();

    const out = payload(result);
    expect(out).toMatchObject({ ok: true, op: "create", id: "new-1", subject: "AI TEST — write probe" });
    expect(out.web_link).toContain("outlook.office365.com");
  });

  it("returns attendees_detail on create, matching get", async () => {
    harness([
      json({ value: [] }),
      json(
        event({
          Id: "new-att",
          Subject: "AI TEST — write probe",
          Attendees: [
            {
              EmailAddress: { Address: "joel@example.edu", Name: "Student, Test" },
              Type: "Required",
              Status: { Response: "notResponded" },
            },
          ],
        }),
        201,
      ),
    ]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "AI TEST — write probe",
      start: "2026-11-18T16:00:00Z",
      attendees: ["joel@example.edu"],
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.attendees).toEqual(["joel@example.edu"]);
    expect(out.attendees_detail).toEqual([
      {
        email: "joel@example.edu",
        name: "Student, Test",
        type: "Required",
        status: "notResponded",
      },
    ]);
  });

  it("REG: converts create response wall-clock times through the returned timezone", async () => {
    const calls = harness([
      json({ value: [] }),
      json(
        event({
          Id: "new-tz",
          Subject: "AI TEST — write probe",
          Start: { DateTime: "2026-11-17T15:00:00.0000000", TimeZone: "America/Chicago" },
          End: { DateTime: "2026-11-17T15:30:00.0000000", TimeZone: "America/Chicago" },
        }),
        201,
      ),
    ]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "AI TEST — write probe",
      start: "2026-11-17T21:00:00Z",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[1]?.body).toMatchObject({
      Start: { DateTime: "2026-11-17T15:00:00.0000000", TimeZone: "America/Chicago" },
    });
    expect(payload(result)).toMatchObject({
      start: "2026-11-17T21:00:00.000Z",
      end: "2026-11-17T21:30:00.000Z",
      local: "2026-11-17 15:00 → 2026-11-17 15:30",
    });
  });

  it("defaults to now+15m and adds the configured duration", async () => {
    await writeFile(join(dir, "config.yaml"), "default_duration_min: 45\n");
    resetConfigCache();
    const calls = harness([json({ value: [] }), json(event({ Id: "new-2" }), 201)]);
    await calendarTool.handle({ op: "create", subject: "Defaults" });
    const body = calls[1]?.body as Record<string, unknown>;
    const start = Date.parse(`${(body.Start as { DateTime: string }).DateTime}Z`);
    const end = Date.parse(`${(body.End as { DateTime: string }).DateTime}Z`);
    expect(Math.round((end - start) / 60_000)).toBe(45);
    expect((body.Start as { TimeZone: string }).TimeZone).toBe("America/Chicago");
  });

  it("create with start + duration_min derives the end", async () => {
    const calls = harness([json({ value: [] }), json(event({ Id: "new-dur" }), 201)]);
    await calendarTool.handle({
      op: "create",
      subject: "One hour",
      start: "2026-11-18T16:00:00Z",
      duration_min: 60,
    });
    const body = calls[1]?.body as Record<string, unknown>;
    expect(body.Start).toEqual({ DateTime: "2026-11-18T10:00:00.0000000", TimeZone: "America/Chicago" });
    expect(body.End).toEqual({ DateTime: "2026-11-18T11:00:00.0000000", TimeZone: "America/Chicago" });
  });

  it("create with no start/end uses duration_min from now+15m", async () => {
    const calls = harness([json({ value: [] }), json(event({ Id: "new-dur2" }), 201)]);
    await calendarTool.handle({ op: "create", subject: "Ninety", duration_min: 90 });
    const body = calls[1]?.body as Record<string, unknown>;
    const start = Date.parse(`${(body.Start as { DateTime: string }).DateTime}Z`);
    const end = Date.parse(`${(body.End as { DateTime: string }).DateTime}Z`);
    expect(Math.round((end - start) / 60_000)).toBe(90);
    expect((body.Start as { TimeZone: string }).TimeZone).toBe("America/Chicago");
  });

  it("create rejects duration_min when both start and end are supplied", async () => {
    const calls = harness([json(event(), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Ambiguous",
      start: "2026-11-18T16:00:00Z",
      end: "2026-11-18T17:00:00Z",
      duration_min: 30,
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("duration_min");
    expect(calls).toHaveLength(0);
  });

  it("normalises all-day dates to local midnight with an exclusive end day", async () => {
    const calls = harness([json({ value: [] }), json(event({ Id: "new-3" }), 201)]);
    await calendarTool.handle({
      op: "create",
      subject: "All day",
      start: "2026-11-20",
      all_day: true,
    });
    const body = calls[1]?.body as Record<string, unknown>;
    expect(body.IsAllDay).toBe(true);
    expect(body.Start).toEqual({ DateTime: "2026-11-20T00:00:00.0000000", TimeZone: "America/Chicago" });
    expect(body.End).toEqual({ DateTime: "2026-11-21T00:00:00.0000000", TimeZone: "America/Chicago" });
  });

  it("rejects end <= start without calling the API", async () => {
    const calls = harness([json(event(), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Bad",
      start: "2026-11-18T16:00:00Z",
      end: "2026-11-18T15:00:00Z",
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("requires a subject", async () => {
    const calls = harness([json(event(), 201)]);
    const result = await calendarTool.handle({ op: "create", start: "2026-11-18T16:00:00Z" });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("rejects attendees outside invite_allowlist", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@example.com"]\n');
    resetConfigCache();
    const calls = harness([json(event(), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Guarded",
      start: "2026-11-18T16:00:00Z",
      attendees: ["outsider@other.test"],
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe("calendar.list row controls", () => {
  it("omits web_link by default and brief keeps only pick/reference fields", async () => {
    const calls = harness([json({ value: [event()] })]);
    const full = payload(await calendarTool.handle({ op: "list", top: 1 }));
    const row = (full.events as Array<Record<string, unknown>>)[0];
    expect(row?.web_link).toBeUndefined();
    expect(row?.organizer).toBeUndefined();
    expect(calls[0]?.url).toContain("calendarview");

    const brief = payload(await calendarTool.handle({ op: "list", top: 1, brief: true }));
    const briefRow = (brief.events as Array<Record<string, unknown>>)[0];
    expect(Object.keys(briefRow ?? {}).sort()).toEqual(["end", "id", "local", "show_as", "start", "subject"]);

    const linked = payload(await calendarTool.handle({ op: "list", top: 1, include_web_link: true }));
    expect((linked.events as Array<Record<string, unknown>>)[0]?.web_link).toContain("outlook.office365.com");
  });
});

describe("calendar.update guard + patch", () => {
  it("refuses on a subject mismatch before any PATCH", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Wrong subject",
      subject: "Renamed",
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("subject_mismatch");
    expect(calls.map((call) => call.method)).toEqual(["GET"]);
  });

  it("treats case/whitespace differences in expected_subject as a match", async () => {
    const calls = harness([json(event()), json(event({ Subject: "Renamed" }))]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "  real   SUBJECT ",
      subject: "Renamed",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[1]?.method).toBe("PATCH");
    expect(calls[1]?.url).toBe(`${BASE}/me/events/evt-1`);
    expect(calls[1]?.body).toEqual({ Subject: "Renamed" });
    expect(payload(result)).toMatchObject({ ok: true, op: "update", subject: "Renamed" });
  });

  it("returns attendees_detail on update, matching get", async () => {
    const calls = harness([
      json(event()),
      json(
        event({
          Subject: "Real subject",
          Attendees: [
            {
              EmailAddress: { Address: "joel@example.edu", Name: "Student, Test" },
              Type: "Required",
              Status: { Response: "accepted" },
            },
          ],
        }),
      ),
    ]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      subject: "Real subject",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[1]?.method).toBe("PATCH");
    expect(payload(result).attendees_detail).toEqual([
      {
        email: "joel@example.edu",
        name: "Student, Test",
        type: "Required",
        status: "accepted",
      },
    ]);
  });

  it("REG: converts update response wall-clock times through the returned timezone", async () => {
    const calls = harness([
      json(event()),
      json({ value: [] }),
      json(
        event({
          Subject: "Moved",
          Start: { DateTime: "2026-11-18T09:00:00.0000000", TimeZone: "America/Chicago" },
          End: { DateTime: "2026-11-18T09:30:00.0000000", TimeZone: "America/Chicago" },
        }),
      ),
    ]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      start: "2026-11-18T15:00:00Z",
      subject: "Moved",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[2]?.body).toMatchObject({
      Start: { DateTime: "2026-11-18T09:00:00.0000000", TimeZone: "America/Chicago" },
    });
    expect(payload(result)).toMatchObject({
      start: "2026-11-18T15:00:00.000Z",
      local: "2026-11-18 09:00 → 2026-11-18 09:30",
    });
  });

  it("start alone preserves the existing duration (not the old end)", async () => {
    const calls = harness([json(event()), json({ value: [] }), json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      start: "2026-11-18T15:00:00Z",
    });
    const patch = calls[2]?.body as Record<string, unknown>;
    expect(patch.Start).toEqual({ DateTime: "2026-11-18T09:00:00.0000000", TimeZone: "America/Chicago" });
    // The original event is 30 minutes long, so the new end is 09:30 local.
    expect(patch.End).toEqual({ DateTime: "2026-11-18T09:30:00.0000000", TimeZone: "America/Chicago" });
    expect(patch.Subject).toBeUndefined();
    expect(payload(result).changed).toEqual({
      start_before: "2026-11-18T16:00:00.000Z",
      start_after: "2026-11-18T15:00:00.000Z",
      duration_before_min: 30,
      duration_after_min: 30,
    });
  });

  it("requires at least one field", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
    });
    expect(result.isError).toBe(true);
    // Local validation runs before the guard read (and before the rate token).
    expect(calls).toHaveLength(0);
  });
});

describe("calendar conflicts", () => {
  const other = (overrides: Record<string, unknown> = {}): Record<string, unknown> =>
    event({
      Id: "other-1",
      Subject: "Standup",
      Start: { DateTime: "2026-11-18T16:15:00.0000000", TimeZone: "UTC" },
      End: { DateTime: "2026-11-18T16:45:00.0000000", TimeZone: "UTC" },
      ...overrides,
    });

  it("reports an empty conflict list when the slot is clear", async () => {
    const calls = harness([json({ value: [] }), json(event({ Id: "new-clear" }), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Clear",
      start: "2026-11-18T16:00:00Z",
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.conflict_policy).toBe("warn");
    expect(out.conflicts).toEqual([]);
    expect(out.conflict_warning).toBeUndefined();
    expect(calls.map((call) => call.method)).toEqual(["GET", "POST"]);
  });

  it("warn books anyway and returns the conflicts", async () => {
    const calls = harness([json({ value: [other()] }), json(event({ Id: "new-warn" }), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Overlapping",
      start: "2026-11-18T16:00:00Z",
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.conflict_warning).toBe(true);
    expect(out.conflict_count).toBe(1);
    expect(out.conflicts).toEqual([
      {
        id: "other-1",
        subject: "Standup",
        start: "2026-11-18T16:15:00.000Z",
        end: "2026-11-18T16:45:00.000Z",
        local: "2026-11-18 10:15 → 2026-11-18 10:45",
      },
    ]);
    expect(calls.filter((call) => call.method === "POST")).toHaveLength(1);
  });

  it("refuse returns conflict_detected and issues no POST", async () => {
    const calls = harness([json({ value: [other()] })]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Blocked",
      start: "2026-11-18T16:00:00Z",
      conflict_policy: "refuse",
    });
    expect(result.isError).toBe(true);
    const out = JSON.parse(textOf(result)) as Record<string, unknown>;
    expect(out.error).toBe("conflict_detected");
    expect((out.conflicts as unknown[]).length).toBe(1);
    expect(calls.map((call) => call.method)).toEqual(["GET"]);
  });

  it("caps the reported conflicts at five but keeps the true count", async () => {
    const many = Array.from({ length: 6 }, (_, index) =>
      other({
        Id: `other-${index}`,
        Start: { DateTime: `2026-11-18T16:0${index}:00.0000000`, TimeZone: "UTC" },
        End: { DateTime: `2026-11-18T16:3${index}:00.0000000`, TimeZone: "UTC" },
      }),
    );
    harness([json({ value: many }), json(event(), 201)]);
    const out = payload(
      await calendarTool.handle({
        op: "create",
        subject: "Crowded",
        start: "2026-11-18T16:00:00Z",
      }),
    );
    expect((out.conflicts as unknown[]).length).toBe(5);
    expect(out.conflict_count).toBe(6);
  });

  it("ignore skips the conflict read entirely", async () => {
    const calls = harness([json(event({ Id: "new-ignore" }), 201)]);
    const result = await calendarTool.handle({
      op: "create",
      subject: "Unchecked",
      start: "2026-11-18T16:00:00Z",
      conflict_policy: "ignore",
    });
    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("POST");
    expect(payload(result)).toMatchObject({ conflict_policy: "ignore" });
    expect(payload(result).conflicts).toBeUndefined();
  });

  it("does not count a show_as=free event as a conflict", async () => {
    harness([json({ value: [other({ ShowAs: "Free" })] }), json(event(), 201)]);
    const out = payload(
      await calendarTool.handle({ op: "create", subject: "Free overlap", start: "2026-11-18T16:00:00Z" }),
    );
    expect(out.conflicts).toEqual([]);
    expect(out.conflict_warning).toBeUndefined();
  });

  it("update excludes the event being moved and reports the other", async () => {
    const calls = harness([
      json(event()),
      json({ value: [event(), other()] }),
      json(event({ Subject: "Moved" })),
    ]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      subject: "Moved",
      start: "2026-11-18T16:00:00Z",
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect((out.conflicts as Array<Record<string, unknown>>).map((row) => row.id)).toEqual(["other-1"]);
    expect(out.conflict_warning).toBe(true);
    expect(calls.map((call) => call.method)).toEqual(["GET", "GET", "PATCH"]);
  });

  it("update refuse issues no PATCH and keeps the event unchanged", async () => {
    const calls = harness([json(event()), json({ value: [other()] })]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      start: "2026-11-18T16:00:00Z",
      conflict_policy: "refuse",
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("conflict_detected");
    expect(calls.map((call) => call.method)).toEqual(["GET", "GET"]);
  });

  it("a subject-only update does not spend a conflict read", async () => {
    const calls = harness([json(event()), json(event({ Subject: "Renamed" }))]);
    await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      subject: "Renamed",
      conflict_policy: "refuse",
    });
    expect(calls.map((call) => call.method)).toEqual(["GET", "PATCH"]);
  });
});

describe("calendar.update time semantics", () => {
  it("end alone keeps the existing start", async () => {
    const calls = harness([json(event()), json({ value: [] }), json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      end: "2026-11-18T17:30:00Z",
    });
    const patch = calls[2]?.body as Record<string, unknown>;
    expect(patch.Start).toEqual({ DateTime: "2026-11-18T10:00:00.0000000", TimeZone: "America/Chicago" });
    expect(patch.End).toEqual({ DateTime: "2026-11-18T11:30:00.0000000", TimeZone: "America/Chicago" });
    expect(payload(result).changed).toEqual({
      start_before: "2026-11-18T16:00:00.000Z",
      start_after: "2026-11-18T16:00:00.000Z",
      duration_before_min: 30,
      duration_after_min: 90,
    });
  });

  it("shift_min moves both bounds and keeps the duration", async () => {
    const calls = harness([json(event()), json({ value: [] }), json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      shift_min: 30,
    });
    const patch = calls[2]?.body as Record<string, unknown>;
    expect(patch.Start).toEqual({ DateTime: "2026-11-18T10:30:00.0000000", TimeZone: "America/Chicago" });
    expect(patch.End).toEqual({ DateTime: "2026-11-18T11:00:00.0000000", TimeZone: "America/Chicago" });
    expect(payload(result).changed).toEqual({
      start_before: "2026-11-18T16:00:00.000Z",
      start_after: "2026-11-18T16:30:00.000Z",
      duration_before_min: 30,
      duration_after_min: 30,
    });
  });

  it("duration_min with start derives the end", async () => {
    const calls = harness([json(event()), json({ value: [] }), json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      start: "2026-11-18T15:00:00Z",
      duration_min: 60,
    });
    const patch = calls[2]?.body as Record<string, unknown>;
    expect(patch.Start).toEqual({ DateTime: "2026-11-18T09:00:00.0000000", TimeZone: "America/Chicago" });
    expect(patch.End).toEqual({ DateTime: "2026-11-18T10:00:00.0000000", TimeZone: "America/Chicago" });
    expect((payload(result).changed as Record<string, unknown>).duration_after_min).toBe(60);
  });

  it("duration_min with end derives the start", async () => {
    const calls = harness([json(event()), json({ value: [] }), json(event())]);
    await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      end: "2026-11-18T18:00:00Z",
      duration_min: 45,
    });
    const patch = calls[2]?.body as Record<string, unknown>;
    expect(patch.Start).toEqual({ DateTime: "2026-11-18T11:15:00.0000000", TimeZone: "America/Chicago" });
    expect(patch.End).toEqual({ DateTime: "2026-11-18T12:00:00.0000000", TimeZone: "America/Chicago" });
  });

  it("rejects start+end+duration_min as ambiguous before any request", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      start: "2026-11-18T16:00:00Z",
      end: "2026-11-18T17:00:00Z",
      duration_min: 30,
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("rejects shift_min combined with start before any request", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      start: "2026-11-18T16:00:00Z",
      shift_min: 15,
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("refuses an end at or before the start without patching", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "Real subject",
      end: "2026-11-18T15:00:00Z",
    });
    expect(result.isError).toBe(true);
    expect(calls.some((call) => call.method === "PATCH")).toBe(false);
  });
});

describe("calendar.delete", () => {
  it("deletes only after the subject guard passes", async () => {
    const calls = harness([json(event()), empty(204)]);
    const result = await calendarTool.handle({
      op: "delete",
      event_id: "evt-1",
      expected_subject: "Real subject",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[1]?.method).toBe("DELETE");
    expect(calls[1]?.url).toBe(`${BASE}/me/events/evt-1`);
    expect(payload(result)).toEqual({ deleted: true, id: "evt-1", subject: "Real subject" });
  });

  it("refuses and issues no DELETE when the subject differs", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "delete",
      event_id: "evt-1",
      expected_subject: "Something else",
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(1);
  });

  it("requires expected_subject", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({ op: "delete", event_id: "evt-1" });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe("calendar.respond", () => {
  it("posts accept with SendResponse after the guard", async () => {
    const calls = harness([json(event({ Attendees: [{ EmailAddress: { Address: "me@example.com" } }] })), empty(202)]);
    const result = await calendarTool.handle({
      op: "respond",
      event_id: "evt-1",
      expected_subject: "Real subject",
      response: "accept",
      comment: "I'll be there",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[1]?.url).toBe(`${BASE}/me/events/evt-1/accept`);
    expect(calls[1]?.body).toEqual({ Comment: "I'll be there", SendResponse: true });
    expect(payload(result)).toEqual({ responded: true, response: "accept", id: "evt-1" });
  });

  it("maps tentative to tentativelyaccept", async () => {
    const calls = harness([json(event()), empty(202)]);
    await calendarTool.handle({
      op: "respond",
      event_id: "evt-1",
      expected_subject: "Real subject",
      response: "tentative",
    });
    expect(calls[1]?.url).toBe(`${BASE}/me/events/evt-1/tentativelyaccept`);
    expect(calls[1]?.body).toEqual({ SendResponse: true });
  });

  it("refuses on subject mismatch", async () => {
    const calls = harness([json(event())]);
    const result = await calendarTool.handle({
      op: "respond",
      event_id: "evt-1",
      expected_subject: "Nope",
      response: "accept",
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(1);
  });
});

describe("calendar.instances", () => {
  it("queries the instances endpoint and maps rows", async () => {
    const calls = harness([
      json({ value: [event({ Id: "inst-1" }), event({ Id: "inst-2" })] }),
    ]);
    const result = await calendarTool.handle({
      op: "instances",
      event_id: "series-1",
      start: "2026-11-01T00:00:00Z",
      end: "2026-12-01T00:00:00Z",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[0]?.url).toContain(`${BASE}/me/events/series-1/instances?`);
    expect(calls[0]?.url).toContain("startDateTime=2026-11-01T00%3A00%3A00.000Z");
    const out = payload(result);
    expect(out.count).toBe(2);
    expect((out.instances as Array<Record<string, unknown>>)[0]?.id).toBe("inst-1");
  });
});
