/**
 * `availability.suggest` for a self-only request.
 *
 * Outlook's findMeetingTimes often returns candidates with an empty
 * AttendeeAvailability for the organizer alone. Before this fix, every
 * self-only candidate was free:"unknown"/verified:false. The local snapshot
 * search is authoritative in that shape and must still return candidates.
 */
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { availabilityTool } from "../src/tools/availability.js";
import { clearSnapshotCache, SNAPSHOT_VERSION, writeSnapshot } from "../src/snapshot/calendar.js";
import { localDate, parseInstant } from "../src/time.js";
import type { FetchLike } from "../src/http/client.js";

const TZ = "America/Chicago";
const SELF = "self@example.com";

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

interface Call {
  url: string;
  body: Record<string, unknown> | undefined;
}

function harness(responses: Array<() => Response>) {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({
      url,
      body: typeof init?.body === "string" ? (JSON.parse(init.body) as Record<string, unknown>) : undefined,
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

/** A weekday two days out, in local time, so "now" filtering never eats it. */
const DAY = localDate(Date.now() + 2 * 86_400_000, TZ);
const DAY_START = new Date(parseInstant(`${DAY}T09:00`, TZ) ?? Date.now()).toISOString();
const DAY_END = new Date(parseInstant(`${DAY}T12:00`, TZ) ?? Date.now()).toISOString();
const BUSY_START = new Date(parseInstant(`${DAY}T10:00`, TZ) ?? Date.now()).toISOString();
const BUSY_END = new Date(parseInstant(`${DAY}T10:30`, TZ) ?? Date.now()).toISOString();

let dir: string;
const saved = { ...process.env };

async function snapshot() {
  const fetchedAt = new Date().toISOString();
  await writeSnapshot({
    version: SNAPSHOT_VERSION,
    metadata: {
      fetched_at: fetchedAt,
      range_start: new Date(Date.parse(DAY_START) - 3_600_000).toISOString(),
      range_end: new Date(Date.parse(DAY_END) + 3_600_000).toISOString(),
      timezone: TZ,
      ttl_seconds: 3600,
      stale_after: new Date(Date.parse(fetchedAt) + 3_600_000).toISOString(),
      source: "outlook",
      attendees: [SELF],
      unknown: [],
      data_available: [SELF],
    },
    events: [],
    busy: { [SELF]: [{ start: BUSY_START, end: BUSY_END, status: "busy" }] },
  });
}

async function token() {
  await writeFile(
    join(dir, "token.json"),
    JSON.stringify({
      version: 1,
      tokens: {
        outlook: { access_token: "t", expires_at: new Date(Date.now() + 3_600_000).toISOString() },
      },
      source: "refresh",
      account: { upn: SELF },
      updated_at: new Date().toISOString(),
    }),
  );
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-suggest-self-"));
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

describe("availability.suggest self-only", () => {
  it("falls back to local search when the service has no self availability", async () => {
    await token();
    await snapshot();
    const calls = harness([
      json({
        MeetingTimeSuggestions: [
          suggestion(DAY_START, new Date(Date.parse(DAY_START) + 1_800_000).toISOString(), []),
        ],
      }),
    ]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        start: DAY_START,
        end: DAY_END,
        duration_min: 30,
        interval_min: 30,
        working_hours_only: false,
      }),
    );
    expect(out.source).toBe("local");
    expect(out.returned).toBeGreaterThan(0);
    expect(out.data_warning).toBeUndefined();
    const candidate = (out.candidates as Array<Record<string, unknown>>)[0];
    expect(candidate).toMatchObject({ free: "1/1", data_available: true, verified: true, source: "local" });
    // 10:00–10:30 is busy, so the earliest free slot is 09:00.
    expect(Date.parse(String(candidate?.start))).toBe(Date.parse(DAY_START));
    // findMeetingTimes was asked about the signed-in user, not an empty list.
    const body = calls.find((call) => call.url.includes("findmeetingtimes"))?.body;
    const attendees = body?.Attendees as Array<Record<string, unknown>> | undefined;
    expect(attendees?.[0]?.EmailAddress).toEqual({ Address: SELF });
  });

  it("keeps service candidates when they carry self availability", async () => {
    await token();
    await snapshot();
    const slotEnd = new Date(Date.parse(DAY_START) + 1_800_000).toISOString();
    harness([
      json({
        MeetingTimeSuggestions: [
          suggestion(DAY_START, slotEnd, [
            { Attendee: { EmailAddress: { Address: SELF } }, Availability: "Free" },
          ]),
        ],
      }),
    ]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        start: DAY_START,
        end: DAY_END,
        duration_min: 30,
        working_hours_only: false,
      }),
    );
    expect(out.source).toBe("findmeetingtimes");
    expect((out.candidates as Array<Record<string, unknown>>)[0]).toMatchObject({
      free: "1/1",
      verified: true,
    });
  });

  it("returns local candidates when the service fails", async () => {
    await token();
    await snapshot();
    harness([json({ error: { code: "MockFailure", message: "no service" } }, 400)]);
    const out = payload(
      await availabilityTool.handle({
        op: "suggest",
        start: DAY_START,
        end: DAY_END,
        duration_min: 30,
        working_hours_only: false,
      }),
    );
    expect(out.source).toBe("local");
    expect(out.returned).toBeGreaterThan(0);
    expect(out.service_error).toBeTruthy();
  });
});
