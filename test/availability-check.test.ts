/**
 * `availability.check`: one exact proposed slot, no write, no rate token.
 */
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { availabilityTool } from "../src/tools/availability.js";
import type { BusyBlock } from "../src/snapshot/calendar.js";
import {
  clearSnapshotCache,
  SNAPSHOT_VERSION,
  writeSnapshot,
} from "../src/snapshot/calendar.js";
import type { FetchLike } from "../src/http/client.js";

const TZ = "America/Chicago";
const SELF = "self@example.com";
const TYPO = "typo@example.com";

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

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

const busy = (start: string, end: string, status: BusyBlock["status"] = "busy"): BusyBlock => ({
  start,
  end,
  status,
});

async function snapshot(entries: Record<string, BusyBlock[]>, unknown: string[] = [], dataAvailable: string[]) {
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
      attendees: Object.keys(entries),
      unknown,
      data_available: dataAvailable,
    },
    events: [],
    busy: entries,
  });
}

async function token(upn: string) {
  const now = new Date().toISOString();
  await writeFile(
    join(dir, "token.json"),
    JSON.stringify({
      version: 1,
      tokens: {
        outlook: { access_token: "t", expires_at: new Date(Date.now() + 3_600_000).toISOString() },
      },
      source: "refresh",
      account: { upn },
      updated_at: now,
    }),
  );
}

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-check-"));
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

describe("availability.check", () => {
  it("returns available when every known attendee is free", async () => {
    await snapshot({ [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z")] }, [], [SELF]);
    const calls = noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "check",
        attendees: [SELF],
        start: "2026-11-20T17:00:00.000Z",
        end: "2026-11-20T17:30:00.000Z",
      }),
    );
    expect(out.available).toBe(true);
    expect(out.conflicts).toEqual([]);
    expect(out.conflict_count).toBe(0);
    expect(out.unknown).toEqual([]);
    expect(out.range).toEqual({
      start: "2026-11-20 11:00",
      end: "2026-11-20 11:30",
      timezone: TZ,
    });
    expect(out.attendees).toEqual([
      { email: SELF, data_available: true, unknown: false, status: "free" },
    ]);
    expect(calls).toEqual([]);
  });

  it("reports the overlapping conflict for a busy slot", async () => {
    await snapshot(
      { [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z")] },
      [],
      [SELF],
    );
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "check",
        attendees: [SELF],
        start: "2026-11-20T16:15:00.000Z",
        end: "2026-11-20T16:45:00.000Z",
      }),
    );
    expect(out.available).toBe(false);
    expect(out.conflict_count).toBe(1);
    expect((out.attendees as Array<Record<string, unknown>>)[0]).toMatchObject({ status: "busy" });
    expect((out.conflicts as Array<Record<string, unknown>>)[0]).toEqual({
      email: SELF,
      status: "busy",
      start: "2026-11-20T16:15:00.000Z",
      end: "2026-11-20T16:30:00.000Z",
      local: "2026-11-20 10:15 → 2026-11-20 10:30",
    });
  });

  it("caps the conflict rows at five but keeps the true count", async () => {
    const blocks = Array.from({ length: 7 }, (_, index) => {
      const start = Date.parse("2026-11-20T15:00:00.000Z") + index * 7 * 60_000;
      return busy(new Date(start).toISOString(), new Date(start + 5 * 60_000).toISOString());
    });
    await snapshot({ [SELF]: blocks }, [], [SELF]);
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "check",
        attendees: [SELF],
        start: "2026-11-20T15:00:00.000Z",
        end: "2026-11-20T16:00:00.000Z",
      }),
    );
    expect(out.available).toBe(false);
    expect(out.conflict_count).toBe(7);
    expect((out.conflicts as unknown[]).length).toBe(5);
  });

  it("treats an unknown attendee as not available, never free", async () => {
    await snapshot(
      { [SELF]: [], [TYPO]: [] },
      [TYPO],
      [SELF],
    );
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "check",
        attendees: [SELF, TYPO],
        start: "2026-11-20T17:00:00.000Z",
        end: "2026-11-20T17:30:00.000Z",
      }),
    );
    expect(out.available).toBe(false);
    expect(out.unknown).toEqual([TYPO]);
    expect((out.attendees as Array<Record<string, unknown>>).find((row) => row.email === TYPO)).toEqual({
      email: TYPO,
      data_available: false,
      unknown: true,
      status: "unknown",
    });
  });

  it("defaults attendees to the signed-in user", async () => {
    await token(SELF);
    await snapshot({ [SELF]: [] }, [], [SELF]);
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "check",
        start: "2026-11-20T17:00:00.000Z",
        end: "2026-11-20T17:30:00.000Z",
      }),
    );
    expect(out.available).toBe(true);
    expect((out.attendees as Array<Record<string, unknown>>).map((row) => row.email)).toEqual([SELF]);
  });

  it("requires an exact start and end", async () => {
    noNetwork();
    const result = await availabilityTool.handle({ op: "check", start: "2026-11-20T17:00:00.000Z" });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("check requires both start and end");
  });

  it("ignores free blocks in the conflict count", async () => {
    await snapshot({ [SELF]: [busy("2026-11-20T16:00:00.000Z", "2026-11-20T16:30:00.000Z", "free")] }, [], [SELF]);
    noNetwork();
    const out = payload(
      await availabilityTool.handle({
        op: "check",
        attendees: [SELF],
        start: "2026-11-20T16:15:00.000Z",
        end: "2026-11-20T16:45:00.000Z",
      }),
    );
    expect(out.available).toBe(true);
    expect(out.conflict_count).toBe(0);
  });
});
