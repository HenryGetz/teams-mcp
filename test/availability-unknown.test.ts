/**
 * `busy` must mean busy, and an unresolved mailbox must never read as free.
 *
 * Before this suite: (a) `freebusy` emitted `status:"free"` intervals under
 * `busy`, and (b) a typo'd/unresolvable attendee (`getSchedule` returns no
 * `WorkingHours` and no items) produced `busy: []` + all-zero
 * `availability_view`, i.e. "free all day".
 */
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { availabilityTool } from "../src/tools/availability.js";
import { summarizeBusyBlocks } from "../src/availability/slots.js";
import {
  buildSnapshot,
  clearSnapshotCache,
  ensureSnapshot,
  readSnapshot,
  SNAPSHOT_VERSION,
  snapshotPath,
  writeSnapshot,
  type CalendarSnapshot,
  type CalendarSource,
} from "../src/snapshot/calendar.js";
import type { FetchLike } from "../src/http/client.js";

// Anchored to the next weekday so the suite does not become a time bomb: slot
// search skips times already in the past.
const DAY = (() => {
  const day = new Date();
  day.setUTCHours(0, 0, 0, 0);
  do day.setUTCDate(day.getUTCDate() + 1);
  while (day.getUTCDay() === 0 || day.getUTCDay() === 6);
  return day;
})();
/** UTC time on that day, e.g. at(15) = 15:00Z. */
const at = (hours: number, minutes = 0): string =>
  new Date(DAY.getTime() + (hours * 60 + minutes) * 60_000).toISOString();
/** The same instant rendered the way summarizeBusyBlocks labels it. */
function localLabel(iso: string): string {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Chicago",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(new Date(iso));
  const value = (type: string): string => parts.find((part) => part.type === type)?.value ?? "";
  return `${value("year")}-${value("month")}-${value("day")} ${value("hour")}:${value("minute")}`;
}

const START = at(15); // 09:00 America/Chicago (CST)
const END = at(18); // 12:00
const KNOWN = "known@example.com";
const TYPO = "typo@example.com";

let dir: string;
const saved = { ...process.env };

function snapshot(): CalendarSnapshot {
  const fetchedAt = new Date().toISOString();
  return {
    version: SNAPSHOT_VERSION,
    metadata: {
      fetched_at: fetchedAt,
      range_start: START,
      range_end: END,
      timezone: "America/Chicago",
      ttl_seconds: 3600,
      stale_after: new Date(Date.parse(fetchedAt) + 3_600_000).toISOString(),
      source: "outlook",
      attendees: [KNOWN, TYPO],
      unknown: [TYPO],
    },
    events: [],
    busy: {
      [KNOWN]: [
        // Legacy shape: an old snapshot may still carry a free interval.
        { start: at(15, 30), end: at(16), status: "free" },
        { start: at(16), end: at(16, 30), status: "busy" },
      ],
      [TYPO]: [],
    },
  };
}

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-unknown-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  process.env.TEAMS_MCP_TIMEZONE = "America/Chicago";
  resetConfigCache();
  clearSnapshotCache();
  // Any service call fails fast; the tool then falls back to local slot search.
  const fetchImpl: FetchLike = async () => new Response("nope", { status: 503 });
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
  await writeSnapshot(snapshot(), snapshotPath());
});

afterEach(async () => {
  setOutlookClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  clearSnapshotCache();
  await rm(dir, { recursive: true, force: true });
});

describe("free intervals never appear under busy", () => {
  it("drops them from a legacy snapshot on read", () => {
    const rows = summarizeBusyBlocks(snapshot().busy[KNOWN] ?? [], "America/Chicago", 20);
    expect(rows).toEqual([
      { start: localLabel(at(16)), end: localLabel(at(16, 30)), status: "busy" },
    ]);
  });

  it("freebusy reports only non-free blocks", async () => {
    const result = await availabilityTool.handle({
      op: "freebusy",
      attendees: [KNOWN],
      start: START,
      end: END,
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    const attendee = (out.attendees as Array<Record<string, unknown>>)[0];
    expect(attendee?.busy).toEqual([
      { start: localLabel(at(16)), end: localLabel(at(16, 30)), status: "busy" },
    ]);
    expect(JSON.stringify(attendee?.busy)).not.toContain("free");
    // Digits still encode the free part of the window.
    expect(out.availability_view).toMatchObject({ [KNOWN]: "002000" });
  });
});

describe("unknown attendees", () => {
  it("freebusy marks the mailbox unknown instead of free", async () => {
    const result = await availabilityTool.handle({
      op: "freebusy",
      attendees: [KNOWN, TYPO],
      start: START,
      end: END,
    });
    const out = payload(result);
    expect(out.unknown).toEqual([TYPO]);
    expect(out.warning).toContain("not evidence of availability");
    const rows = out.attendees as Array<Record<string, unknown>>;
    expect(rows.find((row) => row.email === TYPO)).toMatchObject({ unknown: true, busy: [] });
    expect(rows.find((row) => row.email === KNOWN)?.unknown).toBeUndefined();
    // No all-zero digit string that would read as a free day.
    expect(out.availability_view).not.toHaveProperty(TYPO);
  });

  it("suggest refuses when no attendee resolves", async () => {
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: [TYPO],
      start: START,
      end: END,
      duration_min: 30,
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("unknown_attendees");
    expect(textOf(result)).toContain(TYPO);
  });

  it("suggest excludes unknown attendees with a warning, never ranking them free", async () => {
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: [KNOWN, TYPO],
      start: START,
      end: END,
      duration_min: 30,
      interval_min: 30,
      max_candidates: 5,
    });
    expect(result.isError).toBeUndefined();
    const out = payload(result);
    expect(out.excluded_unknown_attendees).toEqual([TYPO]);
    expect(out.warning).toContain("excluded from ranking");
    const candidates = out.candidates as Array<Record<string, unknown>>;
    expect(candidates.length).toBeGreaterThan(0);
    // Only the known attendee is ranked; no conflict line names the typo.
    expect(JSON.stringify(candidates)).not.toContain(TYPO);
    const ranked = candidates[0];
    expect(String(ranked?.free)).toMatch(/^[01]\/1$/);
  });
});

describe("legacy snapshots cannot hide unknown", () => {
  /** A version-1 document as written before `metadata.unknown` was required. */
  const legacyYaml = () =>
    [
      "version: 1",
      "metadata:",
      `  fetched_at: ${new Date().toISOString()}`,
      `  range_start: ${START}`,
      `  range_end: ${END}`,
      "  timezone: America/Chicago",
      "  ttl_seconds: 3600",
      `  stale_after: ${new Date(Date.now() + 3_600_000).toISOString()}`,
      "  source: outlook",
      `  attendees: [${KNOWN}, ${TYPO}]`,
      "events: []",
      "busy:",
      `  ${KNOWN}: []`,
      `  ${TYPO}: []`,
      "",
    ].join("\n");

  it("rejects a snapshot without unknown metadata on read", async () => {
    await writeFile(snapshotPath(), legacyYaml());
    clearSnapshotCache();
    expect(await readSnapshot(snapshotPath())).toBeNull();
  });

  it("rebuilds instead of reusing a legacy snapshot, then reports the unknown mailbox", async () => {
    await writeFile(snapshotPath(), legacyYaml());
    clearSnapshotCache();

    const source: CalendarSource = {
      async ownEvents() {
        return [];
      },
      async schedules(attendees) {
        return attendees.map((email) => ({
          email,
          blocks: [],
          ...(email === TYPO ? { unknown: true as const } : {}),
        }));
      },
    };
    const rebuilt = await ensureSnapshot({
      source,
      attendees: [KNOWN, TYPO],
      start: START,
      end: END,
    });
    expect(rebuilt.refreshed).toBe(true);
    expect(rebuilt.snapshot.version).toBe(SNAPSHOT_VERSION);
    expect(rebuilt.snapshot.metadata.unknown).toEqual([TYPO]);

    const result = await availabilityTool.handle({
      op: "freebusy",
      attendees: [KNOWN, TYPO],
      start: START,
      end: END,
    });
    const out = payload(result);
    expect(out.unknown).toEqual([TYPO]);
    expect(out.availability_view).not.toHaveProperty(TYPO);
    expect(out.warning).toContain(TYPO);
  });

  it("marks an attendee missing from the schedule response as unknown, not free", async () => {
    const built = await buildSnapshot({
      source: {
        async ownEvents() {
          return [];
        },
        // The provider returned only the first attendee; the second row vanished.
        async schedules(attendees) {
          return [{ email: attendees[0] as string, blocks: [] }];
        },
      },
      attendees: [KNOWN, TYPO],
      range: { start: START, end: END },
    });
    expect(built.busy[TYPO]).toEqual([]);
    expect(built.metadata.unknown).toEqual([TYPO]);
    expect(built.metadata.unknown).not.toContain(KNOWN);
  });
});
