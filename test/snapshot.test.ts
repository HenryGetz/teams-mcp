import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { resetConfigCache } from "../src/config.js";
import {
  buildSnapshot,
  cacheSnapshot,
  clearSnapshotCache,
  ensureSnapshot,
  isStale,
  mergeBlocks,
  rangeCovers,
  readSnapshot,
  summarizeSnapshot,
  writeSnapshot,
  type CalendarSource,
  type CalendarSnapshot,
} from "../src/snapshot/calendar.js";

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-snapshot-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  delete process.env.TEAMS_MCP_SNAPSHOT_DIR;
  delete process.env.TEAMS_MCP_TIMEZONE;
  resetConfigCache();
  clearSnapshotCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  clearSnapshotCache();
  await rm(dir, { recursive: true, force: true });
});

const RANGE = { start: "2026-11-16T15:00:00.000Z", end: "2026-11-23T23:00:00.000Z" };

function fakeSource(overrides: Partial<CalendarSource> = {}): CalendarSource {
  return {
    ownEvents: async () => [
      { id: "e1", subject: "Standup", start: "2026-11-16T15:00:00.000Z", end: "2026-11-16T15:30:00.000Z" },
    ],
    schedules: async (attendees) =>
      attendees.map((email) => ({
        email,
        blocks: [{ start: "2026-11-16T15:00:00.000Z", end: "2026-11-16T16:00:00.000Z", status: "busy" as const }],
      })),
    ...overrides,
  };
}

describe("snapshot persistence", () => {
  it("writes compact YAML 0600 and reads it back", async () => {
    const snapshot = await buildSnapshot({
      source: fakeSource(),
      attendees: ["a@example.com"],
      range: RANGE,
      timezone: "America/Chicago",
      ttlSeconds: 1800,
      now: Date.parse("2026-11-16T12:00:00.000Z"),
    });
    const path = join(dir, "cache", "snapshots", "calendar.yaml");
    await writeSnapshot(snapshot, path);

    const info = await stat(path);
    expect(info.mode & 0o777).toBe(0o600);

    const text = await readFile(path, "utf8");
    expect(text).toContain("version: 2");
    expect(text).toContain("fetched_at: 2026-11-16T12:00:00.000Z");
    expect(text).toContain("attendees:");

    const loaded = await readSnapshot(path);
    expect(loaded?.snapshot.events).toHaveLength(1);
    expect(loaded?.snapshot.busy["a@example.com"]?.[0]?.status).toBe("busy");
  });

  it("returns null for a missing or corrupt snapshot", async () => {
    expect(await readSnapshot(join(dir, "nope.yaml"))).toBeNull();
    const bad = join(dir, "bad.yaml");
    await writeSnapshot(
      await buildSnapshot({ source: fakeSource(), attendees: [], range: RANGE }),
      bad,
    );
    const text = await readFile(bad, "utf8");
    await writeFile(bad, `${text}\nversion: 2\n`);
    expect(await readSnapshot(bad)).toBeNull();
  });

  it("merges overlapping blocks of the same status", () => {
    expect(
      mergeBlocks([
        { start: "2026-11-16T15:00:00.000Z", end: "2026-11-16T16:00:00.000Z", status: "busy" },
        { start: "2026-11-16T15:30:00.000Z", end: "2026-11-16T17:00:00.000Z", status: "busy" },
      ]),
    ).toEqual([{ start: "2026-11-16T15:00:00.000Z", end: "2026-11-16T17:00:00.000Z", status: "busy" }]);
  });
});

describe("freshness policy", () => {
  it("treats own-calendar TTL as the binding limit", async () => {
    const snapshot = await buildSnapshot({
      source: fakeSource(),
      attendees: ["a@example.com"],
      range: RANGE,
      ttlSeconds: 1800,
      now: Date.parse("2026-11-16T12:00:00.000Z"),
    });
    const sixMinutesLater = Date.parse("2026-11-16T12:06:00.000Z");
    expect(isStale(snapshot, sixMinutesLater)).toBe(true);
    expect(isStale(snapshot, sixMinutesLater, 1800)).toBe(false);
    expect(rangeCovers(snapshot, RANGE.start, RANGE.end)).toBe(true);
    expect(rangeCovers(snapshot, "2026-11-01T00:00:00.000Z", RANGE.end)).toBe(false);
  });

  it("refreshes once, then serves memory, then refreshes when stale", async () => {
    let fetches = 0;
    const source = fakeSource({
      ownEvents: async () => {
        fetches += 1;
        return [];
      },
    });
    const base = Date.parse("2026-11-16T12:00:00.000Z");

    const first = await ensureSnapshot({
      source,
      attendees: ["a@example.com"],
      ...RANGE,
      now: base,
    });
    expect(first.refreshed).toBe(true);
    expect(fetches).toBe(1);

    const second = await ensureSnapshot({ attendees: ["a@example.com"], ...RANGE, now: base + 60_000 });
    expect(second.refreshed).toBe(false);
    expect(fetches).toBe(1);

    const stale = await ensureSnapshot({
      source,
      attendees: ["a@example.com"],
      ...RANGE,
      now: base + 6 * 60_000,
    });
    expect(stale.refreshed).toBe(true);
    expect(fetches).toBe(2);
  });

  it("never fabricates a snapshot when no source is available", async () => {
    await expect(ensureSnapshot({ attendees: ["a@example.com"], ...RANGE })).rejects.toMatchObject({
      code: "snapshot_unavailable",
    });
  });
});

describe("summarizeSnapshot", () => {
  it("returns counts and local range, not full arrays", async () => {
    const snapshot: CalendarSnapshot = await buildSnapshot({
      source: fakeSource(),
      attendees: ["a@example.com"],
      range: RANGE,
      timezone: "America/Chicago",
      now: Date.parse("2026-11-16T12:00:00.000Z"),
    });
    cacheSnapshot(snapshot, "/tmp/x.yaml");
    const summary = summarizeSnapshot(snapshot, "/tmp/x.yaml", Date.parse("2026-11-16T12:01:00.000Z"));
    expect(summary.path).toBe("/tmp/x.yaml");
    expect(summary.events).toBe(1);
    expect(summary.stale).toBe(false);
    expect(JSON.stringify(summary)).not.toContain("scheduleItems");
  });
});
