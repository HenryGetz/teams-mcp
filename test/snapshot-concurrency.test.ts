/**
 * Snapshot write/refresh safety.
 *
 * Live failure this locks down: two availability calls in one process shared
 * `<snapshot>.<pid>.tmp`, so the second `rename` lost its source and surfaced
 * `ENOENT: no such file or directory, rename '…calendar.yaml.<pid>.tmp'`.
 * The fixes: unique temp names, one refresh at a time per path, retry once on
 * ENOENT, stale-temp cleanup, and an atomic replace that never drops the prior
 * file.
 */
import { mkdtemp, readdir, readFile, rm, stat, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { resetConfigCache } from "../src/config.js";
import {
  buildSnapshot,
  clearSnapshotCache,
  ensureSnapshot,
  readSnapshot,
  snapshotPath,
  SNAPSHOT_VERSION,
  writeSnapshot,
  type CalendarSnapshot,
  type CalendarSource,
} from "../src/snapshot/calendar.js";

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-snapshot-race-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  resetConfigCache();
  clearSnapshotCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  clearSnapshotCache();
  await rm(dir, { recursive: true, force: true });
});

function snapshot(overrides: Partial<CalendarSnapshot> = {}): CalendarSnapshot {
  const fetchedAt = new Date().toISOString();
  return {
    version: SNAPSHOT_VERSION,
    metadata: {
      fetched_at: fetchedAt,
      range_start: "2026-11-17T15:00:00.000Z",
      range_end: "2026-11-17T18:00:00.000Z",
      timezone: "America/Chicago",
      ttl_seconds: 3600,
      stale_after: new Date(Date.parse(fetchedAt) + 3_600_000).toISOString(),
      source: "outlook",
      attendees: ["student@example.edu"],
      unknown: [],
    },
    events: [],
    busy: {
      "student@example.edu": [
        { start: "2026-11-17T16:00:00.000Z", end: "2026-11-17T16:30:00.000Z", status: "busy" },
      ],
    },
    ...overrides,
  };
}

/**
 * Source that counts builds and parks them on a gate, so overlap is guaranteed
 * without timing guesses: the caller opens the gate only after every request is
 * already waiting.
 */
function gatedSource(): { source: CalendarSource; builds: () => number; open: () => void } {
  let builds = 0;
  const gate = Promise.withResolvers<void>();
  return {
    builds: () => builds,
    open: gate.resolve,
    source: {
      async ownEvents() {
        builds += 1;
        await gate.promise;
        return [];
      },
      async schedules(attendees) {
        await gate.promise;
        return attendees.map((email) => ({
          email,
          blocks: [{ start: "2026-11-17T16:00:00.000Z", end: "2026-11-17T16:30:00.000Z", status: "busy" as const }],
        }));
      },
    },
  };
}

const options = {
  start: "2026-11-17T15:00:00.000Z",
  end: "2026-11-17T18:00:00.000Z",
  attendees: ["student@example.edu"],
  timezone: "America/Chicago",
};

describe("ensureSnapshot concurrency", () => {
  it("builds once for parallel identical calls and writes no temp litter", async () => {
    const gated = gatedSource();
    const pending = Promise.all(Array.from({ length: 6 }, () => ensureSnapshot({ ...options, source: gated.source })));
    gated.open();
    const results = await pending;
    expect(gated.builds()).toBe(1);
    expect(new Set(results.map((result) => result.snapshot.metadata.fetched_at)).size).toBe(1);
    expect(results.every((result) => result.refreshed || result.snapshot === results[0]?.snapshot)).toBe(true);

    const loaded = await readSnapshot();
    expect(loaded?.snapshot.busy["student@example.edu"]).toHaveLength(1);
    const leftovers = (await readdir(join(dir, "cache", "snapshots"))).filter((name) => name.endsWith(".tmp"));
    expect(leftovers).toEqual([]);
  });

  it("serialises concurrent refreshes for different ranges without losing either write", async () => {
    const gated = gatedSource();
    const pending = Promise.all([
      ensureSnapshot({ ...options, source: gated.source }),
      ensureSnapshot({
        ...options,
        start: "2026-11-18T15:00:00.000Z",
        end: "2026-11-18T18:00:00.000Z",
        source: gated.source,
      }),
    ]);
    gated.open();
    const [a, b] = await pending;
    expect(a.snapshot.metadata.range_start).toBe("2026-11-17T15:00:00.000Z");
    expect(b.snapshot.metadata.range_start).toBe("2026-11-18T15:00:00.000Z");

    // The file is a complete document; the later build simply won the rename.
    const loaded = await readSnapshot();
    expect(loaded).not.toBeNull();
    expect(loaded?.snapshot.metadata.range_start).toBe("2026-11-18T15:00:00.000Z");
    const leftovers = (await readdir(join(dir, "cache", "snapshots"))).filter((name) => name.endsWith(".tmp"));
    expect(leftovers).toEqual([]);
  });

  it("honours a forced refresh once for parallel callers", async () => {
    const gated = gatedSource();
    const initial = ensureSnapshot({ ...options, source: gated.source });
    gated.open();
    await initial;
    expect(gated.builds()).toBe(1);

    const forced = Promise.all(
      Array.from({ length: 4 }, () =>
        ensureSnapshot({ ...options, source: gated.source, refresh: true }),
      ),
    );
    const results = await forced;
    expect(gated.builds()).toBe(2); // one initial + one forced, not one per caller
    expect(results.every((result) => result.snapshot.metadata.range_start === options.start)).toBe(true);
  });
});

describe("writeSnapshot", () => {
  it("writes concurrent snapshots with unique temps and a valid final file", async () => {
    const path = snapshotPath();
    const writes = Array.from({ length: 12 }, (_, index) =>
      writeSnapshot(
        snapshot({
          events: [
            {
              id: `e-${index}`,
              subject: `Event ${index}`,
              start: "2026-11-17T16:00:00.000Z",
              end: "2026-11-17T16:30:00.000Z",
            },
          ],
        }),
      ),
    );
    const paths = await Promise.all(writes);
    expect(new Set(paths).size).toBe(1);
    const loaded = await readSnapshot(path);
    expect(loaded?.snapshot.events).toHaveLength(1);
    expect((await stat(path)).mode & 0o777).toBe(0o600);
    const leftovers = (await readdir(join(dir, "cache", "snapshots"))).filter((name) => name.endsWith(".tmp"));
    expect(leftovers).toEqual([]);
  });

  it("keeps the previous snapshot when a new write fails validation", async () => {
    const path = snapshotPath();
    await writeSnapshot(snapshot());
    const broken = { version: SNAPSHOT_VERSION, metadata: {}, events: [], busy: {} };
    // @ts-expect-error deliberately invalid: metadata is required
    await expect(writeSnapshot(broken)).rejects.toThrow();
    const loaded = await readSnapshot(path);
    expect(loaded?.snapshot.metadata.range_start).toBe("2026-11-17T15:00:00.000Z");
  });

  it("cleans abandoned temps older than an hour and leaves fresh ones alone", async () => {
    const path = snapshotPath();
    await writeSnapshot(snapshot());
    const dirPath = join(dir, "cache", "snapshots");
    const stale = join(dirPath, "calendar.yaml.999.deadbeef.tmp");
    const fresh = join(dirPath, "calendar.yaml.999.cafebabe.tmp");
    await writeFile(stale, "stale");
    await writeFile(fresh, "fresh");
    const old = new Date(Date.now() - 2 * 60 * 60 * 1000);
    await utimes(stale, old, old);

    await writeSnapshot(snapshot());
    const names = await readdir(dirPath);
    expect(names).not.toContain("calendar.yaml.999.deadbeef.tmp");
    expect(names).toContain("calendar.yaml.999.cafebabe.tmp");
    expect((await readFile(path, "utf8")).startsWith("# teams-mcp calendar snapshot")).toBe(true);
  });
});

describe("buildSnapshot busy hygiene", () => {
  it("drops free intervals and keeps the rest", async () => {
    const source: CalendarSource = {
      async ownEvents() {
        return [];
      },
      async schedules(attendees) {
        return [
          {
            email: attendees[0] as string,
            blocks: [
              { start: "2026-11-17T09:00:00.000Z", end: "2026-11-17T09:30:00.000Z", status: "free" },
              { start: "2026-11-17T10:00:00.000Z", end: "2026-11-17T10:30:00.000Z", status: "busy" },
              { start: "2026-11-17T11:00:00.000Z", end: "2026-11-17T11:30:00.000Z", status: "oof" },
            ],
          },
        ];
      },
    };
    const built = await buildSnapshot({
      source,
      attendees: ["a@example.com"],
      range: { start: "2026-11-17T08:00:00.000Z", end: "2026-11-17T18:00:00.000Z" },
    });
    expect(built.busy["a@example.com"]?.map((block) => block.status)).toEqual(["busy", "oof"]);
    expect(built.metadata.unknown).toEqual([]);
  });

  it("records unresolved mailboxes in metadata.unknown", async () => {
    const source: CalendarSource = {
      async ownEvents() {
        return [];
      },
      async schedules(attendees) {
        return [
          { email: attendees[0] as string, blocks: [] },
          { email: attendees[1] as string, blocks: [], unknown: true },
        ];
      },
    };
    const built = await buildSnapshot({
      source,
      attendees: ["known@example.com", "typo@example.com"],
      range: { start: "2026-11-17T08:00:00.000Z", end: "2026-11-17T18:00:00.000Z" },
    });
    expect(built.metadata.unknown).toEqual(["typo@example.com"]);
    expect(built.busy["typo@example.com"]).toEqual([]);
  });
});
