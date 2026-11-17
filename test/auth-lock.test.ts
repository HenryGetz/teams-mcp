import { mkdtemp, readFile, rm, stat, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { config, resetConfigCache } from "../src/config.js";
import { acquireTokenLock, tokenLockPath } from "../src/auth/store.js";

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-lock-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  resetConfigCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("token lock", () => {
  it("lives at <configDir>/token.lock", () => {
    expect(tokenLockPath()).toBe(join(config().configDir, "token.lock"));
  });

  it("writes pid + timestamp at 0600 and releases by deleting the file", async () => {
    const release = await acquireTokenLock(1000);
    const path = tokenLockPath();
    const mode = (await stat(path)).mode & 0o777;
    expect(mode).toBe(0o600);
    const body = JSON.parse(await readFile(path, "utf8")) as { pid: number; at: string };
    expect(body.pid).toBe(process.pid);
    expect(Number.isNaN(Date.parse(body.at))).toBe(false);

    await release();
    await expect(stat(path)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("is re-entrant-safe: a released lock can be taken again", async () => {
    const first = await acquireTokenLock(1000);
    await first();
    const second = await acquireTokenLock(1000);
    await second();
  });

  it("waits, then fails with an actionable error while a live lock is held", async () => {
    const held = await acquireTokenLock(60_000);
    const started = Date.now();
    const error = await acquireTokenLock(400).catch((caught: unknown) => caught);
    expect(error).toMatchObject({ code: "auth_required" });
    expect(String((error as Error).message)).toMatch(/token lock .* still held after/);
    expect(String((error as { hint?: string }).hint)).toMatch(/delete the lock file/);
    // It really waited for its budget rather than failing instantly.
    expect(Date.now() - started).toBeGreaterThanOrEqual(350);
    await held();
  });

  it("takes over a stale lock instead of blocking", async () => {
    const path = tokenLockPath();
    await writeFile(path, JSON.stringify({ pid: 999999, at: new Date(0).toISOString() }), { mode: 0o600 });
    // Backdate well past the 120s stale threshold.
    const old = new Date(Date.now() - 200_000);
    await utimes(path, old, old);

    const release = await acquireTokenLock(500);
    const body = JSON.parse(await readFile(path, "utf8")) as { pid: number };
    expect(body.pid).toBe(process.pid);
    await release();
  });
});
