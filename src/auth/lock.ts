/**
 * Advisory cross-process file lock.
 *
 * Refresh tokens are single-use: two processes redeeming the same token would
 * invalidate each other's session. Each credential cache therefore takes its own
 * lock file around read-refresh-persist. A lock older than `staleMs` is presumed
 * abandoned and taken over; a live lock blocks up to `waitMs`, then fails with an
 * actionable error instead of spinning forever.
 */
import { mkdir, open, stat, unlink } from "node:fs/promises";
import { dirname } from "node:path";
import { setTimeout as delay } from "node:timers/promises";

export const LOCK_WAIT_MS = 120_000;
export const LOCK_STALE_MS = 120_000;

/** Raised when a live lock cannot be acquired within the wait budget. */
export class LockTimeoutError extends Error {
  readonly code = "lock_timeout";
  constructor(
    readonly path: string,
    readonly ageMs: number,
    readonly waitMs: number,
    label = "lock",
  ) {
    super(`${label} ${path} is still held after ${Math.round(waitMs / 1000)}s (lock is ${Math.round(ageMs / 1000)}s old)`);
    this.name = "LockTimeoutError";
  }
}

/**
 * Take an exclusive lock file. Returns a release function that is safe to call
 * more than once. The lock stores `{pid, at}` for diagnostics only.
 */
export async function acquireFileLock(
  path: string,
  options: { waitMs?: number; staleMs?: number; label?: string } = {},
): Promise<() => Promise<void>> {
  const waitMs = options.waitMs ?? LOCK_WAIT_MS;
  const staleMs = options.staleMs ?? LOCK_STALE_MS;
  const label = options.label ?? "lock";
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const deadline = Date.now() + waitMs;
  for (;;) {
    try {
      const handle = await open(path, "wx", 0o600);
      await handle.writeFile(JSON.stringify({ pid: process.pid, at: new Date().toISOString() }));
      await handle.close();
      return async () => {
        await unlink(path).catch(() => undefined);
      };
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      const age = await stat(path).then((s) => Date.now() - s.mtimeMs).catch(() => 0);
      if (age > staleMs) {
        await unlink(path).catch(() => undefined);
        continue;
      }
      if (Date.now() > deadline) throw new LockTimeoutError(path, age, waitMs, label);
      await delay(150);
    }
  }
}
