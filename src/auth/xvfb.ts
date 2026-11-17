/**
 * Xvfb management for the headed Chromium login.
 *
 * Headless Chromium is detectable by the Microsoft login front end, so the flow
 * runs headed. On a machine with no display we bring up our own X server on a
 * free display, then tear it down with the login.
 */
import { spawn, spawnSync, type ChildProcess } from "node:child_process";
import { existsSync } from "node:fs";
import { setTimeout as delay } from "node:timers/promises";

export const XVFB_SOCKET_DIR = "/tmp/.X11-unix";
export const DEFAULT_SCREEN = "1280x900x24";

export class XvfbUnavailableError extends Error {
  readonly code = "xvfb_unavailable";
  constructor(message: string, readonly hint?: string) {
    super(message);
    this.name = "XvfbUnavailableError";
  }
}

/** Displays with a live socket, e.g. `:0`, `:99`. */
export function existingDisplays(socketDir = XVFB_SOCKET_DIR): number[] {
  if (!existsSync(socketDir)) return [];
  const probe = spawnSync("sh", ["-c", `ls -1 ${JSON.stringify(socketDir)} 2>/dev/null`], { encoding: "utf8" });
  if (probe.status !== 0 || typeof probe.stdout !== "string") return [];
  return probe.stdout
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => /^X\d+$/.test(line))
    .map((line) => Number(line.slice(1)))
    .filter((display) => Number.isFinite(display));
}

/**
 * First display number in [from, to] with no existing socket. Starting at 99
 * avoids colliding with a real desktop session on :0.
 */
export function pickFreeDisplay(inUse: readonly number[], from = 99, to = 130): number {
  const taken = new Set(inUse);
  for (let display = from; display <= to; display += 1) {
    if (!taken.has(display)) return display;
  }
  throw new XvfbUnavailableError(
    `no free X display in ${from}..${to}`,
    "stop stale Xvfb processes or set DISPLAY to an existing server",
  );
}

export function xvfbAvailable(): boolean {
  return spawnSync("sh", ["-c", "command -v Xvfb"], { encoding: "utf8" }).status === 0;
}

export interface XvfbHandle {
  display: number;
  process: ChildProcess;
  stop(): void;
}

/**
 * Start Xvfb on a free display direct to an `Xvfb :N` call and wait for its
 * socket. Throws `XvfbUnavailableError` with an actionable hint if Xvfb is not
 * installed or fails to come up.
 */
export async function startXvfb(options: { screen?: string; waitMs?: number } = {}): Promise<XvfbHandle> {
  const { screen = DEFAULT_SCREEN, waitMs = 8000 } = options;
  if (!xvfbAvailable()) {
    throw new XvfbUnavailableError(
      "Xvfb is not installed and DISPLAY is unset",
      "install xvfb (Debian/Ubuntu: apt-get install xvfb) or run on a machine with a display; the login window must be headed",
    );
  }
  const display = pickFreeDisplay(existingDisplays());
  const child = spawn("Xvfb", [`:${display}`, "-screen", "0", screen, "-nolisten", "tcp"], {
    stdio: "ignore",
    detached: false,
  });

  const socket = `${XVFB_SOCKET_DIR}/X${display}`;
  const deadline = Date.now() + waitMs;
  while (Date.now() < deadline) {
    if (existsSync(socket)) {
      const handle: XvfbHandle = {
        display,
        process: child,
        stop() {
          if (!child.killed) child.kill("SIGTERM");
        },
      };
      return handle;
    }
    if (child.exitCode !== null) {
      throw new XvfbUnavailableError(
        `Xvfb exited immediately on :${display} (code ${child.exitCode})`,
        "run Xvfb manually to see the error; another session may own that display",
      );
    }
    await delay(100);
  }
  child.kill("SIGTERM");
  throw new XvfbUnavailableError(
    `Xvfb did not create ${socket} within ${waitMs}ms`,
    "check for a stale lock file /tmp/.X${display}-lock",
  );
}

export interface DisplayEnvironment {
  display: string;
  /** Set when this call started Xvfb (and therefore owns teardown). */
  xvfb: XvfbHandle | null;
}

/**
 * Ensure `process.env.DISPLAY` points at something usable. Starts Xvfb only when
 * no display is configured; a pre-existing DISPLAY is left alone.
 */
export async function ensureDisplay(options: { screen?: string } = {}): Promise<DisplayEnvironment> {
  const existing = process.env.DISPLAY?.trim();
  if (existing) return { display: existing, xvfb: null };
  const xvfb = await startXvfb(options);
  process.env.DISPLAY = `:${xvfb.display}`;
  return { display: process.env.DISPLAY, xvfb };
}
