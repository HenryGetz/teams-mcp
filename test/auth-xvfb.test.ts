import { describe, expect, it } from "vitest";
import { existsSync } from "node:fs";
import { pickFreeDisplay, startXvfb, xvfbAvailable, XvfbUnavailableError, existingDisplays } from "../src/auth/xvfb.js";

describe("pickFreeDisplay", () => {
  it("skips displays that already have a socket", () => {
    expect(pickFreeDisplay([99, 100, 101])).toBe(102);
  });

  it("starts at 99 so a real desktop on :0 is never reused", () => {
    expect(pickFreeDisplay([])).toBe(99);
    expect(pickFreeDisplay([0, 1])).toBe(99);
  });

  it("returns the first gap, not simply the next number after the highest", () => {
    expect(pickFreeDisplay([99, 101], 99, 130)).toBe(100);
  });

  it("raises an actionable error when the range is exhausted", () => {
    const inUse = Array.from({ length: 6 }, (_, index) => 99 + index);
    expect(() => pickFreeDisplay(inUse, 99, 104)).toThrow(XvfbUnavailableError);
    expect(() => pickFreeDisplay(inUse, 99, 104)).toThrow(/no free X display/);
  });
});

describe("existingDisplays", () => {
  it("returns numbers only", () => {
    const displays = existingDisplays();
    expect(Array.isArray(displays)).toBe(true);
    expect(displays.every((display) => Number.isInteger(display) && display >= 0)).toBe(true);
  });

  it("tolerates a missing socket directory", () => {
    expect(existingDisplays("/nonexistent-x11-dir")).toEqual([]);
  });
});

describe("startXvfb", () => {
  it("either brings up a display with a socket or fails with a typed, actionable error", async () => {
    if (!xvfbAvailable()) {
      await expect(startXvfb({ waitMs: 2000 })).rejects.toMatchObject({ code: "xvfb_unavailable" });
      const error = await startXvfb({ waitMs: 2000 }).catch((e: unknown) => e);
      expect(String((error as Error).message)).toMatch(/Xvfb is not installed/);
      expect((error as XvfbUnavailableError).hint).toMatch(/install xvfb/);
      return;
    }
    const handle = await startXvfb({ waitMs: 8000 });
    try {
      expect(handle.display).toBeGreaterThanOrEqual(99);
      expect(existsSync(`/tmp/.X11-unix/X${handle.display}`)).toBe(true);
      // The next pick must avoid the display we just took.
      expect(pickFreeDisplay(existingDisplays())).not.toBe(handle.display);
    } finally {
      handle.stop();
    }
  }, 20_000);
});
