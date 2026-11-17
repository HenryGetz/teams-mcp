import { describe, expect, it } from "vitest";
import { RFC6238_SECRET, RFC6238_VECTORS, base32Decode, secondsRemaining, totp, totpForSubmission } from "../src/auth/totp.js";

describe("TOTP (RFC 6238)", () => {
  it("matches the published SHA-1 test vectors", () => {
    for (const vector of RFC6238_VECTORS) {
      expect(totp(RFC6238_SECRET, { at: vector.at, digits: 8 })).toBe(vector.code);
    }
  });

  it("defaults to 6 digits / 30s steps", () => {
    const fixed = 1_700_000_000_000;
    const code = totp(RFC6238_SECRET, { at: fixed });
    expect(code).toMatch(/^\d{6}$/);
    // Same 30s window -> identical code; next window -> different code.
    expect(totp(RFC6238_SECRET, { at: fixed + 1000 })).toBe(code);
    expect(totp(RFC6238_SECRET, { at: fixed + 30_000 })).not.toBe(code);
  });

  it("decodes base32 with case and padding tolerance", () => {
    expect(base32Decode("gezdgnbvgy3tqojqgezdgnbvgy3tqojq").toString("utf8")).toBe("12345678901234567890");
    expect(base32Decode("JBSWY3DP").toString("utf8")).toBe("Hello");
    expect(base32Decode("MZXW6===").toString("utf8")).toBe("foo");
  });

  it("rejects an empty secret", () => {
    expect(() => totp("", { at: 0 })).toThrow(/secret is empty/);
  });

  it("computes the remaining window correctly", () => {
    expect(secondsRemaining({ at: 0 })).toBe(30);
    expect(secondsRemaining({ at: 29_000 })).toBe(1);
    expect(secondsRemaining({ at: 30_000 })).toBe(30);
  });

  it("refuses to submit a code that is about to expire", () => {
    expect(() => totpForSubmission(RFC6238_SECRET, { at: 28_000, minimumSecondsLeft: 5 })).toThrow(/wait for the next window/);
    const submission = totpForSubmission(RFC6238_SECRET, { at: 1_000, minimumSecondsLeft: 5 });
    expect(submission.secondsLeft).toBe(29);
    expect(submission.code).toMatch(/^\d{6}$/);
  });
});
