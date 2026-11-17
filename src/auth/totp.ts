/**
 * RFC 6238 TOTP for the Microsoft Authenticator soft token.
 *
 * Defaults match what the tenant's MFA prompt expects: SHA-1, 6 digits,
 * 30-second step. Verified against the RFC's published SHA-1 vectors in
 * test/auth-totp.test.ts.
 */
import { createHash, createHmac } from "node:crypto";

const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";

export interface TotpOptions {
  digits?: number;
  step?: number;
  algorithm?: "sha1" | "sha256" | "sha512";
  /** Wall clock in ms; injectable for tests. */
  at?: number;
}

/** Decode RFC 4648 base32 (padding optional, case-insensitive). */
export function base32Decode(input: string): Buffer {
  const clean = input.toUpperCase().replace(/[^A-Z2-7]/g, "");
  let bits = 0;
  let value = 0;
  const out: number[] = [];
  for (const char of clean) {
    const index = ALPHABET.indexOf(char);
    if (index === -1) throw new Error(`invalid base32 character: ${char}`);
    value = (value << 5) | index;
    bits += 5;
    if (bits >= 8) {
      out.push((value >>> (bits - 8)) & 0xff);
      bits -= 8;
    }
  }
  return Buffer.from(out);
}

/** Current TOTP code for `secret`. */
export function totp(secret: string, options: TotpOptions = {}): string {
  const { digits = 6, step = 30, algorithm = "sha1", at = Date.now() } = options;
  const key = base32Decode(secret);
  if (key.length === 0) throw new Error("totp secret is empty");
  const counter = Math.floor(at / 1000 / step);
  const counterBuffer = Buffer.alloc(8);
  counterBuffer.writeUInt32BE(Math.floor(counter / 2 ** 32), 0);
  counterBuffer.writeUInt32BE(counter >>> 0, 4);
  const digest = createHmac(algorithm, key).update(counterBuffer).digest();
  // Dynamic truncation: the low nibble of the last byte selects the 4-byte window.
  const offset = (digest[digest.length - 1] ?? 0) & 0x0f;
  const binary =
    (((digest[offset] ?? 0) & 0x7f) << 24) |
    ((digest[offset + 1] ?? 0) << 16) |
    ((digest[offset + 2] ?? 0) << 8) |
    (digest[offset + 3] ?? 0);
  return String(binary % 10 ** digits).padStart(digits, "0");
}

/** Seconds until the current window rolls over. */
export function secondsRemaining(options: Pick<TotpOptions, "step" | "at"> = {}): number {
  const { step = 30, at = Date.now() } = options;
  return step - (Math.floor(at / 1000) % step);
}

/**
 * A code is only safe to submit with enough of its window left; otherwise the
 * form submission can land after rollover and be rejected.
 */
export function totpForSubmission(
  secret: string,
  options: TotpOptions & { minimumSecondsLeft?: number } = {},
): { code: string; secondsLeft: number } {
  const { minimumSecondsLeft = 5, ...rest } = options;
  const secondsLeft = secondsRemaining({ ...(rest.step !== undefined ? { step: rest.step } : {}), ...(rest.at !== undefined ? { at: rest.at } : {}) });
  if (secondsLeft < minimumSecondsLeft) {
    throw new Error(`totp window closes in ${secondsLeft}s; wait for the next window`);
  }
  return { code: totp(secret, rest), secondsLeft };
}

/** SHA-256 helper used by the PKCE builder (kept here so auth has one hash home). */
export function sha256Base64Url(input: string): string {
  return createHash("sha256").update(input).digest("base64url");
}

/** Published RFC 6238 SHA-1 vectors (secret "12345678901234567890"), 8 digits. */
export const RFC6238_VECTORS: ReadonlyArray<{ at: number; code: string }> = [
  { at: 59_000, code: "94287082" },
  { at: 1_111_111_109_000, code: "07081804" },
  { at: 1_111_111_111_000, code: "14050471" },
  { at: 1_234_567_890_000, code: "89005924" },
  { at: 2_000_000_000_000, code: "69279037" },
];

/** Base32 of the RFC's ASCII secret. */
export const RFC6238_SECRET = "GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ";
