/**
 * Reader for the MSAL credential cache that Outlook Web leaves in localStorage.
 *
 * OWA (MSAL.js v3) stores one entry per credential, keyed as:
 *
 *   msal.3|<homeAccountId>|login.windows.net|<credentialType>|<clientId>|<tenantId>|<scopes>
 *
 * with `credentialType` one of AccessToken | RefreshToken | IdToken | Account and
 * the value a JSON object `{ homeAccountId, environment, clientId, credentialType,
 * secret, target, expiresOn, ... }`.
 *
 * This module is deliberately pure (a plain `Record<string,string>` in, typed
 * entries out) so the harvest can be unit-tested from a redacted fixture.
 */
import { decodeJwtClaims, type RefreshClaims } from "./refresh.js";

export type MsalCredentialType = "AccessToken" | "RefreshToken" | "IdToken" | "Account" | string;

export interface MsalEntry {
  key: string;
  homeAccountId?: string;
  environment?: string;
  clientId?: string;
  credentialType: MsalCredentialType;
  secret?: string;
  target?: string;
  expiresOn?: number | string;
}

/** Split an MSAL cache key into its fields, tolerating trailing empty segments. */
export function parseMsalKey(key: string): { homeAccountId?: string; environment?: string; credentialType?: string; clientId?: string; tenantId?: string } {
  const parts = key.split("|");
  if (parts.length < 5 || !parts[0]?.startsWith("msal.")) return {};
  return {
    ...(parts[1] ? { homeAccountId: parts[1] } : {}),
    ...(parts[2] ? { environment: parts[2] } : {}),
    ...(parts[3] ? { credentialType: parts[3] } : {}),
    ...(parts[4] ? { clientId: parts[4] } : {}),
    ...(parts[5] ? { tenantId: parts[5] } : {}),
  };
}

/**
 * Parse every `msal.*` credential entry out of a localStorage-shaped map.
 * Non-credential bookkeeping keys (`*.token.keys`, `account.keys`) and any
 * unparsable value are skipped.
 */
export function parseMsalStorage(storage: Readonly<Record<string, string>>): MsalEntry[] {
  const entries: MsalEntry[] = [];
  for (const [key, raw] of Object.entries(storage)) {
    if (!key.startsWith("msal.")) continue;
    if (/\.token\.keys$|\.account\.keys$/.test(key)) continue;
    let parsed: Record<string, unknown>;
    try {
      const value = JSON.parse(raw) as unknown;
      if (value === null || typeof value !== "object") continue;
      parsed = value as Record<string, unknown>;
    } catch {
      continue;
    }
    const fromKey = parseMsalKey(key);
    const credentialType =
      typeof parsed.credentialType === "string" ? parsed.credentialType : fromKey.credentialType;
    if (!credentialType) continue;
    const entry: MsalEntry = { key, credentialType };
    if (typeof parsed.homeAccountId === "string") entry.homeAccountId = parsed.homeAccountId;
    else if (fromKey.homeAccountId) entry.homeAccountId = fromKey.homeAccountId;
    if (typeof parsed.environment === "string") entry.environment = parsed.environment;
    if (typeof parsed.clientId === "string") entry.clientId = parsed.clientId;
    else if (fromKey.clientId) entry.clientId = fromKey.clientId;
    if (typeof parsed.secret === "string") entry.secret = parsed.secret;
    if (typeof parsed.target === "string") entry.target = parsed.target;
    if (typeof parsed.expiresOn === "number" || typeof parsed.expiresOn === "string") entry.expiresOn = parsed.expiresOn;
    entries.push(entry);
  }
  return entries;
}

/** The OWA client's refresh token, newest account first (single account in practice). */
export function pickRefreshToken(entries: readonly MsalEntry[], clientId?: string): MsalEntry | undefined {
  return entries.find(
    (entry) =>
      entry.credentialType === "RefreshToken" &&
      Boolean(entry.secret) &&
      (clientId === undefined || entry.clientId === undefined || entry.clientId === clientId),
  );
}

/** Identity claims for the cache's `account` block. */
export function idTokenClaims(entries: readonly MsalEntry[]): RefreshClaims | undefined {
  const idToken = entries.find((entry) => entry.credentialType === "IdToken" && Boolean(entry.secret));
  return idToken?.secret ? decodeJwtClaims(idToken.secret) : undefined;
}

/** Access tokens currently cached by OWA, for diagnostics only. */
export function cachedAccessTokenAudiences(entries: readonly MsalEntry[]): string[] {
  const audiences = new Set<string>();
  for (const entry of entries) {
    if (entry.credentialType !== "AccessToken" || !entry.secret) continue;
    const claims = decodeJwtClaims(entry.secret);
    if (claims?.aud) audiences.add(claims.aud);
  }
  return [...audiences];
}
