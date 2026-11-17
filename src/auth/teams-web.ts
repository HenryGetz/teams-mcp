/**
 * Teams-web credential: the Graph token minted by the Teams web client.
 *
 * The OWA-derived Graph token has no `ChannelMessage.Read.All` (docs/teams-recon.md).
 * The Teams web app (client `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`) mints a Graph
 * token with 37 scopes, including channel messages, channel/team members and mail.
 * This module owns that credential:
 *
 *   ~/.config/teams-mcp/teams-web-token.json (0600)
 *   {version, client_id, tenant_id, account, scope, access_token, expires_at,
 *    refresh_token?, source, updated_at}
 *
 * Refresh: the standard refresh-token grant with the Teams `Origin`, rotated and
 * persisted atomically before the access token is used (own lock file, so it
 * never contends with the OWA token lock).
 *
 * Acquisition fallback (fallbackTeamsWebAuth): drive the persistent Chromium
 * profile to `https://teams.cloud.microsoft/`. The SSO session makes that
 * silent (no prompt). MSAL v5 stores its refresh token encrypted; the key is the
 * `msal.cache.encryption` cookie and the scheme is documented in
 * docs/teams-web-auth.md §3.1. If decryption is impossible we persist an
 * **access-only** credential from the cached Graph access token, which still
 * works until it expires.
 */
import { chmod, mkdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { dirname } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import type { Page } from "playwright";
import { config, TOKEN_EXPIRY_SKEW_MS } from "../config.js";
import { decodeJwtClaims, refreshAccessToken, RefreshError } from "./refresh.js";
import { acquireFileLock, LockTimeoutError } from "./lock.js";
import { driveSignIn, openBrowserSession, readCredentials, type LogFn } from "./login.js";
import { AuthRequiredError, type TokenSource } from "./store.js";

export const TEAMS_WEB_CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
export const TEAMS_WEB_ORIGIN = "https://teams.microsoft.com";
export const TEAMS_WEB_SCOPE = "https://graph.microsoft.com/.default";
export const TEAMS_WEB_URL = "https://teams.cloud.microsoft/";
/** The one scope that makes this credential worth a second cache. */
export const TEAMS_WEB_REQUIRED_SCOPE = "ChannelMessage.Read.All";

export const TEAMS_WEB_CACHE_VERSION = 1 as const;

export interface TeamsWebCache {
  version: typeof TEAMS_WEB_CACHE_VERSION;
  client_id: string;
  tenant_id?: string;
  account?: { upn?: string; id?: string };
  scope: string;
  access_token: string;
  /** ISO 8601 UTC. */
  expires_at: string;
  refresh_token?: string;
  source: TokenSource | "access-only";
  updated_at: string;
}

export interface TeamsWebAccessToken {
  token: string;
  expiresAt: string;
  source: TeamsWebCache["source"];
  /** True when a refresh token is stored, so no browser is needed next time. */
  refreshable: boolean;
}

export type TeamsWebCacheRead =
  | { cache: TeamsWebCache; path: string }
  | { cache: null; path: string; reason: string; corrupt?: true };

const stderrLog: LogFn = (message) => process.stderr.write(`${message}\n`);

function isSource(value: unknown): value is TeamsWebCache["source"] {
  return value === "headless" || value === "browser" || value === "refresh" || value === "access-only";
}

function validateCache(raw: unknown): TeamsWebCache {
  if (raw === null || typeof raw !== "object") throw new Error("teams-web cache is not an object");
  const doc = raw as Record<string, unknown>;
  if (doc.version !== TEAMS_WEB_CACHE_VERSION) {
    // Never echo the field: a corrupt document must not put its bytes in an error.
    throw new Error("unsupported teams-web cache version");
  }
  if (typeof doc.access_token !== "string" || doc.access_token === "") {
    throw new Error("teams-web cache has no access_token");
  }
  if (typeof doc.expires_at !== "string" || Number.isNaN(Date.parse(doc.expires_at))) {
    throw new Error("teams-web cache has no valid expires_at");
  }
  if (doc.refresh_token !== undefined && typeof doc.refresh_token !== "string") {
    throw new Error("teams-web cache refresh_token must be a string");
  }
  if (!isSource(doc.source)) throw new Error("teams-web cache has an invalid source");
  const account = doc.account;
  if (account !== undefined && (account === null || typeof account !== "object")) {
    throw new Error("teams-web cache account must be an object");
  }
  return {
    version: TEAMS_WEB_CACHE_VERSION,
    client_id: typeof doc.client_id === "string" ? doc.client_id : TEAMS_WEB_CLIENT_ID,
    ...(typeof doc.tenant_id === "string" ? { tenant_id: doc.tenant_id } : {}),
    ...(account ? { account: account as { upn?: string; id?: string } } : {}),
    scope: typeof doc.scope === "string" && doc.scope !== "" ? doc.scope : TEAMS_WEB_SCOPE,
    access_token: doc.access_token,
    expires_at: new Date(doc.expires_at).toISOString(),
    ...(typeof doc.refresh_token === "string" && doc.refresh_token !== ""
      ? { refresh_token: doc.refresh_token }
      : {}),
    source: doc.source,
    updated_at:
      typeof doc.updated_at === "string" && !Number.isNaN(Date.parse(doc.updated_at))
        ? new Date(doc.updated_at).toISOString()
        : new Date(0).toISOString(),
  };
}

/** Offline read; a missing file is not an error. */
export async function readTeamsWebCache(path = config().teamsWebTokenPath): Promise<TeamsWebCacheRead> {
  let text: string;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    return { cache: null, path, reason: code === "ENOENT" ? "missing" : `unreadable (${code ?? "error"})` };
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text) as unknown;
  } catch {
    // `JSON.parse` errors quote the input on some Node versions; the cache holds
    // bearer material, so the reason is deliberately generic.
    return { cache: null, path, reason: "invalid: not valid JSON", corrupt: true };
  }
  try {
    return { cache: validateCache(parsed), path };
  } catch (error) {
    return { cache: null, path, reason: `invalid: ${(error as Error).message}`, corrupt: true };
  }
}

/** Atomic owner-only write; the caller must already hold the refresh token. */
export async function writeTeamsWebCache(
  cache: TeamsWebCache,
  path = config().teamsWebTokenPath,
): Promise<void> {
  const doc = validateCache({ ...cache, version: TEAMS_WEB_CACHE_VERSION });
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const tmp = `${path}.${process.pid}.${randomUUID()}.tmp`;
  await writeFile(tmp, `${JSON.stringify(doc, null, 2)}\n`, { mode: 0o600 });
  await rename(tmp, path);
  await chmod(path, 0o600).catch(() => undefined);
}

export function isTeamsWebTokenFresh(expiresAt: string, now = Date.now()): boolean {
  const expiry = Date.parse(expiresAt);
  return !Number.isNaN(expiry) && expiry - TOKEN_EXPIRY_SKEW_MS > now;
}

export function teamsWebLockPath(): string {
  return `${config().teamsWebTokenPath}.lock`;
}

/* ------------------------------------------------------------------ */
/* MSAL v5 cache decryption                                            */
/* ------------------------------------------------------------------ */

export interface MsalEncryptionCookie {
  id: string;
  key: string;
}

export interface MsalEncryptedEntry {
  id: string;
  nonce: string;
  data: string;
}

export interface MsalCredentialEntry {
  key: string;
  credentialType: string;
  secret?: string;
  target?: string;
  expiresOn?: number;
}

function base64UrlToBytes(value: string): Uint8Array<ArrayBuffer> {
  const normalised = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalised + "=".repeat((4 - (normalised.length % 4)) % 4);
  const decoded = Buffer.from(padded, "base64");
  const bytes = new Uint8Array(decoded.length);
  bytes.set(decoded);
  return bytes;
}

/**
 * Decrypt one MSAL v5 cache entry.
 *
 *   derived = HKDF(baseKey, salt = nonce, info = clientId, SHA-256) -> AES-256-GCM
 *   plain   = AES-GCM(derived, iv = 12 zero bytes, ciphertext = data)
 *
 * `cookie.key` is base64url, `nonce` and `data` are base64url. Exported so the
 * whole scheme is testable without a browser (docs/teams-web-auth.md §3.1).
 */
export async function decryptMsalCacheEntry(
  entry: MsalEncryptedEntry,
  cookie: MsalEncryptionCookie,
  clientId: string,
): Promise<Record<string, unknown>> {
  if (entry.id !== cookie.id) {
    throw new Error("MSAL encryption key id does not match the cache entry");
  }
  const subtle = globalThis.crypto.subtle;
  const baseKey = base64UrlToBytes(cookie.key);
  const hkdf = await subtle.importKey("raw", baseKey, "HKDF", false, ["deriveKey"]);
  const derived = await subtle.deriveKey(
    {
      name: "HKDF",
      salt: base64UrlToBytes(entry.nonce),
      hash: "SHA-256",
      info: new TextEncoder().encode(clientId),
    },
    hkdf,
    { name: "AES-GCM", length: 256 },
    false,
    ["decrypt"],
  );
  const plain = await subtle.decrypt(
    { name: "AES-GCM", iv: new Uint8Array(12) },
    derived,
    base64UrlToBytes(entry.data),
  );
  const parsed = JSON.parse(new TextDecoder().decode(plain)) as unknown;
  if (parsed === null || typeof parsed !== "object") throw new Error("decrypted MSAL entry is not an object");
  return parsed as Record<string, unknown>;
}

/** Parse `msal.*` credentials from a localStorage-shaped map (v2 keys included). */
export function parseTeamsMsalStorage(storage: Readonly<Record<string, string>>): MsalCredentialEntry[] {
  const entries: MsalCredentialEntry[] = [];
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
    const segments = key.split("|");
    const fromKey = segments[3];
    const credentialType =
      typeof parsed.credentialType === "string"
        ? parsed.credentialType
        : typeof fromKey === "string" && fromKey !== ""
          ? fromKey
          : "";
    if (!credentialType) continue;
    const entry: MsalCredentialEntry = { key, credentialType };
    if (typeof parsed.secret === "string") entry.secret = parsed.secret;
    if (typeof parsed.target === "string") entry.target = parsed.target;
    if (typeof parsed.expiresOn === "number") entry.expiresOn = parsed.expiresOn;
    entries.push(entry);
  }
  return entries;
}

/** `msal.cache.encryption` cookie: `{id, key}` (key is base64url). */
export function parseEncryptionCookie(cookieHeader: string): MsalEncryptionCookie | undefined {
  for (const part of cookieHeader.split(";")) {
    const [name, ...rest] = part.trim().split("=");
    if (name !== "msal.cache.encryption") continue;
    try {
      const parsed = JSON.parse(decodeURIComponent(rest.join("="))) as Record<string, unknown>;
      if (typeof parsed.id === "string" && typeof parsed.key === "string") {
        return { id: parsed.id, key: parsed.key };
      }
    } catch {
      return undefined;
    }
  }
  return undefined;
}

function homeAccountPart(key: string): { oid?: string; tid?: string } {
  const home = key.split("|")[1] ?? "";
  const [oid, tid] = home.split(".");
  return { ...(oid ? { oid } : {}), ...(tid ? { tid } : {}) };
}

export interface HarvestedCredential {
  refreshToken?: string;
  /** Cached Graph access token, when the web app already had one. */
  accessToken?: string;
  expiresAt?: string;
  tenantId?: string;
  upn?: string;
  oid?: string;
}

/** Turn the page's localStorage + cookies into a credential (pure apart from WebCrypto). */
export async function extractTeamsWebCredential(input: {
  storage: Readonly<Record<string, string>>;
  cookieHeader: string;
  clientId?: string;
}): Promise<HarvestedCredential> {
  const clientId = input.clientId ?? TEAMS_WEB_CLIENT_ID;
  const entries = parseTeamsMsalStorage(input.storage);
  const cookie = parseEncryptionCookie(input.cookieHeader);

  const refreshKey = Object.keys(input.storage).find(
    (key) => key.includes("|refreshtoken|") && key.includes(clientId),
  );
  const graphEntry = entries.find((entry) => {
    if (entry.credentialType !== "AccessToken" || !entry.secret) return false;
    const claims = decodeJwtClaims(entry.secret);
    return (
      claims?.aud === "https://graph.microsoft.com" &&
      String(claims.scp ?? "").includes(TEAMS_WEB_REQUIRED_SCOPE)
    );
  });
  const idEntry = entries.find((entry) => entry.credentialType === "IdToken" && entry.secret);
  const idClaims = idEntry?.secret ? decodeJwtClaims(idEntry.secret) : undefined;
  const graphClaims = graphEntry?.secret ? decodeJwtClaims(graphEntry.secret) : undefined;

  let refreshToken: string | undefined;
  let tenantId = graphClaims?.tid ?? idClaims?.tid;
  let oid = idClaims?.oid ?? graphClaims?.oid;

  if (refreshKey) {
    const parts = homeAccountPart(refreshKey);
    tenantId ??= parts.tid;
    oid ??= parts.oid;
    try {
      const blob = JSON.parse(input.storage[refreshKey] ?? "") as Record<string, unknown>;
      if (typeof blob.credentialType === "string" && typeof blob.secret === "string") {
        refreshToken = blob.secret; // plaintext cache (older MSAL layout)
      } else if (cookie && typeof blob.nonce === "string" && typeof blob.data === "string") {
        const decrypted = await decryptMsalCacheEntry(
          { id: typeof blob.id === "string" ? blob.id : "", nonce: blob.nonce, data: blob.data },
          cookie,
          clientId,
        );
        if (typeof decrypted.secret === "string") refreshToken = decrypted.secret;
        if (typeof decrypted.homeAccountId === "string") {
          const [homeOid, homeTid] = decrypted.homeAccountId.split(".");
          oid ??= homeOid;
          tenantId ??= homeTid;
        }
      }
    } catch {
      // Key-id mismatch / corruption: fall back to an access-only credential.
      refreshToken = undefined;
    }
  }

  const upn =
    (typeof idClaims?.preferred_username === "string" ? idClaims.preferred_username : undefined) ??
    (typeof idClaims?.upn === "string" ? idClaims.upn : undefined);

  return {
    ...(refreshToken ? { refreshToken } : {}),
    ...(graphEntry?.secret
      ? {
          accessToken: graphEntry.secret,
          ...(graphClaims?.exp ? { expiresAt: new Date(graphClaims.exp * 1000).toISOString() } : {}),
        }
      : {}),
    ...(tenantId ? { tenantId } : {}),
    ...(upn ? { upn } : {}),
    ...(oid ? { oid } : {}),
  };
}

/* ------------------------------------------------------------------ */
/* Minting and persistence                                             */
/* ------------------------------------------------------------------ */

function toAccessToken(cache: TeamsWebCache): TeamsWebAccessToken {
  return {
    token: cache.access_token,
    expiresAt: cache.expires_at,
    source: cache.source,
    refreshable: Boolean(cache.refresh_token),
  };
}

/** Persist a fresh access token (and any rotated refresh token) atomically. */
async function persist(
  base: Partial<TeamsWebCache>,
  result: { accessToken: string; expiresAt: string; scope: string; refreshToken?: string },
  source: TeamsWebCache["source"],
): Promise<TeamsWebCache> {
  const cache: TeamsWebCache = {
    version: TEAMS_WEB_CACHE_VERSION,
    client_id: base.client_id ?? TEAMS_WEB_CLIENT_ID,
    ...(base.tenant_id ? { tenant_id: base.tenant_id } : {}),
    ...(base.account ? { account: base.account } : {}),
    scope: result.scope || TEAMS_WEB_SCOPE,
    access_token: result.accessToken,
    expires_at: result.expiresAt,
    ...(result.refreshToken ?? base.refresh_token
      ? { refresh_token: (result.refreshToken ?? base.refresh_token) as string }
      : {}),
    source,
    updated_at: new Date().toISOString(),
  };
  // Rotation is persisted BEFORE the caller can use the access token.
  await writeTeamsWebCache(cache);
  return cache;
}

/** Redeem the stored refresh token, holding the dedicated lock. */
async function refreshTeamsWeb(request: { cache: TeamsWebCache; force?: boolean }): Promise<TeamsWebCache> {
  let release: (() => Promise<void>) | undefined;
  try {
    release = await acquireFileLock(teamsWebLockPath(), { label: "teams-web token lock" });
  } catch (error) {
    if (error instanceof LockTimeoutError) {
      throw new AuthRequiredError(error.message, "another process is refreshing the Teams web token; retry shortly");
    }
    throw error;
  }
  try {
    const latest = (await readTeamsWebCache()).cache ?? request.cache;
    if (!request.force && isTeamsWebTokenFresh(latest.expires_at)) return latest;
    const refreshToken = latest.refresh_token;
    if (!refreshToken) {
      throw new RefreshError({ status: 0, code: "no_refresh_token", message: "teams-web cache has no refresh token" });
    }
    const result = await refreshAccessToken({
      tenantId: latest.tenant_id ?? "common",
      refreshToken,
      audience: "graph",
      clientId: TEAMS_WEB_CLIENT_ID,
      origin: TEAMS_WEB_ORIGIN,
      scope: TEAMS_WEB_SCOPE,
    });
    return await persist(latest, result, "refresh");
  } finally {
    await release();
  }
}

/* ------------------------------------------------------------------ */
/* Browser harvest (fallback)                                          */
/* ------------------------------------------------------------------ */

/** Signals the app shell has booted enough for MSAL to hold a credential. */
async function teamsReady(page: Page): Promise<boolean> {
  return page
    .evaluate((clientId: string) => {
      if (!/teams\.(microsoft|cloud)\.microsoft/.test(location.host)) return false;
      for (let index = 0; index < localStorage.length; index += 1) {
        const key = localStorage.key(index);
        if (!key || !key.startsWith("msal.")) continue;
        if (key.includes("|refreshtoken|") && key.includes(clientId)) return true;
        if (!key.includes("|accesstoken|") || !key.includes(clientId)) continue;
        const value = localStorage.getItem(key);
        if (value && value.includes("graph.microsoft.com")) return true;
      }
      return false;
    }, TEAMS_WEB_CLIENT_ID)
    .catch(() => false);
}

async function harvestTeamsWeb(options: { log?: LogFn; timeoutMs?: number } = {}): Promise<TeamsWebCache> {
  const log = options.log ?? stderrLog;
  const credentials = await readCredentials();
  const session = await openBrowserSession({ profileDir: config().profileDir }, log);
  let storage: { values: Record<string, string>; cookie: string };
  try {
    await session.page
      .goto(TEAMS_WEB_URL, { waitUntil: "domcontentloaded", timeout: 90_000 })
      .catch((error: unknown) => log(`[teams-auth] navigation: ${error instanceof Error ? error.message : String(error)}`));
    await driveSignIn(session.page, credentials, {
      ready: async () => teamsReady(session.page),
      timeoutMs: options.timeoutMs ?? 150_000,
      maxSteps: 30,
      stepIntervalMs: 1500,
      log,
    });
    await delay(2000);
    storage = await session.page.evaluate(() => {
      const values: Record<string, string> = {};
      for (let index = 0; index < localStorage.length; index += 1) {
        const key = localStorage.key(index);
        if (key === null) continue;
        const value = localStorage.getItem(key);
        if (value !== null) values[key] = value;
      }
      return { values, cookie: document.cookie };
    });
  } finally {
    await session.close();
  }

  const harvested = await extractTeamsWebCredential({
    storage: storage.values,
    cookieHeader: storage.cookie,
  });
  const account = {
    ...(harvested.upn ? { upn: harvested.upn } : {}),
    ...(harvested.oid ? { id: harvested.oid } : {}),
  };
  const base: Partial<TeamsWebCache> = {
    client_id: TEAMS_WEB_CLIENT_ID,
    ...(harvested.tenantId ? { tenant_id: harvested.tenantId } : {}),
    ...(Object.keys(account).length > 0 ? { account } : {}),
  };

  if (harvested.refreshToken) {
    const result = await refreshAccessToken({
      tenantId: harvested.tenantId ?? "common",
      refreshToken: harvested.refreshToken,
      audience: "graph",
      clientId: TEAMS_WEB_CLIENT_ID,
      origin: TEAMS_WEB_ORIGIN,
      scope: TEAMS_WEB_SCOPE,
    });
    log(`[teams-auth] harvested refresh token; minted Graph token (${result.expiresIn}s)`);
    return persist({ ...base, ...(harvested.refreshToken ? { refresh_token: harvested.refreshToken } : {}) }, result, "browser");
  }

  if (harvested.accessToken && harvested.expiresAt && String(decodeJwtClaims(harvested.accessToken)?.scp ?? "").includes(TEAMS_WEB_REQUIRED_SCOPE)) {
    log("[teams-auth] could not decrypt the refresh token; storing the cached Graph access token only");
    const cache: TeamsWebCache = {
      version: TEAMS_WEB_CACHE_VERSION,
      client_id: TEAMS_WEB_CLIENT_ID,
      ...(harvested.tenantId ? { tenant_id: harvested.tenantId } : {}),
      ...(Object.keys(account).length > 0 ? { account } : {}),
      scope: TEAMS_WEB_SCOPE,
      access_token: harvested.accessToken,
      expires_at: harvested.expiresAt,
      source: "access-only",
      updated_at: new Date().toISOString(),
    };
    await writeTeamsWebCache(cache);
    return cache;
  }

  throw new AuthRequiredError(
    "no Teams web credential could be harvested (no decryptable refresh token and no cached Graph token)",
    "open https://teams.cloud.microsoft/ once in the browser profile so the app caches its tokens, then retry",
  );
}

/* ------------------------------------------------------------------ */
/* Public entry points                                                 */
/* ------------------------------------------------------------------ */

export interface TeamsWebDeps {
  refresh?: (request: { cache: TeamsWebCache; force?: boolean }) => Promise<TeamsWebCache>;
  harvest?: (options?: { log?: LogFn; timeoutMs?: number }) => Promise<TeamsWebCache>;
  now?: () => number;
}

export interface GetTeamsWebTokenOptions {
  now?: number;
  force?: boolean;
  /** False disables the browser fallback (tests, offline status checks). */
  allowBrowser?: boolean;
  log?: LogFn;
  deps?: TeamsWebDeps;
}

/**
 * Return a usable Teams-web Graph token.
 *
 * Order: fresh cache -> refresh (rotation persisted first) -> browser harvest ->
 * error. A transient refresh failure never launches Chromium; only a dead or
 * missing refresh token does (and an access-only cache still answers until it
 * expires).
 */
export async function getTeamsWebToken(options: GetTeamsWebTokenOptions = {}): Promise<TeamsWebAccessToken> {
  const now = options.now ?? options.deps?.now?.() ?? Date.now();
  const read = await readTeamsWebCache();
  const cached = read.cache;
  if (cached && isTeamsWebTokenFresh(cached.expires_at, now) && !options.force) {
    return toAccessToken(cached);
  }

  const reason = cached ? `teams-web token expired at ${cached.expires_at}` : read.reason;

  if (cached?.refresh_token) {
    const refresh = options.deps?.refresh ?? refreshTeamsWeb;
    try {
      const updated = await refresh({ cache: cached, force: options.force === true });
      if (isTeamsWebTokenFresh(updated.expires_at, now)) return toAccessToken(updated);
      throw new AuthRequiredError(
        `refreshed Teams web token expires at ${updated.expires_at}, inside the safety margin`,
        reason,
      );
    } catch (error) {
      if (error instanceof AuthRequiredError) throw error;
      if (error instanceof RefreshError && !error.requiresLogin) {
        throw new AuthRequiredError(
          `could not refresh the Teams web token: ${error.message}`,
          error.hint ?? "transient failure; retry shortly, no browser harvest needed",
        );
      }
      // invalid_grant: the refresh token is dead, fall through to a harvest.
    }
  }

  if (options.allowBrowser === false) {
    throw new AuthRequiredError(
      `no usable Teams web token: ${reason}`,
      "run teams op=status for the cache path; a browser harvest is required and was disabled for this call",
    );
  }

  const harvest = options.deps?.harvest ?? harvestTeamsWeb;
  try {
    const harvested = await harvest({ log: options.log });
    if (!isTeamsWebTokenFresh(harvested.expires_at, now)) {
      throw new AuthRequiredError("harvested Teams web token is already inside the expiry margin", reason);
    }
    return toAccessToken(harvested);
  } catch (error) {
    if (error instanceof AuthRequiredError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    throw new AuthRequiredError(`Teams web harvest failed: ${message}`, reason);
  }
}

export interface TeamsWebStatus {
  authenticated: boolean;
  fresh: boolean;
  path: string;
  expires_at?: string;
  seconds_left: number | null;
  source?: TeamsWebCache["source"];
  refreshable: boolean;
  client_id: string;
  account?: { upn?: string; id?: string };
  scope_count: number;
  reason?: string;
  corrupt?: true;
}

/** Offline status for teams op=status (no network). */
export async function teamsWebStatus(now = Date.now()): Promise<TeamsWebStatus> {
  const read = await readTeamsWebCache();
  const cache = read.cache;
  if (!cache) {
    return {
      authenticated: false,
      fresh: false,
      path: read.path,
      seconds_left: null,
      refreshable: false,
      client_id: TEAMS_WEB_CLIENT_ID,
      scope_count: 0,
      reason: read.reason,
      ...(read.corrupt ? { corrupt: true as const } : {}),
    };
  }
  const expiry = Date.parse(cache.expires_at);
  return {
    authenticated: true,
    fresh: isTeamsWebTokenFresh(cache.expires_at, now),
    path: read.path,
    expires_at: cache.expires_at,
    seconds_left: Number.isNaN(expiry) ? null : Math.round((expiry - now) / 1000),
    source: cache.source,
    refreshable: Boolean(cache.refresh_token),
    client_id: cache.client_id,
    ...(cache.account ? { account: cache.account } : {}),
    scope_count: cache.scope.split(" ").filter(Boolean).length,
  };
}

/** Remove the Teams-web credential (used by auth.logout). */
export async function clearTeamsWebCache(): Promise<{ removed: boolean; path: string }> {
  const path = config().teamsWebTokenPath;
  try {
    await unlink(path);
    return { removed: true, path };
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return { removed: false, path };
    throw error;
  }
}
