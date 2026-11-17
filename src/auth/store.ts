/**
 * Token cache.
 *
 * Canonical file: `${TEAMS_MCP_CONFIG_DIR:-~/.config/teams-mcp}/token.json`, 0600,
 * written atomically. Shape (version 1):
 *
 * {
 *   "version": 1,
 *   "account": {"upn": "user@example.com", "id": "<object id>"},
 *   "refresh_token": "<rotating refresh token>",
 *   "tokens": {
 *     "outlook": {"access_token": "...", "expires_at": "2026-11-16T20:31:00.000Z", "scope": "..."},
 *     "graph":   {"access_token": "...", "expires_at": "2026-11-16T20:31:00.000Z", "scope": "..."}
 *   },
 *   "source": "refresh" | "browser" | "headless",
 *   "updated_at": "2026-11-16T19:31:00.000Z"
 * }
 *
 * Two audiences because the tenant denies Graph app assignment for mail/calendar:
 * Outlook REST calls use the `outlook.office.com` token, directory calls use the
 * Graph token. A token is stale 5 minutes before `expires_at`.
 *
 * Read-only legacy support (never rewritten):
 *  - `tokens.json` (OWA/MSAL capture: accessToken/expiresAt(epoch ms)/claims), and
 *  - the earlier single-audience v1 document with a top-level `access_token`.
 * Both are normalised into `tokens.outlook` (or `tokens.graph` when the captured
 * audience says Graph).
 *
 * The login/refresh flow belongs to the auth workstream. It must persist a
 * rotated refresh token with `writeTokenCache`/`persistTokenUpdate` BEFORE
 * returning a new access token.
 */
import { chmod, mkdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { dirname, join } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { config, TOKEN_EXPIRY_SKEW_MS } from "../config.js";
// Call-time-only cycle: login.ts imports this module for persistence, and this
// module imports login.ts for the fallback. Neither uses the other at module
// evaluation time, so ESM resolves the cycle safely.
import { performLogin } from "./login.js";
import { acquireFileLock, LOCK_WAIT_MS, LockTimeoutError } from "./lock.js";
import { decodeJwtClaims, refreshAccessToken, RefreshError } from "./refresh.js";

export const TOKEN_CACHE_VERSION = 1 as const;

export type Audience = "outlook" | "graph";
export type TokenSource = "headless" | "browser" | "refresh";
export type TokenFormat = "v1" | "legacy";

export interface AudienceToken {
  access_token: string;
  /** ISO 8601 UTC. */
  expires_at: string;
  scope?: string;
}

export interface TokenAccount {
  upn?: string;
  id?: string;
}

export interface TokenCache {
  version: typeof TOKEN_CACHE_VERSION;
  account?: TokenAccount;
  refresh_token?: string;
  tokens: Partial<Record<Audience, AudienceToken>>;
  source: TokenSource;
  updated_at: string;
}

export interface AccessToken {
  token: string;
  expiresAt: string;
  source: TokenSource;
  audience: Audience;
}

export type CacheReadResult =
  | { cache: TokenCache; path: string; format: TokenFormat }
  | {
      cache: null;
      path: string;
      reason: string;
      /** A candidate existed but was unreadable/invalid — not simply absent. */
      corrupt?: true;
    };

/** Raised when a credential is required but the flow that obtains one is not wired up. */
export class AuthNotImplementedError extends Error {
  readonly code = "auth_not_implemented";
  constructor(message: string, readonly hint?: string) {
    super(message);
    this.name = "AuthNotImplementedError";
  }
}

/** Raised when an operation needs a valid token and the cache has none. */
export class AuthRequiredError extends Error {
  readonly code = "auth_required";
  constructor(message: string, readonly hint?: string) {
    super(message);
    this.name = "AuthRequiredError";
  }
}

const SOURCES: Record<TokenSource, true> = { headless: true, browser: true, refresh: true };
const AUDIENCES: Record<Audience, true> = { outlook: true, graph: true };

function isTokenSource(value: unknown): value is TokenSource {
  return typeof value === "string" && Object.hasOwn(SOURCES, value);
}

function isAudience(value: unknown): value is Audience {
  return typeof value === "string" && Object.hasOwn(AUDIENCES, value);
}

function validateAudienceToken(raw: unknown): AudienceToken {
  if (raw === null || typeof raw !== "object") throw new Error("audience token is not an object");
  const token = raw as Record<string, unknown>;
  if (typeof token.access_token !== "string" || token.access_token === "") {
    throw new Error("audience token has no access_token");
  }
  if (typeof token.expires_at !== "string" || Number.isNaN(Date.parse(token.expires_at))) {
    throw new Error("audience token has no valid expires_at");
  }
  return {
    access_token: token.access_token,
    expires_at: new Date(token.expires_at).toISOString(),
    ...(typeof token.scope === "string" ? { scope: token.scope } : {}),
  };
}

function validateDocument(raw: unknown): TokenCache {
  if (raw === null || typeof raw !== "object") throw new Error("token cache is not an object");
  const doc = raw as Record<string, unknown>;
  if (doc.version !== TOKEN_CACHE_VERSION) {
    throw new Error(`unsupported token cache version ${String(doc.version)}`);
  }
  if (!isTokenSource(doc.source)) throw new Error("token cache has an invalid source");
  if (doc.tokens === null || typeof doc.tokens !== "object") {
    throw new Error("token cache has no tokens map");
  }
  const rawTokens = doc.tokens as Record<string, unknown>;
  const tokens: Partial<Record<Audience, AudienceToken>> = {};
  for (const [audience, value] of Object.entries(rawTokens)) {
    if (value === null || value === undefined) continue;
    if (!isAudience(audience)) throw new Error(`token cache has an unknown audience "${audience}"`);
    tokens[audience] = validateAudienceToken(value);
  }
  if (Object.keys(tokens).length === 0) throw new Error("token cache has no audience tokens");

  const account = doc.account;
  if (account !== undefined && (account === null || typeof account !== "object")) {
    throw new Error("token cache account must be an object");
  }
  const updatedAt =
    typeof doc.updated_at === "string" && !Number.isNaN(Date.parse(doc.updated_at))
      ? new Date(doc.updated_at).toISOString()
      : new Date(0).toISOString();

  return {
    version: TOKEN_CACHE_VERSION,
    ...(account ? { account: account as TokenAccount } : {}),
    ...(typeof doc.refresh_token === "string" ? { refresh_token: doc.refresh_token } : {}),
    tokens,
    source: doc.source,
    updated_at: updatedAt,
  };
}

/** Audience implied by a captured token's `aud` claim. */
function audienceFromClaims(claims: Record<string, unknown>): Audience {
  const aud = typeof claims.aud === "string" ? claims.aud : "";
  return /graph\.microsoft\.com/i.test(aud) ? "graph" : "outlook";
}

/** Map an OWA/MSAL capture (tokens.json) onto the canonical shape. */
export function normalizeLegacyDocument(raw: unknown): TokenCache | null {
  if (raw === null || typeof raw !== "object") return null;
  const doc = raw as Record<string, unknown>;

  // Earlier single-audience v1: top-level access_token/expires_at.
  if (typeof doc.access_token === "string" && typeof doc.expires_at === "string") {
    return {
      version: TOKEN_CACHE_VERSION,
      tokens: {
        outlook: {
          access_token: doc.access_token,
          expires_at: new Date(doc.expires_at).toISOString(),
          ...(typeof doc.scope === "string" ? { scope: doc.scope } : {}),
        },
      },
      source: isTokenSource(doc.source) ? doc.source : "headless",
      ...(typeof doc.refresh_token === "string" ? { refresh_token: doc.refresh_token } : {}),
      ...(doc.account && typeof doc.account === "object" ? { account: doc.account as TokenAccount } : {}),
      updated_at:
        typeof doc.updated_at === "string" && !Number.isNaN(Date.parse(doc.updated_at))
          ? new Date(doc.updated_at).toISOString()
          : new Date(0).toISOString(),
    };
  }

  // OWA/MSAL capture.
  const access = doc.accessToken ?? doc.access_token;
  if (typeof access !== "string" || access === "") return null;
  const expiryRaw = doc.expiresAt ?? doc.expires_at;
  const expiry =
    typeof expiryRaw === "number"
      ? new Date(expiryRaw)
      : typeof expiryRaw === "string"
        ? new Date(/^\d+$/.test(expiryRaw) ? Number(expiryRaw) : expiryRaw)
        : undefined;
  if (!expiry || Number.isNaN(expiry.getTime())) return null;

  const claims = (doc.claims ?? {}) as Record<string, unknown>;
  const obtainedFrom = typeof doc.obtainedFrom === "string" ? doc.obtainedFrom : "";
  const scopes = Array.isArray(doc.scopes)
    ? doc.scopes.filter((scope): scope is string => typeof scope === "string")
    : [];
  const upn =
    (typeof claims.upn === "string" ? claims.upn : undefined) ??
    (typeof doc.username === "string" ? doc.username : undefined);
  const oid = typeof claims.oid === "string" ? claims.oid : undefined;

  return {
    version: TOKEN_CACHE_VERSION,
    tokens: {
      [audienceFromClaims(claims)]: {
        access_token: access,
        expires_at: expiry.toISOString(),
        ...(scopes.length > 0 ? { scope: scopes.join(" ") } : {}),
      },
    },
    source: /refresh/i.test(obtainedFrom) ? "refresh" : "headless",
    ...(typeof doc.refreshToken === "string" ? { refresh_token: doc.refreshToken } : {}),
    ...(upn || oid ? { account: { ...(upn ? { upn } : {}), ...(oid ? { id: oid } : {}) } } : {}),
    updated_at:
      typeof doc.obtainedAt === "string" && !Number.isNaN(Date.parse(doc.obtainedAt))
        ? new Date(doc.obtainedAt).toISOString()
        : new Date(0).toISOString(),
  };
}

async function readTokenFile(path: string): Promise<CacheReadResult> {
  let text: string;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    return {
      cache: null,
      path,
      reason: code === "ENOENT" ? "missing" : `token file unreadable (${code ?? "error"})`,
    };
  }
  try {
    const raw = JSON.parse(text) as Record<string, unknown>;
    if (raw.version === TOKEN_CACHE_VERSION && raw.tokens !== undefined) {
      return { cache: validateDocument(raw), path, format: "v1" };
    }
    const normalized = normalizeLegacyDocument(raw);
    if (normalized) return { cache: normalized, path, format: "legacy" };
    throw new Error("unrecognised token document (neither v1 nor a legacy capture)");
  } catch (error) {
    return { cache: null, path, reason: `invalid: ${(error as Error).message}` };
  }
}

/**
 * Read the token cache. Order: canonical `token.json`, then the read-only legacy
 * captures. Missing/corrupt documents yield a reason, never a throw.
 */
export async function readTokenCache(path?: string): Promise<CacheReadResult> {
  const cfg = config();
  const candidates = path ? [path] : [cfg.tokenPath, ...cfg.legacyTokenPaths];
  const reasons: string[] = [];
  let corrupt = false;
  for (const candidate of candidates) {
    const result = await readTokenFile(candidate);
    if (result.cache !== null) return result;
    // `readTokenFile` reports a bare "missing"; every other reason means the file
    // was there and could not be used.
    if (result.reason !== "missing") corrupt = true;
    reasons.push(`${result.path}: ${result.reason}`);
  }
  const allMissing = reasons.every((reason) => reason.endsWith(": missing"));
  return {
    cache: null,
    path: candidates[0] ?? cfg.tokenPath,
    reason: allMissing ? "no token cache (not logged in)" : reasons.join("; "),
    ...(corrupt ? { corrupt: true as const } : {}),
  };
}

/** Atomically persist a token document with owner-only permissions. */
export async function writeTokenCache(cache: TokenCache, path = config().tokenPath): Promise<void> {
  const doc = validateDocument({ ...cache, version: TOKEN_CACHE_VERSION });
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const tmp = `${path}.${process.pid}.${randomUUID()}.tmp`;
  await writeFile(tmp, `${JSON.stringify(doc, null, 2)}\n`, { mode: 0o600 });
  await rename(tmp, path);
  await chmod(path, 0o600).catch(() => undefined);
}

export interface TokenUpdate {
  audience: Audience;
  access_token: string;
  expires_at: string;
  scope?: string;
  /** New refresh token when the provider rotated it. Persisted before returning. */
  refresh_token?: string;
  source?: TokenSource;
  account?: TokenAccount;
  now?: number;
}

/**
 * Merge one audience's token into the existing document (preserving the other
 * audience and a rotated refresh token) and persist it atomically.
 */
export async function persistTokenUpdate(update: TokenUpdate, path = config().tokenPath): Promise<TokenCache> {
  const existing = await readTokenCache(path);
  const current = existing.cache;
  const next: TokenCache = {
    version: TOKEN_CACHE_VERSION,
    tokens: {
      ...(current?.tokens ?? {}),
      [update.audience]: {
        access_token: update.access_token,
        expires_at: new Date(update.expires_at).toISOString(),
        ...(update.scope !== undefined
          ? { scope: update.scope }
          : current?.tokens[update.audience]?.scope !== undefined
            ? { scope: current.tokens[update.audience]?.scope as string }
            : {}),
      },
    },
    source: update.source ?? current?.source ?? "refresh",
    updated_at: new Date(update.now ?? Date.now()).toISOString(),
    ...(update.refresh_token ?? current?.refresh_token
      ? { refresh_token: (update.refresh_token ?? current?.refresh_token) as string }
      : {}),
    ...(update.account ?? current?.account ? { account: (update.account ?? current?.account) as TokenAccount } : {}),
  };
  await writeTokenCache(next, path);
  return next;
}

/** True when the token is still usable after applying the 5 minute expiry skew. */
export function isTokenFresh(expiresAt: string, now = Date.now()): boolean {
  const expiry = Date.parse(expiresAt);
  if (Number.isNaN(expiry)) return false;
  return expiry - TOKEN_EXPIRY_SKEW_MS > now;
}

export interface GetAccessTokenOptions {
  audience?: Audience;
  now?: () => number;
  /** Test seams; production callers omit this. */
  deps?: AuthHooks;
}

/** Injectable collaborators so orchestration is testable without a browser or network. */
export interface AuthHooks {
  /** Redeem the cache's refresh token and persist the result. */
  refresh?: (request: { audience: Audience; cache: TokenCache }) => Promise<TokenCache>;
  /** Run the sign-in flow and return the refreshed cache. */
  login?: (request: LoginRequest) => Promise<TokenCache>;
  now?: () => number;
}

function toAccessToken(entry: AudienceToken, audience: Audience, source: TokenSource): AccessToken {
  return { token: entry.access_token, expiresAt: entry.expires_at, source, audience };
}

/** Tenant id from any cached token's claims; refresh needs it in the URL path. */
function tenantIdFrom(cache: TokenCache | null): string | undefined {
  for (const entry of Object.values(cache?.tokens ?? {})) {
    const tid = entry ? decodeJwtClaims(entry.access_token)?.tid : undefined;
    if (tid) return tid;
  }
  return undefined;
}

/** Advisory cross-process lock guarding refresh-token rotation. */
export function tokenLockPath(): string {
  return join(config().configDir, "token.lock");
}

/**
 * Serialise refreshes across processes. Refresh tokens are single-use, so two
 * concurrent refreshes with the same token would invalidate each other's
 * session.
 *
 * The lock file holds `{pid, at}`. A lock older than two minutes is presumed
 * abandoned and taken over; if a live lock blocks us for two minutes we raise an
 * actionable error rather than spinning forever. The browser login never holds
 * this lock — a 60s interactive flow must not block other instances.
 *
 * `waitMs` is exposed for tests; production callers use the 120 s default.
 */
export async function acquireTokenLock(waitMs = LOCK_WAIT_MS): Promise<() => Promise<void>> {
  try {
    return await acquireFileLock(tokenLockPath(), { waitMs, label: "token lock" });
  } catch (error) {
    if (error instanceof LockTimeoutError) {
      throw new AuthRequiredError(
        error.message,
        "another teams-mcp process is refreshing the session; retry shortly, or delete the lock file if nothing else is running",
      );
    }
    throw error;
  }
}

/** Network hiccups and 5xx are worth one retry; a dead grant is not. */
function isTransientFailure(error: RefreshError): boolean {
  if (error.code === "network_error" || error.code === "timeout") return true;
  return error.status >= 500;
}

async function redeemWithRetry(params: {
  tenantId: string;
  refreshToken: string;
  audience: Audience;
}): Promise<Awaited<ReturnType<typeof refreshAccessToken>>> {
  try {
    return await refreshAccessToken(params);
  } catch (error) {
    if (error instanceof RefreshError && !error.requiresLogin && isTransientFailure(error)) {
      await delay(500);
      return await refreshAccessToken(params);
    }
    throw error;
  }
}

/**
 * Default refresh path: take the lock, re-read the cache (another process may
 * have refreshed while we waited), redeem the NEWEST refresh token, then persist
 * the rotation atomically before returning.
 */
async function refreshFromCache(request: { audience: Audience; cache: TokenCache }): Promise<TokenCache> {
  const path = config().tokenPath;
  const release = await acquireTokenLock();
  try {
    const latest = (await readTokenCache(path)).cache ?? request.cache;
    const alreadyFresh = latest.tokens[request.audience];
    if (alreadyFresh && isTokenFresh(alreadyFresh.expires_at)) return latest;

    const refreshToken = latest.refresh_token;
    if (!refreshToken) {
      throw new RefreshError({ status: 0, code: "no_refresh_token", message: "token cache has no refresh token" });
    }
    const result = await redeemWithRetry({
      tenantId: tenantIdFrom(latest) ?? "common",
      refreshToken,
      audience: request.audience,
    });
    // Rotated refresh token + access token land on disk in one atomic write.
    return await persistTokenUpdate(
      {
        audience: request.audience,
        access_token: result.accessToken,
        expires_at: result.expiresAt,
        ...(result.scope ? { scope: result.scope } : {}),
        refresh_token: result.refreshToken ?? refreshToken,
        source: "refresh",
      },
      path,
    );
  } finally {
    await release();
  }
}

/** A concurrent refresh may have produced a usable token; adopt it before failing. */
async function recoverFromDisk(audience: Audience, now: number): Promise<AccessToken | undefined> {
  const result = await readTokenCache();
  const entry = result.cache?.tokens[audience];
  if (result.cache && entry && isTokenFresh(entry.expires_at, now)) {
    return toAccessToken(entry, audience, result.cache.source);
  }
  return undefined;
}

/** Default login path: run the browser flow, then read back what it persisted. */
async function loginAndRead(request: LoginRequest): Promise<TokenCache> {
  if (request.reason) process.stderr.write(`[auth] signing in (${request.reason})\n`);
  await performLogin();
  const result = await readTokenCache();
  if (!result.cache) throw new AuthRequiredError("login finished but no token cache was written", result.reason);
  return result.cache;
}

function asAuthRequired(error: unknown, audience: Audience, reason: string): AuthRequiredError {
  if (error instanceof AuthRequiredError) return error;
  const message = error instanceof Error ? error.message : String(error);
  const hint = error instanceof Error && "hint" in error && typeof error.hint === "string" ? error.hint : reason;
  return new AuthRequiredError(`could not obtain a ${audience} token: ${message}`, hint);
}

/**
 * Return a token for the requested audience.
 *
 * Order: fresh cache -> refresh (rotation persisted first) -> one sign-in. A
 * transient refresh failure does not trigger a browser launch; a dead refresh
 * token does. Login failures surface as `AuthRequiredError` with an actionable
 * hint, never a raw stack trace.
 */
export async function getAccessToken(options: GetAccessTokenOptions = {}): Promise<AccessToken> {
  const audience = options.audience ?? "outlook";
  const now = options.now?.() ?? options.deps?.now?.() ?? Date.now();
  const result = await readTokenCache();
  const cached = result.cache?.tokens[audience];
  if (result.cache && cached && isTokenFresh(cached.expires_at, now)) {
    return toAccessToken(cached, audience, result.cache.source);
  }

  const reason =
    result.cache === null
      ? result.reason
      : cached
        ? `${audience} token expired at ${cached.expires_at}`
        : `no cached ${audience} token (have: ${Object.keys(result.cache.tokens).join(", ") || "none"})`;

  // A corrupt cache is not "not logged in": launching the interactive flow would
  // pop a browser at an arbitrary request and still leave the bad file behind.
  if (result.cache === null && result.corrupt) {
    throw new AuthRequiredError(
      `token cache is unreadable: ${result.reason}`,
      `delete ${result.path} and run auth login (the browser profile is kept)`,
    );
  }

  if (result.cache?.refresh_token) {
    const refresh = options.deps?.refresh ?? refreshFromCache;
    try {
      const updated = await refresh({ audience, cache: result.cache });
      const entry = updated.tokens[audience];
      // Never hand back a token that dies inside the 5-minute safety margin.
      if (entry && isTokenFresh(entry.expires_at, now)) return toAccessToken(entry, audience, updated.source);
      throw new AuthRequiredError(
        entry
          ? `refreshed ${audience} token expires at ${entry.expires_at}, inside the safety margin`
          : `refresh produced no ${audience} token`,
        reason,
      );
    } catch (error) {
      const recovered = await recoverFromDisk(audience, now);
      if (recovered) return recovered;
      if (error instanceof RefreshError && !error.requiresLogin) {
        throw new AuthRequiredError(
          `could not refresh ${audience} token: ${error.message}`,
          error.hint ?? "transient failure; retry shortly, no new login needed",
        );
      }
      if (error instanceof AuthRequiredError) throw error;
      // invalid_grant and friends: fall through to a single sign-in attempt.
    }
  }

  const login = options.deps?.login ?? loginAndRead;
  try {
    const updated = await login({ audience, reason });
    const entry = updated.tokens[audience];
    if (!entry || !isTokenFresh(entry.expires_at, now)) {
      throw new AuthRequiredError(`login returned no usable ${audience} token`, reason);
    }
    return toAccessToken(entry, audience, updated.source);
  } catch (error) {
    throw asAuthRequired(error, audience, reason);
  }
}

export interface LoginRequest {
  audience?: Audience;
  reason?: string;
  force?: boolean;
}

/**
 * Run the sign-in flow and return the requested audience's token. Always signs
 * in (callers use this for `auth.login`, not for automatic recovery).
 */
export async function forceLogin(request: LoginRequest = {}): Promise<AccessToken> {
  const audience = request.audience ?? "outlook";
  const cache = await loginAndRead(request);
  const entry = cache.tokens[audience];
  if (!entry) {
    throw new AuthRequiredError(`login produced no ${audience} token`, request.reason);
  }
  return toAccessToken(entry, audience, cache.source);
}

/** Remove the canonical cached credential. Browser profile and legacy captures stay. */
export async function logout(): Promise<{ removed: boolean; path: string }> {
  const path = config().tokenPath;
  try {
    await unlink(path);
    return { removed: true, path };
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return { removed: false, path };
    throw error;
  }
}
