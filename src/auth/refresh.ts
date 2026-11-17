/**
 * Refresh-token grant against the Microsoft identity platform.
 *
 * Everything here is shaped by observed tenant behaviour (docs/auth.md):
 *
 *  - The only usable client is Outlook Web: `9199bf20-a13f-4107-85dc-02114787ef48`.
 *  - It is an SPA client type, so the token endpoint REJECTS a request without an
 *    `Origin` header (`AADSTS9002327`). We always send `https://outlook.office.com`.
 *  - There is no client secret (public client).
 *  - Only `.default` scopes are honoured; explicit scopes fail with `AADSTS65002`,
 *    so each call passes exactly one resource `.default` scope.
 *  - Every successful redemption returns a NEW refresh token. The caller must
 *    persist it before using the access token.
 */
import type { Audience } from "./store.js";

export const AUTH_HOST = "https://login.microsoftonline.com";
export const OWA_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48";
export const OWA_ORIGIN = "https://outlook.office.com";
export const OWA_AUTHORITY_ORIGIN = "https://outlook.cloud.microsoft";

/** Resource `.default` scope per audience. */
export const AUDIENCE_SCOPE: Readonly<Record<Audience, string>> = {
  outlook: "https://outlook.office.com/.default",
  graph: "https://graph.microsoft.com/.default",
};

export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

export interface RefreshRequest {
  url: string;
  method: "POST";
  headers: Record<string, string>;
  body: string;
}

export interface BuildRefreshRequestOptions {
  tenantId: string;
  refreshToken: string;
  audience: Audience;
  clientId?: string;
  /** Defaults to `AUDIENCE_SCOPE[audience]`; `null` omits `scope` entirely. */
  scope?: string | null;
  origin?: string;
}

/**
 * Build the exact form POST. Exported so tests can assert the wire format
 * (grant type, client id, scope, and the mandatory Origin header).
 */
export function buildRefreshRequest(options: BuildRefreshRequestOptions): RefreshRequest {
  const {
    tenantId,
    refreshToken,
    audience,
    clientId = OWA_CLIENT_ID,
    scope = AUDIENCE_SCOPE[audience],
    origin = OWA_ORIGIN,
  } = options;
  const form = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: clientId,
    refresh_token: refreshToken,
  });
  if (scope !== null) form.set("scope", scope);
  return {
    url: `${AUTH_HOST}/${tenantId}/oauth2/v2.0/token`,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      // SPA client type: present or AADSTS9002327.
      Origin: origin,
      Accept: "application/json",
    },
    body: form.toString(),
  };
}

export interface RefreshClaims {
  aud?: string;
  tid?: string;
  oid?: string;
  upn?: string;
  preferred_username?: string;
  scp?: string;
  exp?: number;
}

export interface RefreshResult {
  accessToken: string;
  /** Present when the provider rotated it — always persist this. */
  refreshToken?: string;
  /** Seconds the access token is valid for, as reported by the endpoint. */
  expiresIn: number;
  expiresAt: string;
  scope: string;
  claims: RefreshClaims | undefined;
}

/** Decode a JWT payload without verifying it (used only for telemetry/claims). */
export function decodeJwtClaims(jwt: string): RefreshClaims | undefined {
  const payload = jwt.split(".")[1];
  if (!payload) return undefined;
  try {
    return JSON.parse(Buffer.from(payload.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8")) as RefreshClaims;
  } catch {
    return undefined;
  }
}

export interface TokenErrorBody {
  error?: string;
  error_description?: string;
  error_codes?: number[];
  trace_id?: string;
  correlation_id?: string;
}

/** Typed refresh failure, with the AADSTS code pulled out for the caller. */
export class RefreshError extends Error {
  readonly status: number;
  readonly code: string;
  readonly aadsts?: string;
  readonly traceId?: string;
  readonly hint?: string;

  constructor(init: { status: number; code: string; message: string; aadsts?: string; traceId?: string; hint?: string }) {
    super(init.message);
    this.name = "RefreshError";
    this.status = init.status;
    this.code = init.code;
    if (init.aadsts !== undefined) this.aadsts = init.aadsts;
    if (init.traceId !== undefined) this.traceId = init.traceId;
    if (init.hint !== undefined) this.hint = init.hint;
  }

  /**
   * True when the refresh token itself is dead and re-authentication is the only
   * way forward. These are the codes observed/likely for rotation races and
   * revoked sessions.
   */
  get requiresLogin(): boolean {
    if (this.code === "invalid_grant") return true;
    if (!this.aadsts) return false;
    return /^(AADSTS70008|AADSTS700082|AADSTS50173|AADSTS9002313|AADSTS70000|AADSTS50078|AADSTS50107)$/.test(this.aadsts);
  }
}

const AADSTS_RE = /AADSTS\d{5,7}/;

function hintFor(aadsts: string | undefined): string | undefined {
  switch (aadsts) {
    case "AADSTS9002327":
      return "the token endpoint requires the Origin header for this SPA client (see docs/auth.md)";
    case "AADSTS65002":
      return "explicit scopes are not preauthorised for this first-party client; use the .default scope";
    case "AADSTS50105":
      return "this tenant blocks non-assigned apps; only the Outlook Web client can be used";
    case "AADSTS70008":
    case "AADSTS700082":
      return "the refresh token expired or was already redeemed; run auth.login";
    case undefined:
      return undefined;
    default:
      return "run auth.login to re-establish the session";
  }
}

/** Turn a non-2xx token response into a typed error. */
export function parseTokenError(status: number, body: TokenErrorBody | string): RefreshError {
  const parsed: TokenErrorBody = typeof body === "string" ? { error_description: body } : body;
  const description = parsed.error_description ?? "";
  const aadsts = AADSTS_RE.exec(description)?.[0] ?? (parsed.error_codes ? `AADSTS${String(parsed.error_codes[0])}` : undefined);
  const code = parsed.error ?? (status === 400 ? "invalid_request" : `http_${status}`);
  return new RefreshError({
    status,
    code,
    message: description.split(" Trace ID:")[0] ?? code,
    ...(aadsts ? { aadsts } : {}),
    ...(parsed.trace_id ? { traceId: parsed.trace_id } : {}),
    ...(hintFor(aadsts) ? { hint: hintFor(aadsts) } : {}),
  });
}

interface TokenSuccessBody {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number | string;
  scope?: string;
  token_type?: string;
}

/** Validate and normalise a 200 token response. */
export function parseTokenSuccess(body: TokenSuccessBody, now = Date.now()): RefreshResult {
  if (typeof body.access_token !== "string" || body.access_token === "") {
    throw new RefreshError({ status: 200, code: "invalid_response", message: "token response has no access_token" });
  }
  const expiresIn = Number(body.expires_in);
  const ttl = Number.isFinite(expiresIn) && expiresIn > 0 ? expiresIn : 3600;
  const claims = decodeJwtClaims(body.access_token);
  const scope = (claims?.scp ? claims.scp.split(" ") : (body.scope ?? "").split(" ").filter(Boolean)).join(" ");
  return {
    accessToken: body.access_token,
    ...(typeof body.refresh_token === "string" && body.refresh_token !== "" ? { refreshToken: body.refresh_token } : {}),
    expiresIn: ttl,
    expiresAt: new Date(now + ttl * 1000).toISOString(),
    scope,
    claims,
  };
}

export interface RefreshAccessTokenOptions extends BuildRefreshRequestOptions {
  fetchImpl?: FetchLike;
  now?: number;
  timeoutMs?: number;
}

/**
 * Redeem a refresh token. Does NOT persist anything — the caller owns the
 * "persist rotation before use" ordering (see store.refreshAudience).
 */
export async function refreshAccessToken(options: RefreshAccessTokenOptions): Promise<RefreshResult> {
  const { fetchImpl = fetch, now = Date.now(), timeoutMs = 30_000, ...requestOptions } = options;
  const request = buildRefreshRequest(requestOptions);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  let response: Response;
  try {
    response = await fetchImpl(request.url, {
      method: request.method,
      headers: request.headers,
      body: request.body,
      signal: controller.signal,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new RefreshError({
      status: 0,
      code: controller.signal.aborted ? "timeout" : "network_error",
      message: `token endpoint unreachable: ${message}`,
      hint: "check connectivity; a transient failure does not require a new login",
    });
  } finally {
    clearTimeout(timer);
  }

  const text = await response.text();
  let json: unknown;
  try {
    json = JSON.parse(text);
  } catch {
    json = undefined;
  }
  if (!response.ok) {
    throw parseTokenError(response.status, (json ?? text) as TokenErrorBody | string);
  }
  return parseTokenSuccess((json ?? {}) as TokenSuccessBody, now);
}
