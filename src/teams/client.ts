/**
 * Graph client for Teams.
 *
 * Two access tokens reach the same API, with different permissions:
 *
 *  - the OWA-derived `graph` token (`src/graph/client.ts`) — chats, team and
 *    channel metadata, channel create/send;
 *  - the Teams-web token (`src/auth/teams-web.ts`) — channel messages, replies,
 *    channel/team members.
 *
 * `TeamsWebClient` is the second one. A 403 from Graph is normalised here into
 * `missing_scope` naming the permission the token lacks, because "Forbidden" is
 * not actionable for an agent.
 */
import { config } from "../config.js";
import { getTeamsWebToken } from "../auth/teams-web.js";
import {
  ApiError,
  HttpClient,
  type FetchLike,
  type HttpClientOptions,
  type TokenProvider,
} from "../http/client.js";

export interface TeamsWebClientOptions {
  baseUrl?: string;
  tokenProvider?: TokenProvider;
  fetchImpl?: FetchLike;
  maxAttempts?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  sleep?: (ms: number) => Promise<void>;
  random?: () => number;
  timeoutMs?: number;
  userAgent?: string;
}

export class TeamsWebClient extends HttpClient {
  constructor(options: TeamsWebClientOptions = {}) {
    const cfg = config();
    const base: HttpClientOptions = {
      baseUrl: options.baseUrl ?? cfg.graphBaseUrl,
      tokenProvider:
        options.tokenProvider ?? (async () => ({ token: (await getTeamsWebToken()).token })),
      maxAttempts: options.maxAttempts ?? cfg.maxAttempts,
      timeoutMs: options.timeoutMs ?? cfg.requestTimeoutMs,
    };
    if (options.fetchImpl) base.fetchImpl = options.fetchImpl;
    if (options.baseDelayMs !== undefined) base.baseDelayMs = options.baseDelayMs;
    if (options.maxDelayMs !== undefined) base.maxDelayMs = options.maxDelayMs;
    if (options.sleep) base.sleep = options.sleep;
    if (options.random) base.random = options.random;
    if (options.userAgent) base.userAgent = options.userAgent;
    super(base);
  }
}

let shared: TeamsWebClient | undefined;

export function teamsWebClient(): TeamsWebClient {
  shared ??= new TeamsWebClient();
  return shared;
}

/** Test seam: replace the shared client (pass undefined to reset). */
export function setTeamsWebClient(client: TeamsWebClient | undefined): void {
  shared = client;
}

/**
 * Permission a Graph 403 is asking for. Both shapes observed in this tenant:
 *
 *   Missing scope permissions on the request. API requires one of 'X'. Scopes…
 *   Caller does not have the required permissions … AllowedPermissions:'X,Y'
 */
export function requiredScopeOf(message: string): string | undefined {
  const single = /API requires one of '([^']+)'/i.exec(message)?.[1];
  if (single) return single.split(",")[0]?.trim();
  const allowed = /AllowedPermissions:'([^']+)'/i.exec(message)?.[1];
  if (allowed) return allowed.split(",")[0]?.trim();
  return undefined;
}

/**
 * Map a provider failure onto the tool-facing error. Only 403 changes shape;
 * everything else is rethrown untouched.
 */
export function asTeamsError(error: unknown): unknown {
  if (error instanceof ApiError && (error.status === 404 || error.status === 410)) {
    return new ApiError({
      status: error.status,
      code: "not_found",
      message: "Graph did not find this Teams resource (it may have been deleted)",
      hint: "list the team/chats again; ids come from a prior list/read and deleted channels answer Gone",
    });
  }
  if (!(error instanceof ApiError) || error.status !== 403) return error;
  const scope = requiredScopeOf(error.message);
  if (scope) {
    return new ApiError({
      status: 403,
      code: "missing_scope",
      message: `this token lacks ${scope} (Graph refused the call)`,
      hint: `grant the delegated permission ${scope}, or use the token that has it (channel history needs the Teams-web token)`,
    });
  }
  // Private channels answer a bare 403 ("UnknownError") when the account is not
  // a member; that is a boundary, not a missing delegated permission.
  return new ApiError({
    status: 403,
    code: "forbidden",
    message: "Graph refused this Teams resource",
    hint: "the credential is valid but this account cannot see it — typically a private channel you are not a member of; the channel list only proves metadata access",
  });
}
