/**
 * Graph client — DIRECTORY ONLY.
 *
 * In this tenant Graph app assignment is blocked for mail/calendar (403), so all
 * mail, calendar and scheduling calls go through Outlook REST v2.0
 * (`src/outlook/client.ts`). Graph remains the directory source:
 *
 *   GET /users?$search="displayName:..."   (ConsistencyLevel: eventual)
 *
 * It uses the `graph` token audience.
 */
import { config } from "../config.js";
import { getAccessToken } from "../auth/store.js";
import { HttpClient, type FetchLike, type HttpClientOptions, type TokenProvider } from "../http/client.js";

export { ApiError as GraphError } from "../http/client.js";
export type { FetchLike, TokenProvider } from "../http/client.js";

export interface GraphClientOptions {
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

export class GraphClient extends HttpClient {
  constructor(options: GraphClientOptions = {}) {
    const cfg = config();
    const base: HttpClientOptions = {
      baseUrl: options.baseUrl ?? cfg.graphBaseUrl,
      tokenProvider:
        options.tokenProvider ?? (async () => ({ token: (await getAccessToken({ audience: "graph" })).token })),
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

export interface DirectoryUser {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
}

const USER_SELECT = "id,displayName,mail,userPrincipalName,jobTitle,department";

/** `GET /users?$search="..."` — the only Graph call this server needs. */
export function searchDirectory(
  client: GraphClient,
  query: string,
  top: number,
): Promise<{ value?: DirectoryUser[] }> {
  const escaped = query.replace(/"/g, "");
  const search = `"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`;
  const params = new URLSearchParams({
    $search: search,
    $select: USER_SELECT,
    $top: String(top),
  });
  return client.get<{ value?: DirectoryUser[] }>(`users?${params.toString()}`, {
    headers: { ConsistencyLevel: "eventual" },
  });
}

/** `GET /users/{id or UPN}` for exact resolution. */
export function getUser(client: GraphClient, idOrUpn: string): Promise<DirectoryUser> {
  const params = new URLSearchParams({ $select: USER_SELECT });
  return client.get<DirectoryUser>(`users/${encodeURIComponent(idOrUpn)}?${params.toString()}`);
}

let shared: GraphClient | undefined;

export function graphClient(): GraphClient {
  shared ??= new GraphClient();
  return shared;
}

/** Test seam: replace the shared client (pass undefined to reset). */
export function setGraphClient(client: GraphClient | undefined): void {
  shared = client;
}
