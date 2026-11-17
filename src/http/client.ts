/**
 * Shared HTTP transport for the two APIs this server talks to:
 *
 *  - Outlook REST v2.0 (`https://outlook.office.com/api/v2.0`) — mail, calendar,
 *    scheduling, and the token audience `outlook`;
 *  - Microsoft Graph (`https://graph.microsoft.com/v1.0`) — directory only
 *    (`/users`), audience `graph`.
 *
 * Both tenants share retry/backoff, timeout, error normalisation and the
 * "never leak a token" rule; only base URL, audience and payload casing differ.
 */
import { isoTime } from "../format.js";

export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

export type TokenProvider = () => Promise<{ token: string }>;

export interface ApiErrorInit {
  status?: number;
  code: string;
  message: string;
  retryable?: boolean;
  retryAfterMs?: number;
  requestId?: string;
  hint?: string;
}

/** Normalised API failure. `code` is the provider's own code when present. */
export class ApiError extends Error {
  readonly status?: number;
  readonly code: string;
  readonly retryable: boolean;
  readonly retryAfterMs?: number;
  readonly requestId?: string;
  readonly hint?: string;

  constructor(init: ApiErrorInit) {
    super(init.message);
    this.name = "ApiError";
    if (init.status !== undefined) this.status = init.status;
    this.code = init.code;
    this.retryable = init.retryable ?? false;
    if (init.retryAfterMs !== undefined) this.retryAfterMs = init.retryAfterMs;
    if (init.requestId !== undefined) this.requestId = init.requestId;
    if (init.hint !== undefined) this.hint = init.hint;
  }

  /** One-line form for tool output. */
  toLine(): string {
    const where = this.status ? ` (HTTP ${this.status})` : "";
    return `${this.code}${where}: ${this.message}${this.hint ? ` — ${this.hint}` : ""}`;
  }
}

export interface HttpClientOptions {
  baseUrl: string;
  /** Supplies the bearer token for this API's audience. */
  tokenProvider: TokenProvider;
  /** Injectable transport, primarily for tests. */
  fetchImpl?: FetchLike;
  /** Total attempts per request, including the first. */
  maxAttempts?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  /** Injectable sleep so tests observe backoff without waiting. */
  sleep?: (ms: number) => Promise<void>;
  /** Injectable jitter source in [0,1). */
  random?: () => number;
  timeoutMs?: number;
  userAgent?: string;
}

export interface RequestOptions {
  method?: "GET" | "POST" | "PATCH" | "PUT" | "DELETE";
  body?: unknown;
  headers?: Record<string, string>;
  /** Overrides `maxAttempts` for this call. */
  maxAttempts?: number;
  signal?: AbortSignal;
}

const RETRYABLE_STATUS: Record<number, true> = {
  408: true,
  429: true,
  500: true,
  502: true,
  503: true,
  504: true,
};

function sleepDefault(ms: number): Promise<void> {
  const { promise, resolve } = Promise.withResolvers<void>();
  setTimeout(resolve, ms);
  return promise;
}

/** Parse Retry-After as delta seconds or an HTTP date. */
export function parseRetryAfter(value: string | null, now = Date.now()): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (/^\d+$/.test(trimmed)) return Number(trimmed) * 1000;
  const date = Date.parse(trimmed);
  if (Number.isNaN(date)) return undefined;
  return Math.max(0, date - now);
}

/** Bounded exponential backoff with small jitter; Retry-After always wins. */
export function computeDelay(
  attempt: number,
  options: { baseDelayMs: number; maxDelayMs: number; retryAfterMs?: number; random?: () => number },
): number {
  if (options.retryAfterMs !== undefined) return Math.min(options.retryAfterMs, options.maxDelayMs);
  const exponential = Math.min(options.maxDelayMs, options.baseDelayMs * 2 ** attempt);
  const jitter = exponential * 0.25 * ((options.random ?? Math.random)() * 2 - 1);
  return Math.max(0, Math.round(Math.min(options.maxDelayMs, exponential + jitter)));
}

/** Strip anything bearer-shaped before an error text escapes. */
export function sanitizeMessage(message: string): string {
  return message.replace(/\bBearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [redacted]").slice(0, 800);
}

interface ErrorEnvelope {
  error?: {
    code?: string | number;
    message?: string;
    innerError?: { "request-id"?: string; date?: string };
  };
}

export class HttpClient {
  readonly baseUrl: string;
  private readonly tokenProvider: TokenProvider;
  private readonly fetchImpl: FetchLike;
  private readonly maxAttempts: number;
  private readonly baseDelayMs: number;
  private readonly maxDelayMs: number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly random: () => number;
  private readonly timeoutMs: number;
  private readonly userAgent: string;

  constructor(options: HttpClientOptions) {
    this.baseUrl = options.baseUrl.replace(/\/+$/, "");
    this.tokenProvider = options.tokenProvider;
    this.fetchImpl = options.fetchImpl ?? ((input, init) => globalThis.fetch(input, init));
    this.maxAttempts = Math.max(1, options.maxAttempts ?? 4);
    this.baseDelayMs = options.baseDelayMs ?? 500;
    this.maxDelayMs = options.maxDelayMs ?? 8_000;
    this.sleep = options.sleep ?? sleepDefault;
    this.random = options.random ?? Math.random;
    this.timeoutMs = options.timeoutMs ?? 30_000;
    this.userAgent = options.userAgent ?? "teams-mcp/0.1.0";
  }

  get<T>(path: string, options: Omit<RequestOptions, "method" | "body"> = {}): Promise<T> {
    return this.request<T>(path, { ...options, method: "GET" });
  }

  post<T>(path: string, body?: unknown, options: Omit<RequestOptions, "method" | "body"> = {}): Promise<T> {
    return this.request<T>(path, { ...options, method: "POST", body });
  }

  patch<T>(path: string, body?: unknown, options: Omit<RequestOptions, "method" | "body"> = {}): Promise<T> {
    return this.request<T>(path, { ...options, method: "PATCH", body });
  }

  delete<T = void>(path: string, options: Omit<RequestOptions, "method" | "body"> = {}): Promise<T> {
    return this.request<T>(path, { ...options, method: "DELETE" });
  }

  /** Absolute URL for a relative path, or the URL itself. */
  url(pathOrUrl: string): string {
    if (/^https?:\/\//i.test(pathOrUrl)) return pathOrUrl;
    return `${this.baseUrl}/${pathOrUrl.replace(/^\/+/, "")}`;
  }

  /** True when a URL belongs to this client's API (guards `nextLink` reuse). */
  ownsUrl(candidate: string): boolean {
    return this.url(candidate).startsWith(`${this.baseUrl}/`);
  }

  async request<T>(pathOrUrl: string, options: RequestOptions = {}): Promise<T> {
    const url = this.url(pathOrUrl);
    const method = options.method ?? "GET";
    const maxAttempts = Math.max(1, options.maxAttempts ?? this.maxAttempts);
    const hasBody = options.body !== undefined;
    let lastError: ApiError | undefined;

    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
      const { token } = await this.tokenProvider();
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        "User-Agent": this.userAgent,
        ...(hasBody ? { "Content-Type": "application/json" } : {}),
        ...options.headers,
      };

      const timeout = AbortSignal.timeout(this.timeoutMs);
      const signal = options.signal ? AbortSignal.any([options.signal, timeout]) : timeout;

      let response: Response;
      try {
        response = await this.fetchImpl(url, {
          method,
          headers,
          body: hasBody ? JSON.stringify(options.body) : undefined,
          signal,
        });
      } catch (error) {
        const cause = error as Error;
        if (options.signal?.aborted) {
          throw new ApiError({ code: "aborted", message: "request aborted" });
        }
        const isTimeout = cause.name === "TimeoutError" || cause.name === "AbortError";
        const retryable = !isTimeout || attempt + 1 < maxAttempts;
        const apiError = new ApiError({
          code: isTimeout ? "timeout" : "network_error",
          message: isTimeout ? `request timed out after ${this.timeoutMs}ms` : sanitizeMessage(cause.message),
          retryable,
        });
        if (attempt + 1 >= maxAttempts || !retryable) throw apiError;
        lastError = apiError;
        await this.sleep(
          computeDelay(attempt, {
            baseDelayMs: this.baseDelayMs,
            maxDelayMs: this.maxDelayMs,
            random: this.random,
          }),
        );
        continue;
      }

      const requestId =
        response.headers.get("request-id") ?? response.headers.get("client-request-id") ?? undefined;

      if (response.ok) return (await decodeSuccess<T>(response, requestId)) as T;

      const retryAfterMs = parseRetryAfter(response.headers.get("retry-after"));
      const error = await decodeFailure(response, requestId, retryAfterMs);

      if (error.retryable && attempt + 1 < maxAttempts) {
        lastError = error;
        await this.sleep(
          computeDelay(attempt, {
            baseDelayMs: this.baseDelayMs,
            maxDelayMs: this.maxDelayMs,
            retryAfterMs,
            random: this.random,
          }),
        );
        continue;
      }
      throw error;
    }

    throw lastError ?? new ApiError({ code: "request_failed", message: "request failed after retries" });
  }
}

async function decodeSuccess<T>(response: Response, requestId?: string): Promise<T | undefined> {
  if (response.status === 204 || response.status === 205) return undefined;
  const text = await response.text();
  if (text.trim() === "") return undefined;
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new ApiError({
      code: "invalid_response",
      message: "API returned a body that is not valid JSON",
      status: response.status,
      ...(requestId ? { requestId } : {}),
    });
  }
}

async function decodeFailure(
  response: Response,
  requestId: string | undefined,
  retryAfterMs: number | undefined,
): Promise<ApiError> {
  const text = await response.text().catch(() => "");
  let body: ErrorEnvelope | undefined;
  if (text) {
    try {
      body = JSON.parse(text) as ErrorEnvelope;
    } catch {
      body = undefined;
    }
  }
  const status = response.status;
  const rawCode = body?.error?.code;
  const code =
    (typeof rawCode === "string" && rawCode) ||
    (typeof rawCode === "number" ? `error_${rawCode}` : undefined) ||
    (status === 401
      ? "unauthorized"
      : status === 403
        ? "forbidden"
        : status === 404
          ? "not_found"
          : status === 405
            ? "method_not_allowed"
            : status === 429
              ? "throttled"
              : `http_${status}`);

  const hint =
    status === 401
      ? "token missing, expired or rejected; run auth login"
      : status === 403
        ? "permission denied for this API/audience (Graph app assignment may be blocked)"
        : status === 405
          ? "endpoint or operation not supported by this API version"
          : status === 429
            ? "throttled by the service; retry later"
            : undefined;

  const resolvedRequestId = requestId ?? body?.error?.innerError?.["request-id"];

  return new ApiError({
    status,
    code,
    message: sanitizeMessage(
      body?.error?.message ?? response.statusText ?? `request failed with HTTP ${status}`,
    ),
    retryable: RETRYABLE_STATUS[status] === true,
    ...(retryAfterMs !== undefined ? { retryAfterMs } : {}),
    ...(resolvedRequestId ? { requestId: resolvedRequestId } : {}),
    ...(hint ? { hint } : {}),
  });
}

/** ISO 8601 normalisation for provider timestamps, shared by both clients. */
export function normalizeTimestamp(value: unknown): string | undefined {
  return typeof value === "string" ? isoTime(value) : undefined;
}
