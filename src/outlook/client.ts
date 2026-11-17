/**
 * Outlook REST v2.0 client — the PRIMARY API for mail, calendar and scheduling.
 *
 * Base: `https://outlook.office.com/api/v2.0` (override TEAMS_MCP_OUTLOOK_BASE_URL),
 * token audience `outlook`.
 *
 * The REST payloads are PascalCase (`Subject`, `ReceivedDateTime`,
 * `Start.DateTime`, `InternetMessageHeaders`, …) and collections are wrapped in
 * `value` with `@odata.context`/`@odata.nextLink`. All of that stops here: the
 * client camelises keys and the mappers below produce the normalized shapes the
 * tools and snapshots consume. Nothing above this layer sees PascalCase or an
 * `@odata` envelope.
 */
import { config } from "../config.js";
import { getAccessToken } from "../auth/store.js";
import { isoTime } from "../format.js";
import { parseInstant } from "../time.js";
import {
  HttpClient,
  type FetchLike,
  type HttpClientOptions,
  type TokenProvider,
} from "../http/client.js";

export { ApiError as OutlookError } from "../http/client.js";
export type { FetchLike, TokenProvider } from "../http/client.js";

export interface OutlookClientOptions {
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

export class OutlookClient extends HttpClient {
  constructor(options: OutlookClientOptions = {}) {
    const cfg = config();
    const base: HttpClientOptions = {
      baseUrl: options.baseUrl ?? cfg.outlookBaseUrl,
      tokenProvider:
        options.tokenProvider ?? (async () => ({ token: (await getAccessToken({ audience: "outlook" })).token })),
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

/* ------------------------------------------------------------------ */
/* Payload normalisation                                               */
/* ------------------------------------------------------------------ */

function camelizeKey(key: string): string {
  return key.length === 0 ? key : key.charAt(0).toLowerCase() + key.slice(1);
}

/**
 * Deep-camelise keys. `@odata.*` members are dropped: callers that need the next
 * page read `@odata.nextLink` from the raw payload via `nextLinkOf()`.
 */
export function camelizeKeys<T = unknown>(value: unknown): T {
  if (Array.isArray(value)) return value.map((item) => camelizeKeys(item)) as unknown as T;
  if (value === null || typeof value !== "object") return value as T;
  const out: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
    if (key.startsWith("@")) continue;
    out[camelizeKey(key)] = camelizeKeys(item);
  }
  return out as T;
}

/** `@odata.nextLink` from a raw REST payload. */
export function nextLinkOf(payload: unknown): string | undefined {
  if (payload === null || typeof payload !== "object") return undefined;
  const link = (payload as Record<string, unknown>)["@odata.nextLink"];
  return typeof link === "string" ? link : undefined;
}

export interface Collection<T> {
  items: T[];
  nextLink?: string;
}

/**
 * Object rows from a REST collection `value` array.
 *
 * A malformed upstream `value: [null, 42, "x"]` used to reach the row mappers,
 * which dereference fields and threw `TypeError` for the whole page. One bad row
 * is dropped, not the response.
 */
export function objectRows(value: unknown): Array<Record<string, unknown>> {
  if (!Array.isArray(value)) return [];
  return value.filter(
    (row): row is Record<string, unknown> =>
      row !== null && typeof row === "object" && !Array.isArray(row),
  );
}

/** Normalise a REST collection envelope: `value[]` + `@odata.nextLink`. */
export function readCollection<T>(
  payload: unknown,
  map: (row: Record<string, unknown>) => T,
): Collection<T> {
  if (payload === null || typeof payload !== "object") return { items: [] };
  const rows = objectRows((payload as Record<string, unknown>).value);
  const items = rows.map((row) => map(camelizeKeys<Record<string, unknown>>(row)));
  const nextLink = nextLinkOf(payload);
  return nextLink ? { items, nextLink } : { items };
}

export interface EmailAddress {
  name?: string;
  email?: string;
}

export interface Message {
  id: string;
  conversationId?: string;
  subject: string;
  /** ISO 8601 UTC. */
  received?: string;
  unread?: boolean;
  flagged?: boolean;
  preview: string;
  hasAttachments?: boolean;
  from?: EmailAddress;
  to?: EmailAddress[];
  cc?: EmailAddress[];
  webLink?: string;
  body?: { contentType?: string; content?: string };
  headers?: Array<{ name: string; value: string }>;
}

function addressOf(value: unknown): EmailAddress | undefined {
  if (value === null || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const inner = (record.emailAddress ?? record) as Record<string, unknown>;
  const name = typeof inner.name === "string" ? inner.name : undefined;
  const email = typeof inner.address === "string" ? inner.address : undefined;
  if (!name && !email) return undefined;
  return { ...(name ? { name } : {}), ...(email ? { email } : {}) };
}

function addressList(value: unknown): EmailAddress[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const list = value
    .map((entry) => addressOf(entry))
    .filter((entry): entry is EmailAddress => entry !== undefined);
  return list.length > 0 ? list : undefined;
}

function isoOrUndefined(value: unknown): string | undefined {
  return typeof value === "string" ? isoTime(value) : undefined;
}

/** Normalize one REST message (camelised input). */
export function normalizeMessage(row: Record<string, unknown>): Message {
  const flagStatus =
    row.flag && typeof row.flag === "object"
      ? (row.flag as Record<string, unknown>).flagStatus
      : undefined;
  const headers = Array.isArray(row.internetMessageHeaders)
    ? row.internetMessageHeaders.flatMap((header): Array<{ name: string; value: string }> => {
        if (header === null || typeof header !== "object") return [];
        const record = header as Record<string, unknown>;
        if (typeof record.name !== "string" || typeof record.value !== "string") return [];
        return [{ name: record.name, value: record.value }];
      })
    : undefined;
  const body = row.body as Record<string, unknown> | undefined;

  return {
    id: String(row.id ?? ""),
    ...(typeof row.conversationId === "string" ? { conversationId: row.conversationId } : {}),
    subject: typeof row.subject === "string" ? row.subject : "(no subject)",
    ...(isoOrUndefined(row.receivedDateTime) ? { received: isoOrUndefined(row.receivedDateTime) } : {}),
    ...(typeof row.isRead === "boolean" ? { unread: !row.isRead } : {}),
    ...(typeof flagStatus === "string" ? { flagged: flagStatus !== "notFlagged" } : {}),
    preview: typeof row.bodyPreview === "string" ? row.bodyPreview : "",
    ...(typeof row.hasAttachments === "boolean" ? { hasAttachments: row.hasAttachments } : {}),
    ...(addressOf(row.from) ? { from: addressOf(row.from) } : {}),
    ...(addressList(row.toRecipients) ? { to: addressList(row.toRecipients) } : {}),
    ...(addressList(row.ccRecipients) ? { cc: addressList(row.ccRecipients) } : {}),
    ...(typeof row.webLink === "string" ? { webLink: row.webLink } : {}),
    ...(body && typeof body === "object"
      ? {
          body: {
            ...(typeof body.contentType === "string" ? { contentType: body.contentType } : {}),
            ...(typeof body.content === "string" ? { content: body.content } : {}),
          },
        }
      : {}),
    ...(headers && headers.length > 0 ? { headers } : {}),
  };
}

export interface CalendarEvent {
  id: string;
  subject: string;
  /** ISO 8601 UTC. */
  start: string;
  /** ISO 8601 UTC. */
  end: string;
  allDay?: boolean;
  showAs?: string;
  location?: string;
  organizer?: EmailAddress;
  attendees?: Array<EmailAddress & { type?: string; status?: string }>;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  webLink?: string;
  preview?: string;
}

interface RestDateTime {
  dateTime?: string;
  timeZone?: string;
}

/**
 * REST DateTime -> UTC ISO.
 *
 * A value with an offset is already absolute. A value without one is a wall
 * clock in the payload's own `TimeZone` field — and REST v2 echoes the zone you
 * sent on POST/PATCH, so treating it as UTC shifts write responses by the UTC
 * offset. Convert through the IANA zone when one is supplied; only assume UTC
 * when the payload says UTC (or omits the zone on a UTC-normalised GET).
 */
export function restDateTimeToIso(value: unknown, timeZone?: string): string | undefined {
  if (typeof value !== "string" || value === "") return undefined;
  const trimmed = value.replace(/(\.\d{3})\d+/, "$1");
  if (/[zZ]|[+-]\d{2}:\d{2}$/.test(trimmed)) return isoTime(trimmed);
  if (timeZone && !/^(utc|gmt|z|etc\/utc|etc\/gmt|coordinated universal time)$/i.test(timeZone.trim())) {
    try {
      const instant = parseInstant(trimmed.replace(/\.\d+$/, ""), timeZone);
      if (instant !== undefined) return new Date(instant).toISOString();
    } catch {
      // Unknown zone id: fall through to the documented UTC default.
    }
  }
  return isoTime(`${trimmed}Z`);
}

function dateTimeOf(value: unknown): string | undefined {
  if (value === null || typeof value !== "object") return undefined;
  const record = value as RestDateTime;
  return typeof record.dateTime === "string"
    ? restDateTimeToIso(record.dateTime, record.timeZone)
    : undefined;
}

/** Normalize one REST event (camelised input). */
export function normalizeEvent(row: Record<string, unknown>): CalendarEvent {
  const attendees = Array.isArray(row.attendees)
    ? row.attendees.flatMap((entry) => {
        if (entry === null || typeof entry !== "object") return [];
        const record = entry as Record<string, unknown>;
        const address = addressOf(record.emailAddress) ?? addressOf(record);
        if (!address) return [];
        const status =
          record.status && typeof record.status === "object"
            ? ((record.status as Record<string, unknown>).response as string | undefined)
            : undefined;
        const type = typeof record.type === "string" ? record.type : undefined;
        return [{ ...address, ...(type ? { type } : {}), ...(status ? { status } : {}) }];
      })
    : undefined;

  const location =
    row.location && typeof row.location === "object"
      ? ((row.location as Record<string, unknown>).displayName as string | undefined)
      : undefined;

  return {
    id: String(row.id ?? ""),
    subject: typeof row.subject === "string" ? row.subject : "(no subject)",
    start: dateTimeOf(row.start) ?? "",
    end: dateTimeOf(row.end) ?? "",
    ...(typeof row.isAllDay === "boolean" ? { allDay: row.isAllDay } : {}),
    ...(typeof row.showAs === "string" ? { showAs: row.showAs } : {}),
    ...(location ? { location } : {}),
    ...(addressOf(row.organizer) ? { organizer: addressOf(row.organizer) } : {}),
    ...(attendees && attendees.length > 0 ? { attendees } : {}),
    ...(typeof row.isOnlineMeeting === "boolean" ? { isOnlineMeeting: row.isOnlineMeeting } : {}),
    ...(typeof row.onlineMeetingUrl === "string" ? { onlineMeetingUrl: row.onlineMeetingUrl } : {}),
    ...(typeof row.webLink === "string" ? { webLink: row.webLink } : {}),
    ...(typeof row.bodyPreview === "string" ? { preview: row.bodyPreview } : {}),
  };
}

export interface UnsubscribeInfo {
  mailto?: string;
  url?: string;
  oneClick?: boolean;
  /** Raw header value, for display. */
  raw?: string;
}

/** Parse `List-Unsubscribe` / `List-Unsubscribe-Post` from message headers. */
export function parseUnsubscribe(headers: Array<{ name: string; value: string }> | undefined): UnsubscribeInfo | undefined {
  if (!headers) return undefined;
  const find = (name: string): string | undefined =>
    headers.find((header) => header.name.toLowerCase() === name.toLowerCase())?.value;
  const raw = find("List-Unsubscribe");
  if (!raw) return undefined;
  const targets = [...raw.matchAll(/<([^>]+)>/g)].map((match) => match[1] ?? "");
  const mailto = targets.find((target) => target.toLowerCase().startsWith("mailto:"));
  const url = targets.find((target) => /^https?:/i.test(target));
  const post = find("List-Unsubscribe-Post");
  return {
    ...(mailto ? { mailto } : {}),
    ...(url ? { url } : {}),
    ...(/one-click/i.test(post ?? "") ? { oneClick: true } : {}),
    raw,
  };
}

let shared: OutlookClient | undefined;

export function outlookClient(): OutlookClient {
  shared ??= new OutlookClient();
  return shared;
}

/** Test seam: replace the shared client (pass undefined to reset). */
export function setOutlookClient(client: OutlookClient | undefined): void {
  shared = client;
}
