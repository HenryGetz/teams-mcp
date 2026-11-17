/**
 * Outlook REST v2.0 mail writes.
 *
 * Verified in this tenant: `POST /me/sendmail` → 202 with an empty body, so a
 * successful send has no message id to report. `PATCH /me/messages/{id}`,
 * `POST /me/messages/{id}/{reply|forward|move}` follow the v2 contract.
 */
import { lookup as dnsLookup } from "node:dns/promises";
import { BlockList, isIP } from "node:net";
import { camelizeKeys, type OutlookClient } from "./client.js";

export type BodyType = "Text" | "HTML";
export type Importance = "Low" | "Normal" | "High";

export interface MailDraft {
  subject: string;
  body: string;
  bodyType?: BodyType;
  to: string[];
  cc?: string[];
  bcc?: string[];
  importance?: Importance;
  saveToSentItems?: boolean;
}

function recipients(addresses: string[] | undefined): Array<{ EmailAddress: { Address: string } }> | undefined {
  if (!addresses || addresses.length === 0) return undefined;
  return addresses.map((address) => ({ EmailAddress: { Address: address } }));
}

/** PascalCase `/me/sendmail` body. No `IsReminderSet` anywhere near mail. */
export function buildSendMailBody(draft: MailDraft): Record<string, unknown> {
  return {
    Message: {
      Subject: draft.subject,
      Body: {
        ContentType: draft.bodyType ?? "Text",
        Content: draft.body,
      },
      ToRecipients: recipients(draft.to),
      ...(draft.cc && draft.cc.length > 0 ? { CcRecipients: recipients(draft.cc) } : {}),
      ...(draft.bcc && draft.bcc.length > 0 ? { BccRecipients: recipients(draft.bcc) } : {}),
      ...(draft.importance ? { Importance: draft.importance } : {}),
    },
    SaveToSentItems: draft.saveToSentItems ?? true,
  };
}

export async function sendMail(client: OutlookClient, draft: MailDraft): Promise<void> {
  await client.post<void>("me/sendmail", buildSendMailBody(draft));
}

export async function replyToMessage(client: OutlookClient, messageId: string, comment: string): Promise<void> {
  await client.post<void>(`me/messages/${encodeURIComponent(messageId)}/reply`, { Comment: comment });
}

export async function replyAllToMessage(
  client: OutlookClient,
  messageId: string,
  comment: string,
): Promise<void> {
  await client.post<void>(`me/messages/${encodeURIComponent(messageId)}/replyall`, { Comment: comment });
}

/* ------------------------------------------------------------------ */
/* Drafts                                                              */
/* ------------------------------------------------------------------ */

export interface DraftMessage {
  subject: string;
  body: string;
  bodyType?: BodyType;
  to: string[];
  cc?: string[];
  bcc?: string[];
  importance?: Importance;
}

/**
 * PascalCase REST `Message` shape for `POST /me/messages` (saves to Drafts).
 * Deliberately no `SaveToSentItems`: that key belongs to the sendmail envelope.
 */
export function buildDraftMessage(draft: DraftMessage): Record<string, unknown> {
  return {
    Subject: draft.subject,
    Body: {
      ContentType: draft.bodyType ?? "Text",
      Content: draft.body,
    },
    ToRecipients: recipients(draft.to),
    ...(draft.cc && draft.cc.length > 0 ? { CcRecipients: recipients(draft.cc) } : {}),
    ...(draft.bcc && draft.bcc.length > 0 ? { BccRecipients: recipients(draft.bcc) } : {}),
    ...(draft.importance ? { Importance: draft.importance } : {}),
  };
}

/** Create a draft; returns its id and (when the service sends one) web link. */
export async function createDraft(
  client: OutlookClient,
  draft: DraftMessage,
): Promise<{ id: string; webLink?: string }> {
  const payload = await client.post<Record<string, unknown> | null>("me/messages", buildDraftMessage(draft));
  const created = camelizeKeys<Record<string, unknown>>(payload ?? {});
  return {
    id: typeof created.id === "string" ? created.id : "",
    ...(typeof created.webLink === "string" ? { webLink: created.webLink } : {}),
  };
}

export interface DraftRecipients {
  id?: string;
  subject?: string;
  conversationId?: string;
  to: string[];
  cc: string[];
  bcc: string[];
}

/** Read a draft's current recipients (the allowlist is re-applied at send time). */
export async function getDraftRecipients(client: OutlookClient, draftId: string): Promise<DraftRecipients> {
  const payload = await client.get<Record<string, unknown>>(
    `me/messages/${encodeURIComponent(draftId)}?$select=` +
      encodeURIComponent("Id,ConversationId,Subject,ToRecipients,CcRecipients,BccRecipients"),
  );
  const row = camelizeKeys<Record<string, unknown>>(payload ?? {});
  const addresses = (value: unknown): string[] => {
    if (!Array.isArray(value)) return [];
    return value.flatMap((entry): string[] => {
      if (entry === null || typeof entry !== "object") return [];
      const inner = ((entry as Record<string, unknown>).emailAddress ?? entry) as Record<string, unknown>;
      return typeof inner.address === "string" && inner.address.trim() !== "" ? [inner.address] : [];
    });
  };
  return {
    ...(typeof row.id === "string" ? { id: row.id } : {}),
    ...(typeof row.subject === "string" ? { subject: row.subject } : {}),
    ...(typeof row.conversationId === "string" ? { conversationId: row.conversationId } : {}),
    to: addresses(row.toRecipients),
    cc: addresses(row.ccRecipients),
    bcc: addresses(row.bccRecipients),
  };
}

/** `POST /me/messages/{id}/send`: answers 202, moves the draft to Sent Items. */
export async function sendDraft(client: OutlookClient, draftId: string): Promise<void> {
  await client.post<void>(`me/messages/${encodeURIComponent(draftId)}/send`);
}

/* ------------------------------------------------------------------ */
/* Folders                                                             */
/* ------------------------------------------------------------------ */

export interface MailFolderRow {
  id: string;
  name: string;
  total?: number;
  unread?: number;
  child_folders?: number;
}

/** One page of top-level mail folders with unread/total counts (read-only). */
export async function listFolders(client: OutlookClient, top = 50): Promise<MailFolderRow[]> {
  const params = new URLSearchParams({
    $top: String(top),
    $select: "Id,DisplayName,TotalItemCount,UnreadItemCount,ChildFolderCount",
  });
  const payload = await client.get<Record<string, unknown>>(`me/mailfolders?${params.toString()}`);
  const rows = Array.isArray(payload?.value) ? (payload.value as Array<Record<string, unknown>>) : [];
  return rows.map((row) => {
    const folder = camelizeKeys<Record<string, unknown>>(row);
    return {
      id: typeof folder.id === "string" ? folder.id : "",
      name: typeof folder.displayName === "string" ? folder.displayName : "(unnamed)",
      ...(typeof folder.totalItemCount === "number" ? { total: folder.totalItemCount } : {}),
      ...(typeof folder.unreadItemCount === "number" ? { unread: folder.unreadItemCount } : {}),
      ...(typeof folder.childFolderCount === "number" ? { child_folders: folder.childFolderCount } : {}),
    };
  });
}

export interface ForwardRequest {
  to: string[];
  comment?: string;
}

export function buildForwardBody(request: ForwardRequest): Record<string, unknown> {
  return {
    ...(request.comment ? { Comment: request.comment } : {}),
    ToRecipients: recipients(request.to),
  };
}

export async function forwardMessage(
  client: OutlookClient,
  messageId: string,
  request: ForwardRequest,
): Promise<void> {
  await client.post<void>(
    `me/messages/${encodeURIComponent(messageId)}/forward`,
    buildForwardBody(request),
  );
}

export type MarkTarget = "read" | "unread" | "flag" | "unflag";

/** PascalCase PATCH body for `mark`; only the requested field is sent. */
export function buildMarkPatch(mark: MarkTarget): Record<string, unknown> {
  switch (mark) {
    case "read":
      return { IsRead: true };
    case "unread":
      return { IsRead: false };
    case "flag":
      return { Flag: { FlagStatus: "Flagged" } };
    case "unflag":
      return { Flag: { FlagStatus: "NotFlagged" } };
  }
}

export async function patchMessage(
  client: OutlookClient,
  messageId: string,
  patch: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const payload = await client.patch<Record<string, unknown>>(
    `me/messages/${encodeURIComponent(messageId)}`,
    patch,
  );
  return camelizeKeys<Record<string, unknown>>(payload ?? {});
}

/** `POST /me/messages/{id}/move`; returns the message's NEW id (folders change ids). */
export async function moveMessage(
  client: OutlookClient,
  messageId: string,
  destination: string,
): Promise<{ id: string }> {
  const payload = await client.post<Record<string, unknown>>(
    `me/messages/${encodeURIComponent(messageId)}/move`,
    { DestinationId: destination },
  );
  const moved = camelizeKeys<Record<string, unknown>>(payload ?? {});
  return { id: typeof moved.id === "string" ? moved.id : "" };
}

/** Normalise a subject for prefix comparison: case/whitespace insensitive. */
function normalizeSubject(value: string | undefined): string {
  return (value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}

/** Drop leading `RE:`/`FW:` runs, which providers may collapse or repeat. */
function subjectWithoutPrefix(value: string | undefined): string {
  return normalizeSubject(value).replace(/^(?:(?:re|fw|fwd)\s*:\s*)+/i, "");
}

export interface SentLookupRequest {
  /**
   * Conversation to narrow the search. Omitted for a plain send or a sent draft,
   * where the copy is identified by subject + ReceivedDateTime alone.
   */
  conversationId?: string;
  /** Inclusive lower bound on ReceivedDateTime, epoch ms. */
  since: number;
  /** Subject the service will have assigned, e.g. `RE: <original subject>`. */
  expectedSubject: string;
  /** Delays between lookups while Sent Items catches up (default 400/800/1500 ms). */
  retryDelaysMs?: number[];
}

const DEFAULT_SENT_LOOKUP_DELAYS = [400, 900, 1800];

let sentLookupDelays: number[] = DEFAULT_SENT_LOOKUP_DELAYS;

/** Test seam: replace the Sent Items retry schedule (pass undefined to reset). */
export function setSentLookupDelays(delays: number[] | undefined): void {
  sentLookupDelays = delays ?? DEFAULT_SENT_LOOKUP_DELAYS;
}

function sleep(ms: number): Promise<void> {
  if (ms <= 0) return Promise.resolve();
  const { promise, resolve } = Promise.withResolvers<void>();
  setTimeout(resolve, ms);
  return promise;
}

export type SentLookupResult =
  | { id: string; subject: string }
  | { ambiguous: true; candidates: number }
  | { found: false };

/**
 * Find the copy of a just-sent reply/forward in Sent Items.
 *
 * `POST /me/messages/{id}/reply` answers 202 with no body, so the only way to
 * name the created message is to read it back: same conversation, received in a
 * short window after the POST, and the subject the service assigns (`RE:`/`FW:`
 * prefix). Only an unambiguous single match yields an id — a wrong id is worse
 * than none.
 */
export async function findSentMessage(
  client: OutlookClient,
  request: SentLookupRequest,
): Promise<SentLookupResult> {
  // Sent Items is not instantly queryable; a reply can be missing for a second or
  // two after the 202. Bounded retries, then an honest "not found".
  const delays = request.retryDelaysMs ?? sentLookupDelays;
  let result = await lookupSentMessage(client, request);
  for (const delay of delays) {
    if (!("found" in result)) break;
    await sleep(delay);
    result = await lookupSentMessage(client, request);
  }
  return result;
}

async function lookupSentMessage(
  client: OutlookClient,
  request: SentLookupRequest,
): Promise<SentLookupResult> {
  const filters = [`ReceivedDateTime ge ${new Date(request.since).toISOString()}`];
  if (request.conversationId) {
    filters.unshift(`ConversationId eq '${request.conversationId.replace(/'/g, "''")}'`);
  }
  const params = new URLSearchParams({
    $filter: filters.join(" and "),
    $select: "Id,Subject,ReceivedDateTime,ConversationId",
    $top: "20",
  });
  const payload = await client.get<Record<string, unknown> | null>(
    `me/mailfolders/sentitems/messages?${params.toString()}`,
  );
  const rows =
    payload && Array.isArray(payload.value) ? (payload.value as Array<Record<string, unknown>>) : [];
  const wanted = normalizeSubject(request.expectedSubject);
  const subjectOf = (row: Record<string, unknown>): string | undefined =>
    typeof row.Subject === "string" ? row.Subject : undefined;
  let matches = rows.filter((row) => normalizeSubject(subjectOf(row)) === wanted);
  if (matches.length === 0) {
    // Outlook may collapse `RE: RE: X` to `RE: X`; compare without prefixes as a
    // fallback. Only a prefixed expectation can differ from its base — a plain
    // subject (`send`, `send_draft`) must match exactly, or `RE: X` would look
    // like the copy of `X`. A reply and a forward reduced to the same base are
    // then ambiguous, which is the safe outcome.
    const base = subjectWithoutPrefix(request.expectedSubject);
    if (base !== wanted) {
      matches = rows.filter((row) => base !== "" && subjectWithoutPrefix(subjectOf(row)) === base);
    }
  }
  if (matches.length === 1) {
    const id = matches[0]?.Id;
    const subject = subjectOf(matches[0] ?? {});
    if (typeof id === "string" && subject !== undefined) return { id, subject };
  }
  if (matches.length > 1) return { ambiguous: true, candidates: matches.length };
  return { found: false };
}

/** Fetch a message with the full header set (read-only; used by unsubscribe preview). */
export async function getMessageWithHeaders(
  client: OutlookClient,
  messageId: string,
): Promise<Record<string, unknown>> {
  const payload = await client.get<Record<string, unknown>>(
    `me/messages/${encodeURIComponent(messageId)}?$select=` +
      encodeURIComponent("Id,Subject,From,ReceivedDateTime,InternetMessageHeaders,WebLink"),
  );
  return camelizeKeys<Record<string, unknown>>(payload ?? {});
}

/**
 * One-click unsubscribe POST. Deliberately NOT authenticated: the target is a
 * mailing-list endpoint, and the message's own headers authorise the call. The
 * fetch is injectable so tests never touch the network.
 *
 * The target comes from untrusted email headers, so it is treated as hostile:
 * https only, no embedded credentials, no redirect following, a hard timeout, a
 * bounded body read, and a DNS check that refuses loopback/private/link-local
 * hosts (SSRF).
 */
export type OneClickPoster = (target: string, body: string) => Promise<{ status: number }>;

/** Refused before any connection is made. `code` is stable for callers/tests. */
export class UnsubscribeTargetError extends Error {
  readonly code = "unsubscribe_target_refused";
  constructor(readonly reason: string, message: string) {
    super(message);
    this.name = "UnsubscribeTargetError";
  }
}

export interface OneClickOptions {
  /** DNS resolver seam (tests). Returns every address the host resolves to. */
  resolve?: (host: string) => Promise<string[]>;
}

const DEFAULT_ONE_CLICK_TIMEOUT_MS = 10_000;
const DEFAULT_ONE_CLICK_MAX_BODY_BYTES = 64_000;

/** Hostnames that never name a public mailing-list endpoint. */
const DENIED_HOST_SUFFIXES = [
  ".localhost",
  ".local",
  ".localdomain",
  ".internal",
  ".intranet",
  ".home.arpa",
  ".lan",
  ".test",
  ".invalid",
  ".example",
];
const DENIED_HOSTS: Record<string, true> = {
  localhost: true,
  metadata: true,
  "metadata.google.internal": true,
  "instance-data": true,
};

const blockedV4 = new BlockList();
for (const cidr of [
  "0.0.0.0/8",
  "10.0.0.0/8",
  "100.64.0.0/10",
  "127.0.0.0/8",
  "169.254.0.0/16",
  "172.16.0.0/12",
  "192.0.0.0/24",
  "192.0.2.0/24",
  "192.168.0.0/16",
  "198.18.0.0/15",
  "198.51.100.0/24",
  "203.0.113.0/24",
  "224.0.0.0/4",
  "240.0.0.0/4",
]) {
  const [address, prefix] = cidr.split("/");
  if (address && prefix) blockedV4.addSubnet(address, Number(prefix), "ipv4");
}

const blockedV6 = new BlockList();
for (const cidr of [
  "::/128",
  "::1/128",
  "64:ff9b::/96",
  "2001:db8::/32",
  "fc00::/7",
  "fe80::/10",
  "ff00::/8",
]) {
  const [address, prefix] = cidr.split("/");
  if (address && prefix) blockedV6.addSubnet(address, Number(prefix), "ipv6");
}

/** True for an address that is not a public unicast destination. */
export function isPrivateAddress(address: string): boolean {
  const bare = address.replace(/^\[|\]$/g, "");
  const mapped = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/i.exec(bare);
  if (mapped?.[1]) return blockedV4.check(mapped[1], "ipv4");
  const family = isIP(bare);
  if (family === 4) return blockedV4.check(bare, "ipv4");
  if (family === 6) return blockedV6.check(bare, "ipv6");
  return false;
}

async function defaultResolve(host: string): Promise<string[]> {
  const records = await dnsLookup(host, { all: true, verbatim: true });
  return records.map((record) => record.address);
}

let oneClickResolver: (host: string) => Promise<string[]> = defaultResolve;

/** Test seam: replace the DNS resolver used by the guard (pass undefined to reset). */
export function setOneClickResolver(resolver: ((host: string) => Promise<string[]>) | undefined): void {
  oneClickResolver = resolver ?? defaultResolve;
}

/** Parse and vet an unsubscribe URL. Throws `UnsubscribeTargetError`. */
export function assertSafeUnsubscribeTarget(target: string): URL {
  let url: URL;
  try {
    url = new URL(target);
  } catch {
    throw new UnsubscribeTargetError("invalid_url", `not a valid absolute URL: ${target.slice(0, 120)}`);
  }
  if (url.protocol !== "https:") {
    throw new UnsubscribeTargetError("insecure_scheme", `only https targets are allowed (got ${url.protocol})`);
  }
  if (url.username || url.password) {
    throw new UnsubscribeTargetError("credentials_in_url", "unsubscribe URLs must not embed credentials");
  }
  const host = url.hostname.toLowerCase().replace(/^\[|\]$/g, "");
  if (Object.hasOwn(DENIED_HOSTS, host) || DENIED_HOST_SUFFIXES.some((suffix) => host.endsWith(suffix))) {
    throw new UnsubscribeTargetError("non_public_host", `refusing to POST to ${host}`);
  }
  return url;
}

/** Resolve the host and refuse when any answer is private (defeats DNS tricks). */
async function assertPublicHost(url: URL, resolve: (host: string) => Promise<string[]>): Promise<void> {
  const host = url.hostname.replace(/^\[|\]$/g, "");
  if (isIP(host)) {
    if (isPrivateAddress(host)) {
      throw new UnsubscribeTargetError("private_address", `refusing to POST to ${host}`);
    }
    return;
  }
  let addresses: string[];
  try {
    addresses = await resolve(host);
  } catch (error) {
    throw new UnsubscribeTargetError(
      "dns_failed",
      `could not resolve ${host}: ${(error as Error).message.slice(0, 80)}`,
    );
  }
  if (addresses.length === 0 || addresses.some((address) => isPrivateAddress(address))) {
    throw new UnsubscribeTargetError("private_address", `refusing to POST to ${host} (resolves to a non-public address)`);
  }
}

/** Drain at most `limit` bytes: an unsubscribe reply is tiny, a payload may not be. */
async function drainBounded(response: Response, limit: number): Promise<void> {
  const reader = response.body?.getReader();
  if (!reader) return;
  let received = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) return;
      received += value.byteLength;
      if (received > limit) return;
    }
  } finally {
    await reader.cancel().catch(() => undefined);
  }
}

export const defaultOneClickPoster: OneClickPoster = async (target, body) => {
  const response = await fetch(target, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    redirect: "manual",
    credentials: "omit",
    referrerPolicy: "no-referrer",
    signal: AbortSignal.timeout(DEFAULT_ONE_CLICK_TIMEOUT_MS),
  });
  await drainBounded(response, DEFAULT_ONE_CLICK_MAX_BODY_BYTES);
  return { status: response.status };
};

let oneClickPoster: OneClickPoster = defaultOneClickPoster;

export function setOneClickPoster(poster: OneClickPoster | undefined): void {
  oneClickPoster = poster ?? defaultOneClickPoster;
}

/**
 * Vetted one-click unsubscribe. The guard runs here (not only in the tool), so
 * any caller — including a future one — inherits it.
 */
export async function postOneClickUnsubscribe(
  target: string,
  options: OneClickOptions = {},
): Promise<{ status: number }> {
  const url = assertSafeUnsubscribeTarget(target);
  await assertPublicHost(url, options.resolve ?? oneClickResolver);

  const response = await oneClickPoster(url.toString(), "List-Unsubscribe=One-Click");
  if (response.status >= 300 && response.status < 400) {
    // `redirect: "manual"` keeps the POST on the vetted host; a redirect is a
    // refusal, never a hop.
    throw new UnsubscribeTargetError(
      "redirect_not_followed",
      `target answered HTTP ${response.status} redirect; not followed`,
    );
  }
  return { status: response.status };
}
