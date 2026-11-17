/**
 * `mail` tool: Outlook REST v2.0 (`/me/messages`, `/me/mailfolders/...`).
 *
 * Reads and writes are live. Sends are explicit (recipients/subject/body),
 * allowlist-checked, and never bulk; deletion of mail is not exposed.
 */
import { z } from "zod";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { config } from "../config.js";
import { cutText, htmlToText, truncate } from "../format.js";
import {
  buildMarkPatch,
  createDraft,
  findSentMessage,
  forwardMessage,
  getDraftRecipients,
  getMessageWithHeaders,
  listFolders,
  moveMessage,
  patchMessage,
  postOneClickUnsubscribe,
  replyAllToMessage,
  replyToMessage,
  sendDraft,
  sendMail,
  UnsubscribeTargetError,
} from "../outlook/mail.js";
import {
  camelizeKeys,
  nextLinkOf,
  normalizeMessage,
  outlookClient,
  parseUnsubscribe,
  readCollection,
  type Message,
  type OutlookClient,
} from "../outlook/client.js";
import {
  ATTACHMENT_LIST_CAP,
  ATTACHMENT_TEXT_CHARS,
  decodeContentBytes,
  fetchAttachment,
  isTextualContentType,
  listAttachments,
  MAX_ATTACHMENT_BYTES,
  saveAttachment,
} from "../outlook/attachments.js";
import {
  arrayArg,
  booleanArg,
  disallowedRecipients,
  encodeCursor,
  errorResult,
  numberArg,
  opOf,
  pageSize,
  resolveCursor,
  stringArg,
  textResult,
  toolError,
  unknownOp,
  type ToolModule,
} from "./shared.js";
import { resolveRecipients, isSingleAddress } from "./recipients.js";
import { formatLocal, parseInstantDetailed } from "../time.js";
import {
  acquireWriteToken,
  auditWrite,
  auditedWrite,
  errorOutcome,
  peekUnsubscribeQuota,
  peekWriteToken,
  readOnlyGuard,
  refundWriteToken,
  reserveUnsubscribeSlots,
  writeRateGuard,
} from "../safety.js";

const OPS = [
  "list",
  "search",
  "get",
  "thread",
  "folders",
  "attachments",
  "get_attachment",
  "send",
  "draft",
  "send_draft",
  "reply",
  "forward",
  "mark",
  "move",
  "unsubscribe_preview",
  "unsubscribe_apply",
] as const;

const LIST_SELECT =
  "Id,ConversationId,Subject,From,ReceivedDateTime,IsRead,Flag,HasAttachments,BodyPreview,WebLink";
const DETAIL_SELECT =
  "Id,ConversationId,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,IsRead,Flag,HasAttachments,BodyPreview,Body,InternetMessageHeaders,WebLink";

const BODY_CHARS = 4_000;
/** `thread` bodies are for summarizing, so they stay far below `get`'s cap. */
const THREAD_BODY_CHARS = 500;
const THREAD_BODY_MAX = 2_000;

interface MailRow {
  id: string;
  conversation_id?: string;
  subject: string;
  from?: string;
  received?: string;
  received_local?: string;
  unread?: boolean;
  flag?: boolean;
  preview?: string;
  has_attachments?: boolean;
  web_link?: string;
  unsub?: { has: boolean; one_click?: true; url?: string; mailto?: string };
}

interface RowOptions {
  previewChars: number;
  timezone: string;
  webLink?: boolean;
  unsubscribe?: boolean;
  unsubscribeTargets?: boolean;
  /** list/search: only the fields needed to pick or summarize a message. */
  brief?: boolean;
}

function unsubscribeSummary(
  message: Message,
  includeTargets: boolean | undefined,
): MailRow["unsub"] {
  const parsed = parseUnsubscribe(message.headers);
  if (!parsed) return { has: false };
  return {
    has: true,
    ...(parsed.oneClick ? { one_click: true as const } : {}),
    ...(includeTargets && parsed.url ? { url: parsed.url } : {}),
    ...(includeTargets && parsed.mailto ? { mailto: parsed.mailto } : {}),
  };
}

/**
 * `brief: true` keeps only the fields needed to pick, sort, or summarize a
 * message: the id to read it, the sender/time to judge it, and the unread bit.
 */
function briefRow(message: Message, options: RowOptions): MailRow {
  const from = message.from?.email ?? message.from?.name;
  return {
    id: message.id,
    subject: message.subject,
    ...(from ? { from } : {}),
    ...(message.received ? { received: message.received } : {}),
    ...(message.received ? { received_local: formatLocal(message.received, options.timezone) } : {}),
    unread: message.unread === true,
    ...(message.hasAttachments ? { has_attachments: true } : {}),
  };
}

function mailRow(message: Message, options: RowOptions): MailRow {
  if (options.brief) return briefRow(message, options);
  const from = message.from?.email ?? message.from?.name;
  return {
    id: message.id,
    ...(message.conversationId ? { conversation_id: message.conversationId } : {}),
    subject: message.subject,
    ...(from ? { from } : {}),
    ...(message.received ? { received: message.received } : {}),
    ...(message.received ? { received_local: formatLocal(message.received, options.timezone) } : {}),
    ...(message.unread !== undefined ? { unread: message.unread } : {}),
    ...(message.flagged !== undefined ? { flag: message.flagged } : {}),
    ...(options.previewChars > 0 ? { preview: truncate(message.preview, options.previewChars) } : {}),
    ...(message.hasAttachments === true ? { has_attachments: true } : {}),
    ...(options.webLink && message.webLink ? { web_link: message.webLink } : {}),
    ...(options.unsubscribe
      ? { unsub: unsubscribeSummary(message, options.unsubscribeTargets) }
      : {}),
  };
}

interface ThreadRow extends MailRow {
  body?: string;
  body_truncated?: boolean;
}

function threadRow(
  message: Message,
  options: RowOptions,
  includeBody: boolean,
  bodyChars: number,
): ThreadRow {
  const base = mailRow(message, options);
  const unread = message.unread === true;
  if (!includeBody) return { ...base, unread };
  const content = message.body?.content ?? "";
  // HTML bodies (every reply/forward) need entity decoding and newlines.
  const text =
    message.body?.contentType?.toLowerCase() === "html" ? htmlToText(content) : content;
  return {
    ...base,
    unread,
    body: cutText(text, bodyChars),
    ...(text.length > bodyChars ? { body_truncated: true } : {}),
  };
}

/** Resolved addresses for a write acknowledgement; empty lists are omitted. */
function recipientsField(to: string[], cc: string[] = [], bcc: string[] = []): Record<string, unknown> {
  const lists: Record<string, string[]> = {};
  if (to.length > 0) lists.to = to;
  if (cc.length > 0) lists.cc = cc;
  if (bcc.length > 0) lists.bcc = bcc;
  return Object.keys(lists).length > 0 ? { recipients: lists } : {};
}

/** Quote a value for an OData filter literal. */
function quote(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/**
 * Normalise a caller date into a UTC ISO literal for the OData filter.
 *
 * `since`/`until` used to be interpolated verbatim, so `"… or IsRead eq true"`
 * became part of the provider filter. Canonicalising the parsed instant closes
 * that injection and rejects typo'd dates before any request.
 */
function dateBound(
  args: Record<string, unknown>,
  key: "since" | "until",
  timezone: string,
): { value?: string; result?: CallToolResult } {
  const raw = stringArg(args, key);
  if (!raw) return {};
  const parsed = parseInstantDetailed(raw, timezone);
  if ("error" in parsed) {
    return {
      result: errorResult(
        "invalid_argument",
        `${key} "${raw}" is not an ISO 8601 datetime`,
        "use ISO 8601 (e.g. 2026-11-16T00:00:00Z) or YYYY-MM-DD",
      ),
    };
  }
  return { value: new Date(parsed.instant).toISOString() };
}

function messageFilter(
  args: Record<string, unknown>,
  timezone: string,
): { filter?: string; result?: CallToolResult } {
  const filters: string[] = [];
  if (booleanArg(args, "unread_only")) filters.push("IsRead eq false");
  const from = stringArg(args, "from");
  if (from) filters.push(`From/EmailAddress/Address eq ${quote(from)}`);
  const since = dateBound(args, "since", timezone);
  if (since.result) return { result: since.result };
  const until = dateBound(args, "until", timezone);
  if (until.result) return { result: until.result };
  if (since.value && until.value && Date.parse(since.value) > Date.parse(until.value)) {
    return { result: errorResult("invalid_argument", "since must not be after until") };
  }
  if (since.value) filters.push(`ReceivedDateTime ge ${since.value}`);
  if (until.value) filters.push(`ReceivedDateTime le ${until.value}`);
  return filters.length > 0 ? { filter: filters.join(" and ") } : {};
}

function folderPath(folder: string): string {
  return `me/mailfolders/${encodeURIComponent(folder)}/messages`;
}

/**
 * Id fields for a reply/forward acknowledgement.
 *
 * Outlook answers 202 with no message id, so the sent copy is looked up in
 * Sent Items. When that is not unambiguous the caller gets a `note` instead —
 * never a guessed id.
 */
async function sentMessageFields(
  client: OutlookClient,
  request: { conversationId?: string; expectedSubject: string; since: number },
): Promise<Record<string, unknown>> {
  // `findSentMessage` retries internally while Sent Items catches up.
  const match = await findSentMessage(client, {
    ...(request.conversationId ? { conversationId: request.conversationId } : {}),
    expectedSubject: request.expectedSubject,
    since: request.since,
  });
  if ("id" in match) return { id: match.id };
  if ("ambiguous" in match) {
    return {
      note: `${match.candidates} copies of this "${request.expectedSubject}" message are in Sent Items; read sentitems to pick the exact id`,
    };
  }
  return {
    note: "sent; no single copy could be identified in Sent Items yet, so no id is reported",
  };
}

/** Page size and cap for the `from` scan (Exchange will not sort a From filter). */
const SENDER_SCAN_PAGE = 50;
const SENDER_SCAN_PAGES = 5;

/**
 * Newest-N page for a `From` filter.
 *
 * Exchange rejects `$orderby` with that filter (`InefficientFilter`), and without
 * it a small `$top` was observed to return the *tail* of the match set — so
 * `from=X, top=1` answered "the OLDEST mail from X". The fix: scan whole pages
 * (50 at a time, at most 5), order the union newest-first client-side, then slice
 * the caller's `top`. A cap hit returns `next_cursor` so the scan can continue.
 */
async function senderPage(
  client: OutlookClient,
  firstLink: string | undefined,
  entry: { path: string; filter?: string; select: string },
  top: number,
  folder: string,
  options: RowOptions,
) {
  let link: string | undefined =
    firstLink ??
    `${entry.path}?${new URLSearchParams({
      ...(entry.filter ? { $filter: entry.filter } : {}),
      $select: entry.select,
      $top: String(SENDER_SCAN_PAGE),
    }).toString()}`;
  const collected: Message[] = [];
  let scanned = 0;
  let nextLink: string | undefined;
  while (link && scanned < SENDER_SCAN_PAGES) {
    const payload = await client.get<unknown>(link);
    scanned += 1;
    collected.push(...readCollection(payload, normalizeMessage).items);
    nextLink = nextLinkOf(payload);
    link = nextLink;
  }
  const items = collected
    .sort((a, b) => (b.received ?? "").localeCompare(a.received ?? ""))
    .slice(0, top);
  return textResult({
    folder,
    count: items.length,
    messages: items.map((message) => mailRow(message, options)),
    scanned_pages: scanned,
    next_cursor: nextLink ? encodeCursor(nextLink) : undefined,
    // More matches than returned, or the scan stopped before exhausting them.
    truncated: nextLink !== undefined || collected.length > top,
  });
}

async function listPage(
  client: OutlookClient,
  path: string,
  args: Record<string, unknown>,
  options: RowOptions,
) {
  const top = pageSize(args, "top", 10, config().maxResults.mail);
  const cursorLookup = resolveCursor(args);
  if (cursorLookup.result) return cursorLookup.result;
  const cursor = cursorLookup.link;
  if (cursor && !client.ownsUrl(cursor)) {
    return errorResult("invalid_cursor", "cursor does not belong to Outlook REST");
  }
  const select = `${LIST_SELECT}${options.unsubscribe && !options.brief ? ",InternetMessageHeaders" : ""}`;
  const filter = messageFilter(args, config().timezone);
  if (filter.result) return filter.result;
  const folder = stringArg(args, "folder") ?? "inbox";

  if (stringArg(args, "from") !== undefined) {
    return senderPage(
      client,
      cursor,
      { path, ...(filter.filter ? { filter: filter.filter } : {}), select },
      top,
      folder,
      options,
    );
  }

  const params = new URLSearchParams({ $select: select, $top: String(top) });
  if (filter.filter) params.set("$filter", filter.filter);
  params.set("$orderby", "ReceivedDateTime desc");

  const payload = await client.get<unknown>(cursor ?? `${path}?${params.toString()}`);
  const collection = readCollection(payload, normalizeMessage);
  const nextLink = nextLinkOf(payload);

  return textResult({
    folder,
    count: collection.items.length,
    messages: collection.items.map((message) => mailRow(message, options)),
    next_cursor: nextLink ? encodeCursor(nextLink) : undefined,
    truncated: nextLink !== undefined,
  });
}

export const mailTool: ToolModule = {
  name: "mail",
  title: "Mail",
  description:
    "Outlook mail (REST v2). op=list (folder inbox; top<=50; unread_only; from, which scans newest-first; since/until; cursor; brief) | search (query) | get (message_id; body, body_chars<=20000) | thread (thread_id) | folders (per-folder unread/total) | attachments (metadata, no content) | get_attachment (text/* decoded; files saved under cache/) | send | draft | send_draft (confirm; allowlist re-checked) | reply (reply_all) | forward | mark | move | unsubscribe_preview | unsubscribe_apply (confirm, expected_sender). Rows: id, conversation_id, subject, from, received, unread, preview<=preview_chars (default 300), has_attachments; web_link and unsub are opt-in.",
  shape: {
    op: z.enum(OPS),
    folder: z.string().optional().describe("inbox (default), sentitems, drafts, archive, deleteditems, or folder id"),
    query: z.string().optional().describe("search: free text"),
    top: z.number().int().min(1).max(50).optional().describe("page size, default 10, max 50 (capped by max_results.mail, default 25)"),
    cursor: z.string().optional().describe("opaque next_cursor from a prior call"),
    unread_only: z.boolean().optional().describe("list: only unread messages"),
    from: z.string().optional().describe("filter: sender address"),
    since: z.string().optional().describe("ISO 8601 lower bound on received time"),
    until: z.string().optional().describe("ISO 8601 upper bound on received time"),
    brief: z.boolean().optional().describe("list/search: compact rows (id,subject,from,received,received_local,unread,has_attachments?)"),
    body_format: z.enum(["text", "html"]).optional().describe("get: preferred body format, default text"),
    body_chars: z.number().int().min(100).max(20000).optional().describe("get: cap default 4000 (max 20000); thread: cap default 500, max 2000"),
    include_body: z.boolean().optional().describe("thread: include each message body text, default false"),
    include_body_preview: z.boolean().optional().describe("list/search: attach preview, default true"),
    preview_chars: z.number().int().min(40).max(1000).optional().describe("list/search: preview cap, default max_results.preview_chars; include_body_preview=false wins"),
    include_web_link: z.boolean().optional().describe("list/search/thread: attach web_link, default false"),
    include_unsubscribe: z
      .boolean()
      .optional()
      .describe("add unsub{has,one_click}"),
    include_unsubscribe_targets: z
      .boolean()
      .optional()
      .describe("inline unsub url/mailto"),
    message_id: z.string().optional().describe("message id for get/attachments/reply/mark/move"),
    attachment_id: z.string().optional().describe("get_attachment: id from op=attachments"),
    text: z
      .boolean()
      .optional()
      .describe("get_attachment: force UTF-8 decode"),
    thread_id: z.string().optional().describe("thread: conversation id"),
    to: z.array(z.string()).optional().describe("send/draft/forward: recipients"),
    cc: z.array(z.string()).optional().describe("send/draft: cc"),
    bcc: z.array(z.string()).optional().describe("send/draft: bcc"),
    subject: z.string().optional().describe("send/draft: subject"),
    body: z.string().optional().describe("send/draft/reply/forward: body text"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("send/draft: importance"),
    draft_id: z.string().optional().describe("send_draft: id of a draft saved by op=draft"),
    reply_all: z.boolean().optional().describe("reply: reply to everyone on the original (allowlist-checked)"),
    mark: z.enum(["read", "unread", "flag", "unflag"]).optional().describe("mark: target state"),
    destination: z
      .string()
      .optional()
      .describe("move: folder id or name; one message per call"),
    confirm: z.boolean().optional().describe("send_draft/unsubscribe_apply: must be true"),
    expected_sender: z
      .string()
      .optional()
      .describe("unsubscribe_apply: the From address you expect; a mismatch refuses"),
  },
  async handle(args) {
    const op = opOf(args, OPS);
    if (!op) return unknownOp("mail", String(args.op ?? ""), OPS);
    const blocked = readOnlyGuard("mail", op, args);
    if (blocked) return blocked;

    const cfg = config();
    const requestedPreview = numberArg(args, "preview_chars");
    // `include_body_preview:false` always wins: no preview, whatever the cap.
    const previewChars =
      booleanArg(args, "include_body_preview") === false
        ? 0
        : requestedPreview === undefined
          ? cfg.maxResults.previewChars
          : Math.min(Math.max(40, Math.trunc(requestedPreview)), 1_000);
    const rowOptions: RowOptions = {
      previewChars,
      timezone: cfg.timezone,
      webLink: booleanArg(args, "include_web_link") === true,
      unsubscribe: booleanArg(args, "include_unsubscribe") === true,
      unsubscribeTargets: booleanArg(args, "include_unsubscribe_targets") === true,
      brief: booleanArg(args, "brief") === true,
    };
    const client = outlookClient();

    try {
      if (op === "list") {
        return await listPage(client, folderPath(stringArg(args, "folder") ?? "inbox"), args, rowOptions);
      }

      if (op === "search") {
        const query = stringArg(args, "query");
        if (!query) return errorResult("invalid_argument", "search requires query");
        const cursorLookup = resolveCursor(args);
        if (cursorLookup.result) return cursorLookup.result;
        const cursor = cursorLookup.link;
        const top = pageSize(args, "top", 10, cfg.maxResults.mail);
        if (cursor) {
          if (!client.ownsUrl(cursor)) return errorResult("invalid_cursor", "cursor does not belong to Outlook REST");
          const payload = await client.get<unknown>(cursor);
          const collection = readCollection(payload, normalizeMessage);
          const nextLink = nextLinkOf(payload);
          return textResult({
            count: collection.items.length,
            messages: collection.items.map((message) => mailRow(message, rowOptions)),
            next_cursor: nextLink ? encodeCursor(nextLink) : undefined,
            truncated: nextLink !== undefined,
          });
        }
        const params = new URLSearchParams({
          $search: `"${query.replace(/"/g, "")}"`,
          $select: `${LIST_SELECT}${rowOptions.unsubscribe && !rowOptions.brief ? ",InternetMessageHeaders" : ""}`,
          $top: String(top),
        });
        const payload = await client.get<unknown>(`me/messages?${params.toString()}`);
        const collection = readCollection(payload, normalizeMessage);
        const nextLink = nextLinkOf(payload);
        return textResult({
          query,
          count: collection.items.length,
          messages: collection.items.map((message) => mailRow(message, rowOptions)),
          next_cursor: nextLink ? encodeCursor(nextLink) : undefined,
          truncated: nextLink !== undefined,
        });
      }

      if (op === "get") {
        const messageId = stringArg(args, "message_id");
        if (!messageId) return errorResult("invalid_argument", "get requires message_id");
        const payload = await client.get<unknown>(
          `me/messages/${encodeURIComponent(messageId)}?$select=${encodeURIComponent(DETAIL_SELECT)}`,
        );
        const message = normalizeMessage(camelizeKeys<Record<string, unknown>>(payload));
        const wanted = args.body_format === "html" ? "html" : "text";
        const bodyChars = Math.min(Math.max(100, numberArg(args, "body_chars") ?? BODY_CHARS), 20_000);
        const content = message.body?.content ?? "";
        // HTML bodies (every reply/forward) need entity decoding and newlines;
        // text bodies are already plain.
        const bodyText =
          wanted === "text" && message.body?.contentType?.toLowerCase() === "html"
            ? htmlToText(content)
            : content;
        return textResult({
          ...mailRow(message, {
            previewChars: cfg.maxResults.previewChars,
            timezone: cfg.timezone,
            webLink: true,
          }),
          to: message.to?.map((entry) => entry.email ?? entry.name),
          cc: message.cc?.map((entry) => entry.email ?? entry.name),
          body_format: wanted,
          body: cutText(bodyText, bodyChars),
          body_truncated: bodyText.length > bodyChars,
        });
      }

      if (op === "thread") {
        const threadId = stringArg(args, "thread_id");
        if (!threadId) return errorResult("invalid_argument", "thread requires thread_id");
        const cursorLookup = resolveCursor(args);
        if (cursorLookup.result) return cursorLookup.result;
        const cursor = cursorLookup.link;
        if (cursor && !client.ownsUrl(cursor)) {
          return errorResult("invalid_cursor", "cursor does not belong to Outlook REST");
        }
        const top = pageSize(args, "top", 10, cfg.maxResults.mail);
        const includeBody = booleanArg(args, "include_body") === true;
        const bodyChars = Math.min(
          Math.max(100, Math.trunc(numberArg(args, "body_chars") ?? THREAD_BODY_CHARS)),
          THREAD_BODY_MAX,
        );
        const threadOptions: RowOptions = {
          previewChars: cfg.maxResults.previewChars,
          timezone: cfg.timezone,
          webLink: rowOptions.webLink,
        };
        // `$orderby ReceivedDateTime` together with a ConversationId filter is
        // rejected live with InefficientFilter, so sort the page locally.
        let payload: unknown;
        if (cursor) {
          payload = await client.get<unknown>(cursor);
        } else {
          const params = new URLSearchParams({
            $filter: `ConversationId eq ${quote(threadId)}`,
            // Body is a large field: request it only when the caller asked.
            $select: includeBody ? `${LIST_SELECT},Body` : LIST_SELECT,
            $top: String(top),
          });
          payload = await client.get<unknown>(`me/messages?${params.toString()}`);
        }
        const collection = readCollection(payload, normalizeMessage);
        const nextLink = nextLinkOf(payload);
        const rows = collection.items
          .map((message) => threadRow(message, threadOptions, includeBody, bodyChars))
          .sort((a, b) => (a.received ?? "").localeCompare(b.received ?? ""));
        return textResult({
          thread_id: threadId,
          count: rows.length,
          messages: rows,
          next_cursor: nextLink ? encodeCursor(nextLink) : undefined,
          // Without this a `top`-limited conversation read looks complete.
          truncated: nextLink !== undefined,
        });
      }

      if (op === "folders") {
        const folders = await listFolders(client, 50);
        return textResult({
          count: folders.length,
          folders,
          // One page only; a full page likely means more folders exist.
          truncated: folders.length >= 50,
        });
      }

      if (op === "attachments") {
        const messageId = stringArg(args, "message_id");
        if (!messageId) return errorResult("invalid_argument", "attachments requires message_id");
        const listed = await listAttachments(client, messageId);
        return textResult({
          message_id: messageId,
          count: listed.attachments.length,
          attachments: listed.attachments,
          truncated: listed.truncated,
          ...(listed.truncated
            ? { note: `showing the first ${ATTACHMENT_LIST_CAP}; call get_attachment for one by id` }
            : {}),
        });
      }

      if (op === "get_attachment") {
        const messageId = stringArg(args, "message_id");
        const attachmentId = stringArg(args, "attachment_id");
        if (!messageId) return errorResult("invalid_argument", "get_attachment requires message_id");
        if (!attachmentId) return errorResult("invalid_argument", "get_attachment requires attachment_id");

        const fetched = await fetchAttachment(client, messageId, attachmentId);
        const meta = fetched.meta;
        if (fetched.contentBytes === undefined) {
          return textResult({ message_id: messageId, ...meta, note: "content_unavailable" });
        }
        // A declared size above the cap is refused before decoding; the actual
        // decoded length is checked again below in case the service lied.
        if (meta.size > MAX_ATTACHMENT_BYTES) {
          return textResult({
            message_id: messageId,
            ...meta,
            note: "too_large",
            limit_bytes: MAX_ATTACHMENT_BYTES,
          });
        }
        const decoded = decodeContentBytes(fetched.contentBytes);
        if ("error" in decoded) {
          return errorResult(
            "invalid_attachment_content",
            `attachment ${attachmentId}: ${decoded.error}`,
            "the service returned ContentBytes that cannot be base64-decoded; nothing was saved",
          );
        }
        const bytes = decoded.bytes;
        if (bytes.length > MAX_ATTACHMENT_BYTES) {
          return textResult({
            message_id: messageId,
            ...meta,
            size: bytes.length,
            note: "too_large",
            limit_bytes: MAX_ATTACHMENT_BYTES,
          });
        }

        // Text-ish media is decoded in place; `text:true` is an explicit
        // override for a textual type the media type does not advertise.
        if (isTextualContentType(meta.content_type) || booleanArg(args, "text") === true) {
          const text = bytes.toString("utf8");
          return textResult({
            message_id: messageId,
            ...meta,
            size: bytes.length,
            text: text.slice(0, ATTACHMENT_TEXT_CHARS),
            text_truncated: text.length > ATTACHMENT_TEXT_CHARS,
          });
        }

        const path = await saveAttachment(config().cacheDir, messageId, attachmentId, meta.name, bytes);
        return textResult({
          message_id: messageId,
          ...meta,
          size: bytes.length,
          saved: true,
          path,
          note: "saved to the cache; the file was not opened or executed",
        });
      }

      if (op === "unsubscribe_preview") {
        const messageId = stringArg(args, "message_id");
        if (!messageId) return errorResult("invalid_argument", "unsubscribe_preview requires message_id");
        const message = normalizeMessage(await getMessageWithHeaders(client, messageId));
        const unsubscribe = parseUnsubscribe(message.headers);
        return textResult({
          message_id: message.id,
          subject: message.subject,
          from: message.from?.email,
          web_link: message.webLink,
          unsubscribe,
          ...(unsubscribe
            ? { note: "read-only preview; unsubscribe_apply needs confirm=true" }
            : { note: "message has no List-Unsubscribe header" }),
        });
      }

      if (op === "unsubscribe_apply") {
        const messageId = stringArg(args, "message_id");
        if (!messageId) return errorResult("invalid_argument", "unsubscribe_apply requires message_id");

        // Gate order is deliberate: every refusal that needs no message data
        // happens before the header read, so a disabled/capped/rate-limited call
        // issues no network request at all.
        if (!cfg.unsubscribe.enabled) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: messageId },
            outcome: "unsubscribe_disabled",
          });
          return errorResult(
            "unsubscribe_disabled",
            "unsubscribe_apply is disabled (safety.unsubscribe.enabled=false); nothing was sent",
            "an operator must enable it in config.yaml or with TEAMS_MCP_UNSUBSCRIBE_ENABLED=1; use unsubscribe_preview to inspect the targets",
          );
        }

        const quota = await peekUnsubscribeQuota();
        if (!quota.allowed) {
          auditWrite({ tool: "mail", op, target: { message_id: messageId }, outcome: "cap_reached" });
          return errorResult(
            "cap_reached",
            `daily unsubscribe cap reached (${quota.used}/${quota.limit}); no external action was taken`,
            `the counter resets at ${quota.reset_at}`,
          );
        }

        const peeked = peekWriteToken();
        if (!peeked.ok) {
          auditWrite({ tool: "mail", op, target: { message_id: messageId }, outcome: "rate_limited" });
          return errorResult(
            "rate_limited",
            `write rate limit reached (${cfg.writeRateLimitPerMinute}/minute); nothing was sent`,
            `retry in ${peeked.retryAfterSeconds}s`,
          );
        }

        const message = normalizeMessage(await getMessageWithHeaders(client, messageId));
        const unsubscribe = parseUnsubscribe(message.headers);
        const preview = {
          message_id: message.id,
          subject: message.subject,
          from: message.from?.email,
          unsubscribe,
        };
        if (!unsubscribe) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id },
            outcome: "nothing_to_unsubscribe",
          });
          return errorResult("nothing_to_unsubscribe", "message has no List-Unsubscribe header");
        }
        if (booleanArg(args, "confirm") !== true) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id },
            outcome: "confirm_required",
          });
          return textResult({
            refused: true,
            reason: "confirm=true required; nothing was sent",
            preview,
          });
        }
        const expectedSender = stringArg(args, "expected_sender");
        const sender = message.from?.email;
        if (!expectedSender) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id },
            outcome: "expected_sender_required",
          });
          return textResult({
            refused: true,
            reason:
              "expected_sender is required; pass the From address you intend to unsubscribe so a stale id cannot act on a different sender",
            preview,
          });
        }
        if (!sender || sender.toLowerCase() !== expectedSender.toLowerCase()) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id },
            outcome: "sender_mismatch",
          });
          return textResult({
            refused: true,
            reason: `expected_sender "${expectedSender}" does not match this message's From address${sender ? ` (${sender})` : " (none)"}; nothing was sent`,
            preview,
          });
        }

        // One rate token and one cap slot per external action, reserved
        // immediately before acting. A message with both an https URL and a
        // mailto target therefore costs two slots, so the daily cap bounds
        // external actions rather than apply calls.
        const mailtoAddress = unsubscribe.mailto
          ? (unsubscribe.mailto.replace(/^mailto:/i, "").split("?")[0]?.trim() ?? "")
          : "";
        const plannedTargets =
          (unsubscribe.url ? 1 : 0) +
          (mailtoAddress !== "" && isSingleAddress(mailtoAddress) ? 1 : 0);
        if (plannedTargets === 0) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id },
            outcome: "nothing_to_unsubscribe",
          });
          return errorResult(
            "nothing_to_unsubscribe",
            "the List-Unsubscribe header has no usable target; nothing was sent",
          );
        }

        const token = acquireWriteToken();
        if (!token.ok) {
          auditWrite({ tool: "mail", op, target: { message_id: message.id }, outcome: "rate_limited" });
          return errorResult(
            "rate_limited",
            `write rate limit reached (${cfg.writeRateLimitPerMinute}/minute); nothing was sent`,
            `retry in ${token.retryAfterSeconds}s`,
          );
        }
        const reserved = await reserveUnsubscribeSlots(plannedTargets);
        if (!reserved.allowed) {
          refundWriteToken();
          auditWrite({ tool: "mail", op, target: { message_id: message.id }, outcome: "cap_reached" });
          return errorResult(
            "cap_reached",
            `daily unsubscribe cap reached (${reserved.used}/${reserved.limit}); ${plannedTargets} external action(s) would exceed it, nothing was sent`,
            `the counter resets at ${reserved.reset_at}`,
          );
        }

        try {
          const results: Array<{
            target: string;
            type: "https" | "mailto";
            status: string;
            http_status?: number;
            refused_reason?: string;
          }> = [];
          if (unsubscribe.url) {
            if (!unsubscribe.url.startsWith("https://")) {
              results.push({ target: unsubscribe.url, type: "https", status: "refused_insecure_target" });
            } else {
              try {
                const { status } = await postOneClickUnsubscribe(unsubscribe.url);
                results.push({
                  target: unsubscribe.url,
                  type: "https",
                  status: status >= 200 && status < 300 ? "posted" : "failed",
                  http_status: status,
                });
              } catch (error) {
                if (error instanceof UnsubscribeTargetError) {
                  results.push({
                    target: unsubscribe.url,
                    type: "https",
                    status: "refused",
                    refused_reason: error.reason,
                  });
                } else {
                  results.push({
                    target: unsubscribe.url,
                    type: "https",
                    status: `failed: ${(error as Error).message.slice(0, 120)}`,
                  });
                }
              }
            }
          }
          if (unsubscribe.mailto) {
            const address = unsubscribe.mailto.replace(/^mailto:/i, "").split("?")[0]?.trim() ?? "";
            if (!address || !isSingleAddress(address)) {
              results.push({ target: unsubscribe.mailto, type: "mailto", status: "invalid_address" });
            } else {
              await sendMail(client, { subject: "unsubscribe", body: "unsubscribe", to: [address] });
              results.push({ target: address, type: "mailto", status: "sent" });
            }
          }
          const ok = results.length > 0 && results.every((entry) => entry.status === "posted" || entry.status === "sent");
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id, targets: results.length },
            outcome: ok ? "ok" : "failed",
          });
          return textResult({ ok, message_id: message.id, results, partial: !ok && results.length > 1 });
        } catch (error) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: message.id },
            outcome: errorOutcome(error),
          });
          throw error;
        }
      }

      if (op === "send") {
        const rawTo = arrayArg(args, "to") ?? [];
        const rawCc = arrayArg(args, "cc") ?? [];
        const rawBcc = arrayArg(args, "bcc") ?? [];
        const subject = stringArg(args, "subject");
        const body = typeof args.body === "string" ? args.body : undefined;
        if (rawTo.length === 0 || !subject || body === undefined) {
          return errorResult("invalid_argument", "send requires to[], subject and body");
        }
        // Rate gate before directory resolution: a rate-limited send makes no
        // request at all, and an invalid call above consumed no token.
        const limited = writeRateGuard("mail", op, {
          recipients: rawTo.length + rawCc.length + rawBcc.length,
        });
        if (limited) return limited;
        const resolved = await resolveRecipients([...rawTo, ...rawCc, ...rawBcc]);
        if (!resolved.ok) {
          refundWriteToken();
          return resolved.result;
        }
        const to = resolved.addresses.slice(0, rawTo.length);
        const cc = resolved.addresses.slice(rawTo.length, rawTo.length + rawCc.length);
        const bcc = resolved.addresses.slice(rawTo.length + rawCc.length);
        const blocked = disallowedRecipients([...to, ...cc, ...bcc]);
        if (blocked.length > 0) {
          refundWriteToken();
          auditWrite({
            tool: "mail",
            op,
            target: { recipients: blocked.length },
            outcome: "recipient_not_allowed",
          });
          return errorResult(
            "recipient_not_allowed",
            `recipient(s) outside invite_allowlist: ${blocked.join(", ")}`,
            "edit invite_allowlist in settings",
          );
        }
        const importance =
          args.importance === "high" ? "High" : args.importance === "low" ? "Low" : undefined;
        const recipientCount = to.length + cc.length + bcc.length;
        const startedAt = Date.now() - 2000;
        await auditedWrite("mail", op, { recipients: recipientCount }, () =>
          sendMail(client, {
            subject,
            body,
            to,
            ...(cc.length > 0 ? { cc } : {}),
            ...(bcc.length > 0 ? { bcc } : {}),
            ...(importance ? { importance } : {}),
          }),
        );
        // 202 carries no body; identify the copy by exact subject + receive time.
        return textResult({
          ok: true,
          op: "send",
          recipient_count: recipientCount,
          ...recipientsField(to, cc, bcc),
          ...(await sentMessageFields(client, { expectedSubject: subject, since: startedAt })),
        });
      }

      if (op === "draft") {
        const rawTo = arrayArg(args, "to") ?? [];
        const rawCc = arrayArg(args, "cc") ?? [];
        const rawBcc = arrayArg(args, "bcc") ?? [];
        const subject = stringArg(args, "subject");
        const body = typeof args.body === "string" ? args.body : undefined;
        if (rawTo.length === 0 || !subject || body === undefined) {
          return errorResult("invalid_argument", "draft requires to[], subject and body");
        }
        // A draft is a write (it lands in Drafts) but never leaves the mailbox.
        const limited = writeRateGuard("mail", op, {
          recipients: rawTo.length + rawCc.length + rawBcc.length,
        });
        if (limited) return limited;
        const resolved = await resolveRecipients([...rawTo, ...rawCc, ...rawBcc]);
        if (!resolved.ok) {
          refundWriteToken();
          return resolved.result;
        }
        const to = resolved.addresses.slice(0, rawTo.length);
        const cc = resolved.addresses.slice(rawTo.length, rawTo.length + rawCc.length);
        const bcc = resolved.addresses.slice(rawTo.length + rawCc.length);
        const blocked = disallowedRecipients([...to, ...cc, ...bcc]);
        if (blocked.length > 0) {
          refundWriteToken();
          auditWrite({
            tool: "mail",
            op,
            target: { recipients: blocked.length },
            outcome: "recipient_not_allowed",
          });
          return errorResult(
            "recipient_not_allowed",
            `recipient(s) outside invite_allowlist: ${blocked.join(", ")}`,
            "edit invite_allowlist in settings",
          );
        }
        const importance =
          args.importance === "high" ? "High" : args.importance === "low" ? "Low" : undefined;
        const recipientCount = to.length + cc.length + bcc.length;
        const created = await auditedWrite("mail", op, { recipients: recipientCount }, () =>
          createDraft(client, {
            subject,
            body,
            to,
            ...(cc.length > 0 ? { cc } : {}),
            ...(bcc.length > 0 ? { bcc } : {}),
            ...(importance ? { importance } : {}),
          }),
        );
        return textResult({
          ok: true,
          op: "draft",
          recipient_count: recipientCount,
          ...(created.id ? { id: created.id } : {}),
          ...(created.webLink ? { web_link: created.webLink } : {}),
          note: "saved to Drafts; nothing was sent",
        });
      }

      if (op === "send_draft") {
        const draftId = stringArg(args, "draft_id");
        if (!draftId) return errorResult("invalid_argument", "send_draft requires draft_id");
        if (booleanArg(args, "confirm") !== true) {
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: draftId },
            outcome: "confirmation_required",
          });
          return errorResult(
            "confirmation_required",
            `mail.send_draft would send draft ${draftId}; nothing was sent`,
            "review the draft, then re-call with confirm:true",
          );
        }
        const limited = writeRateGuard("mail", op, { message_id: draftId });
        if (limited) return limited;
        // A draft edited in Outlook can name recipients that were never checked,
        // so the allowlist is re-applied against the draft as it is NOW.
        let draft;
        try {
          draft = await getDraftRecipients(client, draftId);
        } catch (error) {
          refundWriteToken();
          throw error;
        }
        const recipients = [...draft.to, ...draft.cc, ...draft.bcc];
        if (recipients.length === 0) {
          refundWriteToken();
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: draftId },
            outcome: "no_recipients",
          });
          return errorResult(
            "invalid_argument",
            "draft has no recipients; nothing was sent",
            "open the draft in Outlook (or create a new one) and add recipients",
          );
        }
        const blocked = disallowedRecipients(recipients);
        if (blocked.length > 0) {
          refundWriteToken();
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: draftId, recipients: blocked.length },
            outcome: "recipient_not_allowed",
          });
          return errorResult(
            "recipient_not_allowed",
            `draft recipient(s) outside invite_allowlist: ${blocked.join(", ")}; nothing was sent`,
            "remove those recipients from the draft or edit invite_allowlist in settings",
          );
        }
        const startedAt = Date.now() - 2000;
        await auditedWrite("mail", op, { message_id: draftId, recipients: recipients.length }, () =>
          sendDraft(client, draftId),
        );
        return textResult({
          ok: true,
          op: "send_draft",
          draft_id: draftId,
          recipient_count: recipients.length,
          ...recipientsField(draft.to, draft.cc, draft.bcc),
          ...(draft.subject !== undefined && draft.subject !== ""
            ? await sentMessageFields(client, {
                ...(draft.conversationId ? { conversationId: draft.conversationId } : {}),
                expectedSubject: draft.subject,
                since: startedAt,
              })
            : { note: "sent; the draft had no subject, so no Sent Items id was looked up" }),
        });
      }

      if (op === "reply") {
        const messageId = stringArg(args, "message_id");
        const body = typeof args.body === "string" ? args.body : undefined;
        if (!messageId || body === undefined) {
          return errorResult("invalid_argument", "reply requires message_id and body");
        }
        const replyAll = booleanArg(args, "reply_all") === true;
        // Rate gate before the destination lookup so a refused reply is fully
        // offline; refund if that lookup fails (nothing was sent).
        const limited = writeRateGuard("mail", op, { message_id: messageId });
        if (limited) return limited;
        // Read the original so the success report reflects a real message, and the
        // reply has somewhere to go. Reply recipients are derived by the service.
        let original: unknown;
        try {
          original = await client.get<unknown>(
            `me/messages/${encodeURIComponent(messageId)}?$select=` +
              encodeURIComponent("Id,ConversationId,Subject,From,ToRecipients,CcRecipients"),
          );
        } catch (error) {
          refundWriteToken();
          throw error;
        }
        const message = normalizeMessage(camelizeKeys<Record<string, unknown>>(original ?? {}));
        // Reply-all reaches everyone the original did. With an allowlist in
        // force, every derived address is checked BEFORE the service derives
        // (and sends to) the same set.
        const replyTo: string[] = [];
        const replyCc: string[] = [];
        const seen = new Set<string>();
        if (replyAll) {
          const add = (into: string[], address: string | undefined): void => {
            if (typeof address !== "string" || address.trim() === "") return;
            const key = address.trim().toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            into.push(address.trim());
          };
          add(replyTo, message.from?.email);
          for (const entry of message.to ?? []) add(replyTo, entry.email);
          for (const entry of message.cc ?? []) add(replyCc, entry.email);
          const blocked = disallowedRecipients([...replyTo, ...replyCc]);
          if (blocked.length > 0) {
            refundWriteToken();
            auditWrite({
              tool: "mail",
              op,
              target: { message_id: messageId, recipients: blocked.length },
              outcome: "recipient_not_allowed",
            });
            return errorResult(
              "recipient_not_allowed",
              `reply_all would reach recipient(s) outside invite_allowlist: ${blocked.join(", ")}; nothing was sent`,
              "reply only to the sender, or edit invite_allowlist in settings",
            );
          }
        } else if (message.from?.email) {
          replyTo.push(message.from.email);
        }
        const recipientCount = replyAll ? replyTo.length + replyCc.length : message.from ? 1 : 0;
        const startedAt = Date.now() - 2000;
        await auditedWrite("mail", op, { message_id: messageId, recipients: recipientCount }, () =>
          replyAll ? replyAllToMessage(client, messageId, body) : replyToMessage(client, messageId, body),
        );
        return textResult({
          ok: true,
          op: "reply",
          ...(replyAll ? { reply_all: true } : {}),
          recipient_count: recipientCount,
          ...recipientsField(replyTo, replyCc),
          ...(await sentMessageFields(client, {
            ...(message.conversationId ? { conversationId: message.conversationId } : {}),
            expectedSubject: `RE: ${message.subject}`,
            since: startedAt,
          })),
        });
      }

      if (op === "forward") {
        const messageId = stringArg(args, "message_id");
        const rawTo = arrayArg(args, "to") ?? [];
        if (!messageId || rawTo.length === 0) {
          return errorResult("invalid_argument", "forward requires message_id and to[]");
        }
        const limited = writeRateGuard("mail", op, { message_id: messageId, recipients: rawTo.length });
        if (limited) return limited;
        const resolved = await resolveRecipients(rawTo);
        if (!resolved.ok) {
          refundWriteToken();
          return resolved.result;
        }
        const to = resolved.addresses;
        const blocked = disallowedRecipients(to);
        if (blocked.length > 0) {
          refundWriteToken();
          auditWrite({
            tool: "mail",
            op,
            target: { message_id: messageId, recipients: blocked.length },
            outcome: "recipient_not_allowed",
          });
          return errorResult(
            "recipient_not_allowed",
            `recipient(s) outside invite_allowlist: ${blocked.join(", ")}`,
          );
        }
        // Conversation + subject identify the forwarded copy afterwards.
        let source: unknown;
        try {
          source = await client.get<unknown>(
            `me/messages/${encodeURIComponent(messageId)}?$select=` + encodeURIComponent("Id,ConversationId,Subject"),
          );
        } catch (error) {
          refundWriteToken();
          throw error;
        }
        const original = normalizeMessage(camelizeKeys<Record<string, unknown>>(source ?? {}));
        const comment = typeof args.body === "string" ? args.body : undefined;
        const startedAt = Date.now() - 2000;
        await auditedWrite("mail", op, { message_id: messageId, recipients: to.length }, () =>
          forwardMessage(client, messageId, { to, ...(comment ? { comment } : {}) }),
        );
        return textResult({
          ok: true,
          op: "forward",
          recipient_count: to.length,
          ...recipientsField(to),
          ...(await sentMessageFields(client, {
            conversationId: original.conversationId,
            expectedSubject: `FW: ${original.subject}`,
            since: startedAt,
          })),
        });
      }

      if (op === "mark") {
        const messageId = stringArg(args, "message_id");
        const mark = args.mark;
        if (!messageId) return errorResult("invalid_argument", "mark requires message_id");
        if (mark !== "read" && mark !== "unread" && mark !== "flag" && mark !== "unflag") {
          return errorResult("invalid_argument", "mark requires mark=read|unread|flag|unflag");
        }
        const limited = writeRateGuard("mail", op, { message_id: messageId, mark });
        if (limited) return limited;
        const patch = buildMarkPatch(mark);
        await auditedWrite("mail", op, { message_id: messageId, mark }, () =>
          patchMessage(client, messageId, patch),
        );
        return textResult({ ok: true, op: "mark", id: messageId, mark });
      }

      if (op === "move") {
        const messageId = stringArg(args, "message_id");
        const destination = stringArg(args, "destination");
        if (!messageId || !destination) {
          return errorResult("invalid_argument", "move requires message_id and destination");
        }
        const limited = writeRateGuard("mail", op, { message_id: messageId, destination });
        if (limited) return limited;
        const moved = await auditedWrite("mail", op, { message_id: messageId, destination }, () =>
          moveMessage(client, messageId, destination),
        );
        return textResult({ ok: true, op: "move", id: moved.id, from_id: messageId, destination });
      }

      return unknownOp("mail", op, OPS);
    } catch (error) {
      const mapped = toolError(error);
      if (mapped) return mapped;
      throw error;
    }
  },
};
