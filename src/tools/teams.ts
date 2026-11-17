/**
 * `teams` tool: Microsoft Teams chats, channels and messages over Graph v1.0.
 *
 * Two credentials, one API:
 *
 *  - chats + team/channel metadata + channel create/send/delete use the
 *    OWA-derived `graph` token (`src/graph/client.ts`);
 *  - channel messages/replies use the Teams-web token
 *    (`src/auth/teams-web.ts`), the only token in this tenant with
 *    `ChannelMessage.Read.All`.
 *
 * Writes are explicit and gated: `confirm:true` is mandatory, sending is limited
 * to targets this server created or an operator listed in
 * `teams.allowed_targets`, deleting is limited to channels recorded in
 * `teams-state.json`, read-only mode
 * blocks every write before any request, and every write — including policy
 * refusals — goes through the shared rate limiter and audit journal.
 */
import { z } from "zod";
import { config } from "../config.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { decodeJwtClaims } from "../auth/refresh.js";
import { getAccessToken, readTokenCache } from "../auth/store.js";
import { teamsWebStatus, type TeamsWebStatus } from "../auth/teams-web.js";
import { graphClient, type GraphClient } from "../graph/client.js";
import { nextLinkOf } from "../outlook/client.js";
import { ApiError } from "../http/client.js";
import {
  asTeamsError,
  requiredScopeOf,
  teamsWebClient,
  type TeamsWebClient,
} from "../teams/client.js";
import {
  channelRow,
  chatRow,
  isSystemMessage,
  MESSAGE_TEXT_CHARS,
  messageRow,
  teamRow,
  type ChatRow,
  type MessageRow,
} from "../teams/normalize.js";
import { isCreatedTarget, recordCreatedTarget, removeCreatedTarget } from "../teams/state.js";
import {
  arrayArg,
  booleanArg,
  errorResult,
  opOf,
  pageSize,
  stringArg,
  textResult,
  toolError,
  unknownOp,
  type ToolModule,
} from "./shared.js";
import {
  auditWrite,
  auditedWrite,
  errorOutcome,
  readOnlyGuard,
  refundWriteToken,
  writeRateGuard,
} from "../safety.js";

const OPS = ["status", "list", "read", "replies", "send", "create", "delete"] as const;
const KINDS = ["teams", "chats", "channels", "channel"] as const;
const CHAT_TYPES = ["group", "oneOnOne", "meeting"] as const;
const CAP = 25;
/** `read` pays for every row it returns: ten messages answer most questions. */
const READ_TOP = 10;
/** Per-message text cap for `read`; `text_chars` may raise it to the hard cap. */
const READ_TEXT_CHARS = 500;
/** Provider page size when filtering; output caps stay `top`. */
const FETCH_TOP = 50;
const SCAN_PAGES = 15;
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Declared in the tool schema as `members.max(20)`; handlers must enforce it too. */
const MAX_MEMBERS = 20;
/**
 * Graph rejects oversized message bodies and the limit is not available as a
 * constant; 20 000 characters is below the documented chat-message ceiling and
 * bounds the request before any network call or rate-token spend.
 */
const MAX_TEXT_CHARS = 20_000;

type Record_ = Record<string, unknown>;

function asRecord(value: unknown): Record_ | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record_)
    : undefined;
}

/** `value[]` + `@odata.nextLink` from a Graph collection, without the envelope. */
function collection(payload: unknown): { rows: Record_[]; next?: string } {
  const value = asRecord(payload)?.value;
  const rows = Array.isArray(value)
    ? value.map(asRecord).filter((row): row is Record_ => row !== undefined)
    : [];
  const next = nextLinkOf(payload);
  return next ? { rows, next } : { rows };
}

/* ------------------------------------------------------------------ */
/* Cursors                                                             */
/* ------------------------------------------------------------------ */

/**
 * A page cursor plus the number of rows already consumed from it. Filtering
 * (system events, chat queries) can stop mid-page; a plain nextLink would then
 * skip the rest of that page.
 *
 * `context` binds the cursor to the call that minted it — op, surface, ids and
 * filters. Without it a cursor from an unfiltered list could be replayed with a
 * filter (silently skipping the rows the old offset hid) or a chat-read cursor
 * could be used on `replies`.
 */
interface TeamsCursor {
  link: string;
  skip: number;
  context: string;
}

const CURSOR_VERSION = 2;

/**
 * Stable fingerprint of the call shape. `null` stands for "absent"; callers
 * normalise case-insensitive inputs (queries) before passing them in.
 */
function cursorFingerprint(parts: ReadonlyArray<string | boolean | undefined>): string {
  return JSON.stringify(parts.map((part) => (part === undefined ? null : part)));
}

/**
 * Mint a cursor from an absolute URL. Mint sites hold relative paths on the
 * first page (`me/chats?…`, `teams/{id}/channels`); the decoder only accepts
 * `http(s)://`, so storing the relative form produced an unusable `next_cursor`
 * as soon as a page was cut. `client.url()` is the identity for provider
 * `@odata.nextLink` values.
 */
function mintCursor(
  client: { url(pathOrUrl: string): string },
  link: string,
  skip: number,
  context: string,
): string {
  return encodeTeamsCursor({ link: client.url(link), skip, context });
}

function encodeTeamsCursor(cursor: TeamsCursor): string {
  return Buffer.from(
    JSON.stringify({ v: CURSOR_VERSION, u: cursor.link, s: cursor.skip, c: cursor.context }),
    "utf8",
  ).toString("base64url");
}

/** Versioned, context-bound cursors only; a legacy/unbound value is rejected. */
function decodeTeamsCursor(value: string): TeamsCursor | undefined {
  let decoded: string;
  try {
    decoded = Buffer.from(value, "base64url").toString("utf8");
  } catch {
    return undefined;
  }
  try {
    const record = asRecord(JSON.parse(decoded) as unknown);
    if (record?.v !== CURSOR_VERSION) return undefined;
    const link = typeof record.u === "string" ? record.u : "";
    if (!/^https?:\/\//i.test(link)) return undefined;
    if (typeof record.c !== "string" || record.c === "") return undefined;
    const skip = typeof record.s === "number" && Number.isInteger(record.s) && record.s >= 0 ? record.s : 0;
    return { link, skip, context: record.c };
  } catch {
    return undefined;
  }
}

function cursorOf(
  client: GraphClient | TeamsWebClient,
  args: Record_,
  context: string,
): { cursor?: TeamsCursor; result?: CallToolResult } {
  const supplied = stringArg(args, "cursor");
  if (!supplied) return {};
  const cursor = decodeTeamsCursor(supplied);
  if (!cursor) {
    return {
      result: errorResult(
        "invalid_cursor",
        "cursor is not a teams-mcp cursor",
        "omit it to start from the first page, or pass a next_cursor from a prior teams call",
      ),
    };
  }
  if (cursor.context !== context) {
    return {
      result: errorResult(
        "invalid_cursor",
        "cursor was minted for a different call (op, target or filters changed)",
        "cursors only continue the exact call that produced them; repeat the first page with the new arguments",
      ),
    };
  }
  if (!client.ownsUrl(cursor.link)) {
    return {
      result: errorResult(
        "invalid_cursor",
        "cursor does not belong to this API",
        "use the next_cursor returned by a prior teams call",
      ),
    };
  }
  return { cursor };
}

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */

function requireTeam(args: Record_): string | undefined {
  return stringArg(args, "team_id");
}

function missing(field: string, op: string): CallToolResult {
  return errorResult(
    "invalid_argument",
    `${op} requires ${field}`,
    op === "read" ? "pass chat_id, or team_id together with channel_id" : "ids come from teams op=list",
  );
}

function normalizedName(value: unknown): string {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

/** Resolve a member list to AAD object ids; names are rejected as ambiguous. */
async function memberIds(
  client: GraphClient,
  members: string[],
): Promise<{ ids: string[]; result?: CallToolResult }> {
  const ids: string[] = [];
  for (const raw of members) {
    const value = raw.trim();
    if (value === "") continue;
    if (UUID.test(value)) {
      ids.push(value);
      continue;
    }
    try {
      const user = await client.get<Record_>(`users/${encodeURIComponent(value)}?$select=id`);
      const id = typeof user.id === "string" ? user.id : "";
      if (id === "") return { ids: [], result: errorResult("member_not_found", `no directory id for "${value}"`) };
      ids.push(id);
    } catch (error) {
      const mapped = toolError(asTeamsError(error));
      if (mapped) return { ids: [], result: mapped };
      throw error;
    }
  }
  return { ids: [...new Set(ids)] };
}

/** The signed-in user's AAD id: cache account, else the graph token's `oid`. */
async function selfUserId(): Promise<string | undefined> {
  const cached = await readTokenCache();
  if (cached.cache?.account?.id) return cached.cache.account.id;
  try {
    const token = await getAccessToken({ audience: "graph" });
    return decodeJwtClaims(token.token)?.oid;
  } catch {
    return undefined;
  }
}

function chatPayload(text: string): Record_ {
  return { body: { contentType: "text", content: text } };
}

/* ------------------------------------------------------------------ */
/* Message paging                                                      */
/* ------------------------------------------------------------------ */

interface MessagePage {
  messages: MessageRow[];
  next_cursor?: string;
  /** A cursor exists, so more rows can be fetched. */
  has_more: boolean;
  /** The requested answer was cut short (page budget), not merely paged. */
  truncated: boolean;
}

/**
 * Read one message collection, filtering system events before applying `top`.
 * System rows are skipped without consuming the cap, and paging continues until
 * enough real messages are collected or the provider runs out of pages. The
 * cursor records the exact row offset, so nothing is lost when a page is cut.
 */
async function collectMessages(
  client: GraphClient | TeamsWebClient,
  baseUrl: string,
  cursor: TeamsCursor | undefined,
  top: number,
  includeSystem: boolean,
  context: string,
  maxChars = MESSAGE_TEXT_CHARS,
): Promise<MessagePage> {
  const messages: MessageRow[] = [];
  let url = cursor?.link ?? baseUrl;
  let skip = cursor?.skip ?? 0;
  let pages = 0;
  for (;;) {
    const payload = await client.get<unknown>(url);
    const page = collection(payload);
    for (let index = skip; index < page.rows.length; index += 1) {
      const row = page.rows[index];
      if (!row) continue;
      if (!includeSystem && isSystemMessage(row)) continue;
      messages.push(messageRow(row, maxChars));
      if (messages.length >= top) {
        const remaining = index + 1;
        const nextCursor = remaining < page.rows.length
          ? mintCursor(client, url, remaining, context)
          : page.next
            ? mintCursor(client, page.next, 0, context)
            : undefined;
        return {
          messages,
          ...(nextCursor ? { next_cursor: nextCursor } : {}),
          has_more: Boolean(nextCursor),
          truncated: false,
        };
      }
    }
    pages += 1;
    skip = 0;
    if (!page.next) {
      return { messages, has_more: false, truncated: false };
    }
    if (pages >= 10) {
      return {
        messages,
        next_cursor: mintCursor(client, page.next, 0, context),
        has_more: true,
        truncated: messages.length < top,
      };
    }
    url = page.next;
  }
}

/* ------------------------------------------------------------------ */
/* Tool                                                                */
/* ------------------------------------------------------------------ */

export const teamsTool: ToolModule = {
  name: "teams",
  title: "Microsoft Teams",
  description:
    "Read/write Teams chats, channels and messages (Graph). op=status (offline) | list kind=teams|chats|channels (chats: chat_type, query; teams/channels: query) | read (chat_id, or team_id+channel_id; text_chars<=1000, default 500; system events hidden unless include_system; paging) | replies (message_id) | send (confirm; text<=20000; target must be allowed) | create (private channel only) | delete (server-created channels; confirm + expected_name). Cursors are bound to the call that minted them; has_more/truncated report a cut answer.",
  shape: {
    op: z.enum(OPS),
    kind: z.enum(KINDS).optional().describe("list: teams|chats|channels; create: channel"),
    team_id: z.string().optional().describe("team id (channels, channel read/replies/send/create/delete)"),
    channel_id: z.string().optional().describe("channel id (read/replies/send/delete)"),
    chat_id: z.string().optional().describe("chat id (read/send)"),
    message_id: z.string().optional().describe("replies: parent channel message id"),
    chat_type: z.enum(CHAT_TYPES).optional().describe("list kind=chats: filter by chat type"),
    query: z.string().optional().describe("list: match chat topic/member, or team/channel name"),
    include_system: z.boolean().optional().describe("read: include roster/system events (default false)"),
    top: z.number().int().min(1).max(CAP).optional().describe(`page size, max ${CAP}; read default ${READ_TOP}`),
    text_chars: z
      .number()
      .int()
      .min(1)
      .max(MESSAGE_TEXT_CHARS)
      .optional()
      .describe("read: per-message text cap, default 500"),
    cursor: z.string().optional().describe("opaque next_cursor from a prior call"),
    text: z.string().optional().describe("send: message body (plain text, max 20000 chars)"),
    name: z.string().optional().describe("create channel: display name"),
    expected_name: z.string().optional().describe("delete: exact channel name guard"),
    members: z
      .array(z.string().min(1))
      .max(20)
      .optional()
      .describe("create: extra members (UPN/id, max 20); self is owner"),
    private: z.boolean().optional().describe("create channel: must be true (only private channels)"),
    confirm: z.boolean().optional().describe("write ops: must be true; without it nothing is written"),
  },

  async handle(args): Promise<CallToolResult> {
    const op = opOf(args, OPS);
    if (!op) return unknownOp("teams", typeof args.op === "string" ? args.op : "", OPS);

    const blocked = readOnlyGuard("teams", op, args);
    if (blocked) return blocked;

    try {
      if (op === "status") {
        const status: TeamsWebStatus = await teamsWebStatus();
        return textResult({
          teams_web: status,
          chat_token: "graph (OWA)",
          channel_history_token: status.fresh ? "teams-web" : "teams-web (stale: a read may harvest)",
        });
      }

      if (op === "list") return await listOp(args);
      if (op === "read") return await readOp(args);
      if (op === "replies") return await repliesOp(args);
      if (op === "send") return await sendOp(args);
      if (op === "delete") return await deleteOp(args);
      return await createOp(args);
    } catch (error) {
      const mapped = toolError(asTeamsError(error));
      if (mapped) return mapped;
      throw error;
    }
  },
};

/* ------------------------------------------------------------------ */
/* list                                                                */
/* ------------------------------------------------------------------ */

function queryMatches(row: Record_, query: string): boolean {
  const haystack = [row.topic];
  if (Array.isArray(row.members)) {
    for (const member of row.members) {
      if (member === null || typeof member !== "object") continue;
      const record = member as Record_;
      haystack.push(record.displayName, record.email);
    }
  }
  return haystack.some((value) => typeof value === "string" && value.toLowerCase().includes(query));
}

/** Teams and channels carry their name as `displayName`; accept `name` too. */
function nameMatches(row: Record_, query: string): boolean {
  return [row.displayName, row.name].some(
    (value) => typeof value === "string" && value.toLowerCase().includes(query),
  );
}

interface ScannedRows<Row> {
  rows: Row[];
  next_cursor?: string;
  pages: number;
}

/**
 * Scan a collection that rejects `$top`, keeping matching rows until `top` are
 * collected, the provider is exhausted, or the page budget is spent. The
 * returned cursor holds the exact page URL + row offset of the first unseen
 * row, so a filter never skips rows an earlier call hid.
 */
async function scanRows<Row>(
  client: GraphClient,
  pageUrl: string,
  cursor: TeamsCursor | undefined,
  top: number,
  keep: (row: Record_) => boolean,
  map: (row: Record_) => Row,
  context: string,
): Promise<ScannedRows<Row>> {
  const rows: Row[] = [];
  let url = cursor?.link ?? pageUrl;
  let skip = cursor?.skip ?? 0;
  let pages = 0;
  for (;;) {
    const payload = await client.get<unknown>(url);
    const page = collection(payload);
    for (let index = skip; index < page.rows.length; index += 1) {
      const row = page.rows[index];
      if (!row || !keep(row)) continue;
      rows.push(map(row));
      if (rows.length >= top) {
        const remaining = index + 1;
        const nextCursor = remaining < page.rows.length
          ? mintCursor(client, url, remaining, context)
          : page.next
            ? mintCursor(client, page.next, 0, context)
            : undefined;
        return { rows, ...(nextCursor ? { next_cursor: nextCursor } : {}), pages: pages + 1 };
      }
    }
    pages += 1;
    skip = 0;
    if (!page.next) return { rows, pages };
    if (pages >= SCAN_PAGES) {
      return { rows, next_cursor: mintCursor(client, page.next, 0, context), pages };
    }
    url = page.next;
  }
}

async function listChats(args: Record_, top: number): Promise<CallToolResult> {
  const client = graphClient();
  const chatType = stringArg(args, "chat_type");
  if (chatType && !(CHAT_TYPES as readonly string[]).includes(chatType)) {
    return errorResult("invalid_argument", `chat_type must be one of ${CHAT_TYPES.join(", ")}`);
  }
  const query = stringArg(args, "query")?.toLowerCase();
  const select = "id,topic,chatType,lastUpdatedDateTime,lastMessagePreview";
  const filter = chatType ? `&$filter=chatType eq '${chatType}'` : "";
  const context = cursorFingerprint(["list", "chats", chatType, query]);
  const cursor = cursorOf(client, args, context);
  if (cursor.result) return cursor.result;

  // Membership churn mid-scan can return the same chat twice (eval observed
  // count:3 for 2 chats). Keep one row per id, preferring the newest snapshot.
  const seen = new Map<string, Record_>();
  const absorb = (row: Record_): void => {
    const id = String(row.id ?? "");
    if (id === "") return;
    const previous = seen.get(id);
    if (!previous) {
      seen.set(id, row);
      return;
    }
    const before = Date.parse(String(previous.lastUpdatedDateTime ?? ""));
    const after = Date.parse(String(row.lastUpdatedDateTime ?? ""));
    if (Number.isNaN(before) || (!Number.isNaN(after) && after > before)) seen.set(id, row);
  };
  const chats = (): ChatRow[] => [...seen.values()].map(chatRow);

  if (!chatType && !query) {
    const pageUrl = cursor.cursor?.link ?? `me/chats?$top=${top}&$select=${select}`;
    const payload = await client.get<unknown>(pageUrl);
    const page = collection(payload);
    let consumed = cursor.cursor?.skip ?? 0;
    for (; consumed < page.rows.length && seen.size < top; consumed += 1) {
      const row = page.rows[consumed];
      if (row) absorb(row);
    }
    const rows = chats();
    const nextCursor =
      consumed < page.rows.length
        ? mintCursor(client, pageUrl, consumed, context)
        : page.next
          ? mintCursor(client, page.next, 0, context)
          : undefined;
    const hasMore = Boolean(nextCursor);
    return textResult({
      kind: "chats",
      count: rows.length,
      chats: rows,
      has_more: hasMore,
      truncated: rows.length < top && hasMore,
      ...(nextCursor ? { next_cursor: nextCursor } : {}),
    });
  }

  // A filter is a promise of a complete answer, so scan pages until `top`
  // matches are collected or the provider runs out. `$expand=members` is only
  // needed when the query may match a member name.
  let url: string =
    cursor.cursor?.link ??
    `me/chats?$top=${FETCH_TOP}&$select=${select}${query ? "&$expand=members" : ""}${filter}`;
  let skip = cursor.cursor?.skip ?? 0;
  let pages = 0;
  for (;;) {
    const payload = await client.get<unknown>(url);
    const page = collection(payload);
    for (let index = skip; index < page.rows.length; index += 1) {
      const row = page.rows[index];
      if (!row) continue;
      // `$filter` is accepted by Graph but its semantics are unverified; never
      // return another chat type when chat_type was asked for.
      if (chatType && row.chatType !== chatType) continue;
      if (query && !queryMatches(row, query)) continue;
      absorb(row);
      if (seen.size >= top) {
        const rows = chats();
        const remaining = index + 1;
        const nextCursor = remaining < page.rows.length
          ? mintCursor(client, url, remaining, context)
          : page.next
            ? mintCursor(client, page.next, 0, context)
            : undefined;
        return textResult({
          kind: "chats",
          count: rows.length,
          chats: rows,
          scanned_pages: pages + 1,
          has_more: Boolean(nextCursor),
          truncated: false,
          ...(nextCursor ? { next_cursor: nextCursor } : {}),
        });
      }
    }
    pages += 1;
    skip = 0;
    if (!page.next || pages >= SCAN_PAGES) {
      const rows = chats();
      const nextCursor = page.next ? mintCursor(client, page.next, 0, context) : undefined;
      const hasMore = Boolean(nextCursor);
      return textResult({
        kind: "chats",
        count: rows.length,
        chats: rows,
        scanned_pages: pages,
        has_more: hasMore,
        truncated: rows.length < top && hasMore,
        ...(nextCursor ? { next_cursor: nextCursor } : {}),
      });
    }
    url = page.next;
  }
}

async function listOp(args: Record_): Promise<CallToolResult> {
  const kind = stringArg(args, "kind");
  if (kind !== "teams" && kind !== "chats" && kind !== "channels") {
    return errorResult("invalid_argument", "list requires kind=teams|chats|channels", "e.g. {op:'list', kind:'chats'}");
  }
  const top = pageSize(args, "top", CAP, CAP);

  if (kind === "chats") return listChats(args, top);

  const client = graphClient();
  const query = stringArg(args, "query")?.toLowerCase();

  if (kind === "channels") {
    const teamId = requireTeam(args);
    if (!teamId) return missing("team_id", "list kind=channels");
    // Graph rejects $top on this collection; cap client-side.
    const context = cursorFingerprint(["list", "channels", teamId, query]);
    const cursor = cursorOf(client, args, context);
    if (cursor.result) return cursor.result;
    const scan = await scanRows(
      client,
      `teams/${encodeURIComponent(teamId)}/channels`,
      cursor.cursor,
      top,
      (row) => !query || nameMatches(row, query),
      channelRow,
      context,
    );
    const hasMore = Boolean(scan.next_cursor);
    return textResult({
      kind,
      team_id: teamId,
      count: scan.rows.length,
      channels: scan.rows,
      has_more: hasMore,
      truncated: scan.rows.length < top && hasMore,
      ...(query ? { scanned_pages: scan.pages } : {}),
      ...(scan.next_cursor ? { next_cursor: scan.next_cursor } : {}),
    });
  }

  // joinedTeams also rejects $top.
  const context = cursorFingerprint(["list", "teams", query]);
  const cursor = cursorOf(client, args, context);
  if (cursor.result) return cursor.result;
  const scan = await scanRows(
    client,
    "me/joinedTeams",
    cursor.cursor,
    top,
    (row) => !query || nameMatches(row, query),
    teamRow,
    context,
  );
  const hasMore = Boolean(scan.next_cursor);
  return textResult({
    kind,
    count: scan.rows.length,
    teams: scan.rows,
    has_more: hasMore,
    truncated: scan.rows.length < top && hasMore,
    ...(query ? { scanned_pages: scan.pages } : {}),
    ...(scan.next_cursor ? { next_cursor: scan.next_cursor } : {}),
  });
}

/* ------------------------------------------------------------------ */
/* read / replies                                                      */
/* ------------------------------------------------------------------ */

async function readOp(args: Record_): Promise<CallToolResult> {
  const top = pageSize(args, "top", READ_TOP, CAP);
  const maxChars = pageSize(args, "text_chars", READ_TEXT_CHARS, MESSAGE_TEXT_CHARS);
  const includeSystem = booleanArg(args, "include_system") === true;
  const chatId = stringArg(args, "chat_id");
  const teamId = requireTeam(args);
  const channelId = stringArg(args, "channel_id");

  if (chatId) {
    const client = graphClient();
    const context = cursorFingerprint(["read", "chat", chatId, includeSystem]);
    const cursor = cursorOf(client, args, context);
    if (cursor.result) return cursor.result;
    const page = await collectMessages(
      client,
      `chats/${encodeURIComponent(chatId)}/messages?$top=${FETCH_TOP}`,
      cursor.cursor,
      top,
      includeSystem,
      context,
      maxChars,
    );
    return textResult({
      surface: "chat",
      chat_id: chatId,
      count: page.messages.length,
      messages: page.messages,
      has_more: page.has_more,
      truncated: page.truncated,
      ...(page.next_cursor ? { next_cursor: page.next_cursor } : {}),
    });
  }

  if (!teamId || !channelId) return missing("chat_id, or team_id and channel_id", "read");
  const client = teamsWebClient();
  const context = cursorFingerprint(["read", "channel", teamId, channelId, includeSystem]);
  const cursor = cursorOf(client, args, context);
  if (cursor.result) return cursor.result;
  const page = await collectMessages(
    client,
    `teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages?$top=${FETCH_TOP}`,
    cursor.cursor,
    top,
    includeSystem,
    context,
    maxChars,
  );
  return textResult({
    surface: "channel",
    team_id: teamId,
    channel_id: channelId,
    count: page.messages.length,
    messages: page.messages,
    has_more: page.has_more,
    truncated: page.truncated,
    ...(page.next_cursor ? { next_cursor: page.next_cursor } : {}),
  });
}

async function repliesOp(args: Record_): Promise<CallToolResult> {
  const teamId = requireTeam(args);
  const channelId = stringArg(args, "channel_id");
  const messageId = stringArg(args, "message_id");
  if (!teamId || !channelId) return missing("team_id and channel_id", "replies");
  if (!messageId) return missing("message_id", "replies");

  const top = pageSize(args, "top", CAP, CAP);
  const client = teamsWebClient();
  const context = cursorFingerprint(["replies", "channel", teamId, channelId, messageId]);
  const cursor = cursorOf(client, args, context);
  if (cursor.result) return cursor.result;
  const page = await collectMessages(
    client,
    `teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies?$top=${FETCH_TOP}`,
    cursor.cursor,
    top,
    true,
    context,
  );
  return textResult({
    surface: "channel_replies",
    team_id: teamId,
    channel_id: channelId,
    message_id: messageId,
    count: page.messages.length,
    messages: page.messages,
    has_more: page.has_more,
    truncated: page.truncated,
    ...(page.next_cursor ? { next_cursor: page.next_cursor } : {}),
  });
}

/* ------------------------------------------------------------------ */
/* send                                                                */
/* ------------------------------------------------------------------ */

interface SendOutcome {
  ok: true;
  surface: "chat" | "channel";
  target_id: string;
  message_id?: string;
  team_id?: string;
}

async function sendOp(args: Record_): Promise<CallToolResult> {
  const text = stringArg(args, "text");
  if (!text) return errorResult("invalid_argument", "send requires text", "plain text, no attachments or mentions");
  const chatId = stringArg(args, "chat_id");
  const teamId = requireTeam(args);
  const channelId = stringArg(args, "channel_id");
  const surface = chatId ? "chat" : teamId && channelId ? "channel" : undefined;
  if (!surface) return missing("chat_id, or team_id and channel_id", "send");

  const targetId = chatId ?? (channelId as string);
  const target = {
    surface,
    target_id: targetId,
    ...(teamId ? { team_id: teamId } : {}),
    chars: text.length,
  };

  // Shape refusal before the confirm gate, the rate bucket or any request.
  if (text.length > MAX_TEXT_CHARS) {
    auditWrite({ tool: "teams", op: "send", target, outcome: "invalid_argument" });
    return errorResult(
      "invalid_argument",
      `send text is ${text.length} characters; the limit is ${MAX_TEXT_CHARS}`,
      "split long content into a file or a shorter message; nothing was sent",
    );
  }

  if (booleanArg(args, "confirm") !== true) {
    auditWrite({ tool: "teams", op: "send", target, outcome: "confirmation_required" });
    return errorResult(
      "confirmation_required",
      `teams.send would post to ${surface} ${targetId}; nothing was sent`,
      "re-call with confirm:true once the target is right",
    );
  }

  const created = await isCreatedTarget({ surface, id: targetId, ...(teamId ? { teamId } : {}) });
  if (!created) {
    // The model cannot widen its own write scope: only an operator allowlist
    // entry (`chat:<id>` / `channel:<team>/<channel>`) opens a pre-existing target.
    const allowKey = surface === "chat" ? `chat:${targetId}` : `channel:${teamId}/${targetId}`;
    if (!config().teams.allowedTargets.includes(allowKey)) {
      auditWrite({ tool: "teams", op: "send", target, outcome: "target_not_allowed" });
      return errorResult(
        "target_not_allowed",
        `${surface} ${targetId} is not owned by this server and is not in the operator allowlist; nothing was sent`,
        "an operator must add this exact target to teams.allowed_targets in config.yaml (chat:<chat_id> or channel:<team_id>/<channel_id>); the model cannot grant it",
      );
    }
  }

  const limited = writeRateGuard("teams", "send", target);
  if (limited) return limited;

  const outcome = await auditedWrite("teams", "send", target, async (): Promise<SendOutcome> => {
    if (surface === "chat") {
      const response = await graphClient().post<Record_>(
        `chats/${encodeURIComponent(targetId)}/messages`,
        chatPayload(text),
      );
      return {
        ok: true,
        surface,
        target_id: targetId,
        ...(typeof response?.id === "string" ? { message_id: response.id } : {}),
      };
    }
    const response = await teamsWebClient().post<Record_>(
      `teams/${encodeURIComponent(teamId as string)}/channels/${encodeURIComponent(targetId)}/messages`,
      chatPayload(text),
    );
    return {
      ok: true,
      surface,
      target_id: targetId,
      team_id: teamId,
      ...(typeof response?.id === "string" ? { message_id: response.id } : {}),
    };
  });

  return textResult(outcome);
}

/* ------------------------------------------------------------------ */
/* create                                                              */
/* ------------------------------------------------------------------ */

async function createOp(args: Record_): Promise<CallToolResult> {
  const kind = stringArg(args, "kind");
  if (kind === "chat") {
    // Live Graph rejects `roles: []` (and omitted roles, and
    // "member") on POST /chats: "Each chat member must have at least one role
    // specified. Supported roles are 'owner' and 'guest'." Creating a chat here
    // would therefore hand every invitee ownership, so it is refused.
    auditWrite({ tool: "teams", op: "create", target: { kind }, outcome: "invalid_argument" });
    return errorResult(
      "invalid_argument",
      "teams.create does not create chats: Graph requires every chat member to be an owner or guest, and this server will not grant chat ownership silently",
      "create the group chat in Microsoft Teams and ask the operator to add it to teams.allowed_targets; kind=channel is still supported",
    );
  }
  if (kind !== "channel") {
    return errorResult(
      "invalid_argument",
      "create requires kind=channel",
      "only private channels are created by this server; create chats in the Teams UI",
    );
  }
  const target = { kind, name: stringArg(args, "name") };
  if (booleanArg(args, "confirm") !== true) {
    auditWrite({ tool: "teams", op: "create", target, outcome: "confirmation_required" });
    return errorResult(
      "confirmation_required",
      `teams.create would create a ${kind}; nothing was created`,
      "re-call with confirm:true once the name/members are right",
    );
  }

  // Validate the whole shape before spending a write token or touching the network.
  const name = stringArg(args, "name");
  const members = arrayArg(args, "members");
  const teamId = requireTeam(args);
  const invalid = (message: string, hint?: string): CallToolResult => {
    auditWrite({ tool: "teams", op: "create", target, outcome: "invalid_argument" });
    return errorResult("invalid_argument", message, hint);
  };
  if (!teamId) return invalid("create kind=channel requires team_id", "ids come from teams op=list");
  if (!name) return invalid("create kind=channel requires name");
  if (members && members.length > MAX_MEMBERS) {
    return invalid(`create kind=channel accepts at most ${MAX_MEMBERS} members (got ${members.length})`);
  }
  if (booleanArg(args, "private") !== true) {
    auditWrite({ tool: "teams", op: "create", target, outcome: "private_required" });
    return errorResult(
      "private_required",
      "only private channels may be created by this server",
      "pass private:true; team-wide channels affect everyone",
    );
  }

  const limited = writeRateGuard("teams", "create", target);
  if (limited) return limited;

  const client = graphClient();
  const self = await selfUserId();
  if (!self) {
    refundWriteToken();
    return errorResult("auth_required", "cannot determine the signed-in user id; run auth.login");
  }

  // Private channel in an existing team only; this server never creates teams.
  const owners = await client.get<unknown>(`groups/${encodeURIComponent(teamId as string)}/owners?$select=id&$top=999`);
  const ownerIds = collection(owners).rows.map((row) => String(row.id ?? ""));
  if (!ownerIds.includes(self)) {
    refundWriteToken();
    auditWrite({ tool: "teams", op: "create", target: { ...target, team_id: teamId }, outcome: "not_owner" });
    return errorResult(
      "not_owner",
      `the signed-in user is not an owner of team ${teamId}`,
      "only team owners may create channels; pick another team or ask an owner",
    );
  }

  const resolved = members ? await memberIds(client, members) : { ids: [] as string[] };
  if (resolved.result) {
    refundWriteToken();
    return resolved.result;
  }
  const extraIds = resolved.ids.filter((id) => id !== self);
  const bind = (id: string, roles: string[]): Record_ => ({
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles,
    "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${id}')`,
  });

  const body: Record_ = {
    displayName: name,
    membershipType: "private",
    members: [bind(self, ["owner"]), ...extraIds.map((id) => bind(id, []))],
  };
  const response = await auditedWrite("teams", "create", { ...target, team_id: teamId }, () =>
    client.post<Record_>(`teams/${encodeURIComponent(teamId as string)}/channels`, body),
  );
  const id = typeof response?.id === "string" ? response.id : "";
  if (id === "") return errorResult("invalid_response", "Graph created the channel but returned no id");
  await recordCreatedTarget({
    surface: "channel",
    id,
    team_id: teamId as string,
    name,
    created_at: new Date().toISOString(),
  });
  return textResult({
    ok: true,
    kind,
    id,
    name,
    team_id: teamId,
    membership: "private",
    ...(typeof response?.webUrl === "string" ? { web_url: response.webUrl } : {}),
  });
}

/* ------------------------------------------------------------------ */
/* delete (channels this server created)                               */
/* ------------------------------------------------------------------ */

async function deleteOp(args: Record_): Promise<CallToolResult> {
  const teamId = requireTeam(args);
  const channelId = stringArg(args, "channel_id");
  const expectedName = stringArg(args, "expected_name");
  if (!teamId || !channelId) return missing("team_id and channel_id", "delete");
  if (!expectedName) {
    return errorResult(
      "invalid_argument",
      "delete requires expected_name",
      "pass the channel name you created; it is compared before anything is deleted",
    );
  }

  const target = { team_id: teamId, channel_id: channelId, expected_name: expectedName };
  if (booleanArg(args, "confirm") !== true) {
    auditWrite({ tool: "teams", op: "delete", target, outcome: "confirmation_required" });
    return errorResult(
      "confirmation_required",
      `teams.delete would remove channel ${channelId}; nothing was deleted`,
      "re-call with confirm:true once the id and expected_name are right",
    );
  }

  const created = await isCreatedTarget({ surface: "channel", id: channelId, teamId });
  if (!created) {
    auditWrite({ tool: "teams", op: "delete", target, outcome: "target_not_created" });
    return errorResult(
      "target_not_created",
      `channel ${channelId} was not created by this server; refusing to delete it`,
      "only channels recorded in teams-state.json can be deleted here; remove anything else in the Teams UI",
    );
  }

  const limited = writeRateGuard("teams", "delete", target);
  if (limited) return limited;

  const client = graphClient();
  // Read first: the name guard must compare a live object, and a 404 means the
  // ledger is stale, which is not an error.
  let current: Record_ | undefined;
  try {
    current = await client.get<Record_>(
      `teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}`,
    );
  } catch (error) {
    if (error instanceof ApiError && (error.status === 404 || error.status === 410)) {
      refundWriteToken();
      // A ledger prune is still a state mutation; journal it.
      auditWrite({ tool: "teams", op: "delete", target, outcome: "already_absent" });
      await removeCreatedTarget({ surface: "channel", id: channelId, teamId });
      return textResult({ deleted: false, already_absent: true, team_id: teamId, channel_id: channelId });
    }
    refundWriteToken();
    throw error;
  }

  const liveName = current.displayName;
  if (normalizedName(liveName) !== normalizedName(expectedName)) {
    refundWriteToken();
    auditWrite({ tool: "teams", op: "delete", target, outcome: "name_mismatch" });
    return errorResult(
      "name_mismatch",
      `channel ${channelId} is named "${String(liveName ?? "")}", not "${expectedName}"; nothing was deleted`,
      "list the channels and pass the exact display name",
    );
  }

  // Exactly one journal line per attempt (including the already-absent prune).
  try {
    await client.delete(`teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}`);
  } catch (error) {
    if (error instanceof ApiError && (error.status === 404 || error.status === 410)) {
      refundWriteToken();
      auditWrite({ tool: "teams", op: "delete", target, outcome: "already_absent" });
      await removeCreatedTarget({ surface: "channel", id: channelId, teamId });
      return textResult({ deleted: false, already_absent: true, team_id: teamId, channel_id: channelId });
    }
    auditWrite({ tool: "teams", op: "delete", target, outcome: errorOutcome(error) });
    throw error;
  }

  auditWrite({ tool: "teams", op: "delete", target, outcome: "ok" });
  await removeCreatedTarget({ surface: "channel", id: channelId, teamId });
  return textResult({
    deleted: true,
    team_id: teamId,
    channel_id: channelId,
    name: typeof liveName === "string" ? liveName : expectedName,
  });
}

/** Re-exported for tests: the scope a Graph 403 names, if it names one. */
export { requiredScopeOf };
