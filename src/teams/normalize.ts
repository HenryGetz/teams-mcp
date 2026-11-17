/**
 * Compact rows for the `teams` tool.
 *
 * Graph returns large envelopes (etag, policyViolation, mentions, reactions,
 * viewpoint, …). Tools emit only what an agent needs, HTML is converted to text,
 * and everything null disappears (see `src/format.ts`).
 */
import { cutText, htmlToText, previewFromBody, truncate } from "../format.js";

export interface TeamRow {
  id: string;
  name: string;
  description?: string;
  web_url?: string;
}

export interface ChannelRow {
  id: string;
  name: string;
  membership?: string;
  description?: string;
  web_url?: string;
}

export interface ChatRow {
  id: string;
  name?: string;
  chat_type?: string;
  /** Graph `lastUpdatedDateTime` (not necessarily a message). */
  last_updated?: string;
  /** Text preview of the newest message, or null when Graph exposes none. */
  last_message_preview: string | null;
  /** Timestamp of that preview, or null. */
  last_message_at: string | null;
  members?: string[];
}

export interface MessageRow {
  id: string;
  from?: string;
  created?: string;
  text?: string;
  web_url?: string;
  reply_to?: string;
  message_type?: string;
  has_attachments?: boolean;
  /** A membership/roster event rather than a human message. */
  system?: true;
  /** Short event label, e.g. "member added". */
  event?: string;
}

export const MESSAGE_TEXT_CHARS = 1000;

function text(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() !== "" ? value : undefined;
}

function nested(source: Record<string, unknown>, ...path: string[]): unknown {
  let current: unknown = source;
  for (const key of path) {
    if (current === null || typeof current !== "object") return undefined;
    current = (current as Record<string, unknown>)[key];
  }
  return current;
}

export function teamRow(raw: Record<string, unknown>): TeamRow {
  return {
    id: String(raw.id ?? ""),
    name: String(raw.displayName ?? raw.name ?? ""),
    ...(text(raw.description) ? { description: truncate(text(raw.description), 200) } : {}),
    ...(text(raw.webUrl) ? { web_url: raw.webUrl as string } : {}),
  };
}

export function channelRow(raw: Record<string, unknown>): ChannelRow {
  return {
    id: String(raw.id ?? ""),
    name: String(raw.displayName ?? raw.name ?? ""),
    ...(text(raw.membershipType) ? { membership: raw.membershipType as string } : {}),
    ...(text(raw.description) ? { description: truncate(text(raw.description), 200) } : {}),
    ...(text(raw.webUrl) ? { web_url: raw.webUrl as string } : {}),
  };
}

/**
 * `lastMessagePreview` is a `chatMessageInfo`; it is optional and often absent
 * (this tenant returns no preview for some chats), so both preview fields stay
 * present as null — an absent preview must be distinguishable from "not asked".
 */
function chatPreview(raw: Record<string, unknown>): { text: string | null; at: string | null } {
  const preview = raw.lastMessagePreview;
  if (preview === null || typeof preview !== "object") return { text: null, at: null };
  const record = preview as Record<string, unknown>;
  const content = nested(record, "body", "content");
  const rendered = typeof content === "string" ? previewFromBody(content, 200) : "";
  return {
    text: rendered === "" ? null : rendered,
    at: text(record.createdDateTime) ?? null,
  };
}

export function chatRow(raw: Record<string, unknown>): ChatRow {
  const members = Array.isArray(raw.members)
    ? raw.members
        .map((member) => (member !== null && typeof member === "object"
          ? text((member as Record<string, unknown>).displayName)
          : undefined))
        .filter((name): name is string => Boolean(name))
    : [];
  const preview = chatPreview(raw);
  return {
    id: String(raw.id ?? ""),
    ...(text(raw.topic) ? { name: raw.topic as string } : {}),
    ...(text(raw.chatType) ? { chat_type: raw.chatType as string } : {}),
    ...(text(raw.lastUpdatedDateTime) ? { last_updated: raw.lastUpdatedDateTime as string } : {}),
    last_message_preview: preview.text,
    last_message_at: preview.at,
    ...(members.length > 0 ? { members } : {}),
  };
}

/**
 * True for the roster/system rows Teams writes instead of a message body.
 *
 * `messageType` is authoritative (`unknownFutureValue` is what this tenant
 * returns; `systemEventMessage` is the documented enum). The `<systemEventMessage`
 * content sniff is only a fallback for tenants that label events as plain
 * messages, and it requires a real `eventDetail` payload so a user message that
 * merely starts with that markup cannot hide itself from reads.
 */
export function isSystemMessage(raw: Record<string, unknown>): boolean {
  if (raw.messageType === "unknownFutureValue" || raw.messageType === "systemEventMessage") {
    return true;
  }
  const detail = nested(raw, "eventDetail");
  if (detail === null || typeof detail !== "object") return false;
  const content = nested(raw, "body", "content");
  return typeof content === "string" && content.trim().startsWith("<systemEventMessage");
}

const EVENT_LABELS: Record<string, string> = {
  "membersAddedEventMessageDetail": "member added",
  "membersRemovedEventMessageDetail": "member removed",
  "membersChangedEventMessageDetail": "member role changed",
  "membersJoinedEventMessageDetail": "member joined",
  "membersLeftEventMessageDetail": "member left",
  "chatRenamedEventMessageDetail": "chat renamed",
  "teamRenamedEventMessageDetail": "team renamed",
  "channelAddedEventMessageDetail": "channel added",
  "channelRemovedEventMessageDetail": "channel removed",
  "channelRenamedEventMessageDetail": "channel renamed",
  "channelDescriptionUpdatedEventMessageDetail": "channel description updated",
  "channelSetAsFavoriteByDefaultEventMessageDetail": "channel favourited",
  "channelUnsetAsFavoriteByDefaultEventMessageDetail": "channel unfavourited",
  "channelSharchangedEventMessageDetail": "channel membership changed",
  "primaryChannelSetEventMessageDetail": "primary channel set",
  "teamArchivedEventMessageDetail": "team archived",
  "teamUnarchivedEventMessageDetail": "team unarchived",
};

/** Human label for an `eventDetail` payload. */
export function systemEventLabel(eventDetail: Record<string, unknown> | undefined): string {
  const type = typeof eventDetail?.["@odata.type"] === "string" ? eventDetail["@odata.type"] : "";
  const short = type.replace(/^#microsoft\.graph\./, "");
  const known = EVENT_LABELS[short];
  if (known) return known;
  const words = short
    .replace(/EventMessageDetail$/, "")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .toLowerCase();
  return words === "" ? "system event" : words;
}

/** Compact text for a system event: label plus the members it names. */
export function systemEventText(raw: Record<string, unknown>): string {
  const detail = nested(raw, "eventDetail");
  const record = detail !== null && typeof detail === "object" ? (detail as Record<string, unknown>) : undefined;
  const label = systemEventLabel(record);
  const members = Array.isArray(record?.members)
    ? record.members
        .map((member) => {
          if (member === null || typeof member !== "object") return undefined;
          const row = member as Record<string, unknown>;
          return text(row.displayName) ?? text(row.id);
        })
        .filter((name): name is string => Boolean(name))
    : [];
  return members.length > 0 ? `${label}: ${members.join(", ")}` : label;
}

/**
 * One message. Channel messages carry `channelIdentity`, chat messages do not;
 * `replyToId` is present on replies. System events keep `message_type` but no
 * text, so a caller can tell them apart without reading markup.
 */
export function messageRow(raw: Record<string, unknown>, maxChars = MESSAGE_TEXT_CHARS): MessageRow {
  const contentType = nested(raw, "body", "contentType");
  const content = nested(raw, "body", "content");
  const system = isSystemMessage(raw);
  const rendered = system
    ? cutText(systemEventText(raw), maxChars)
    : typeof content === "string" && content !== ""
      ? cutText(contentType === "html" ? htmlToText(content) : content, maxChars)
      : undefined;
  const from =
    text(nested(raw, "from", "user", "displayName")) ??
    text(nested(raw, "from", "application", "displayName")) ??
    text(nested(raw, "from", "device", "displayName"));
  const attachments = Array.isArray(raw.attachments) ? raw.attachments.length : 0;
  return {
    id: String(raw.id ?? ""),
    ...(from ? { from } : {}),
    ...(text(raw.createdDateTime) ? { created: raw.createdDateTime as string } : {}),
    ...(rendered ? { text: rendered } : {}),
    ...(text(raw.webUrl) ? { web_url: raw.webUrl as string } : {}),
    ...(text(raw.replyToId) ? { reply_to: raw.replyToId as string } : {}),
    ...(text(raw.messageType) && raw.messageType !== "message"
      ? { message_type: raw.messageType as string }
      : {}),
    ...(attachments > 0 ? { has_attachments: true } : {}),
    ...(system
      ? { system: true as const, event: systemEventLabel(
          (nested(raw, "eventDetail") as Record<string, unknown> | undefined) ?? undefined,
        ) }
      : {}),
  };
}

/** Row mapper shared by the collection readers. */
export function messageRows(value: unknown, maxChars = MESSAGE_TEXT_CHARS): MessageRow[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((row): row is Record<string, unknown> => row !== null && typeof row === "object" && !Array.isArray(row))
    .map((row) => messageRow(row, maxChars));
}
