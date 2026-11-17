/**
 * Read-only Outlook REST v2.0 attachment access.
 *
 * Two capabilities, both GETs on the message's attachments collection:
 *  - metadata (`$select=Id,Name,ContentType,Size,IsInline`, no content), capped
 *    at 20 rows per call;
 *  - one attachment including `ContentBytes`, which the caller turns into text
 *    or a cached file.
 *
 * Nothing here mutates the mailbox, sends, or opens a downloaded file. Binary
 * payloads are written by `saveAttachment` with 0600 inside the cache directory
 * only, under a sanitised name and a random suffix.
 */
import { randomBytes } from "node:crypto";
import { chmod, mkdir, writeFile } from "node:fs/promises";
import { resolve, sep } from "node:path";
import { camelizeKeys, objectRows, type OutlookClient } from "./client.js";

/** Maximum attachment metadata rows returned per `attachments` call. */
export const ATTACHMENT_LIST_CAP = 20;

/** Hard cap on a decoded attachment we are willing to decode/save. */
export const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024;

/** Text returned from a textual attachment is cut at this many characters. */
export const ATTACHMENT_TEXT_CHARS = 20_000;

export interface AttachmentMeta {
  id: string;
  name: string;
  content_type: string;
  size: number;
  is_inline: boolean;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" && value !== "" ? value : undefined;
}

function metaOf(row: Record<string, unknown>, fallbackId = ""): AttachmentMeta {
  const size = typeof row.size === "number" && Number.isFinite(row.size) ? row.size : 0;
  return {
    id: asString(row.id) ?? fallbackId,
    name: asString(row.name) ?? "attachment",
    content_type: asString(row.contentType) ?? "application/octet-stream",
    size,
    is_inline: row.isInline === true,
  };
}

const ATTACHMENT_SELECT = "Id,Name,ContentType,Size,IsInline";

/** Metadata only — `ContentBytes` is never requested or returned here. */
export async function listAttachments(
  client: OutlookClient,
  messageId: string,
): Promise<{ attachments: AttachmentMeta[]; truncated: boolean }> {
  const params = new URLSearchParams({
    $select: ATTACHMENT_SELECT,
    $top: String(ATTACHMENT_LIST_CAP),
  });
  const payload = await client.get<unknown>(
    `me/messages/${encodeURIComponent(messageId)}/attachments?${params.toString()}`,
  );
  const rows =
    payload !== null && typeof payload === "object"
      ? objectRows((payload as Record<string, unknown>).value)
      : [];
  const attachments = rows
    .map((row) => metaOf(camelizeKeys<Record<string, unknown>>(row)))
    .slice(0, ATTACHMENT_LIST_CAP);
  return { attachments, truncated: rows.length >= ATTACHMENT_LIST_CAP };
}

export interface FetchedAttachment {
  meta: AttachmentMeta;
  /** Base64 payload, or undefined for a reference attachment. */
  contentBytes?: string;
}

/** One attachment, including `ContentBytes` when the service has them. */
export async function fetchAttachment(
  client: OutlookClient,
  messageId: string,
  attachmentId: string,
): Promise<FetchedAttachment> {
  const payload = await client.get<unknown>(
    `me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
  );
  const row =
    payload !== null && typeof payload === "object"
      ? camelizeKeys<Record<string, unknown>>(payload)
      : {};
  const contentBytes = typeof row.contentBytes === "string" ? row.contentBytes : undefined;
  return {
    meta: metaOf(row, attachmentId),
    ...(contentBytes !== undefined ? { contentBytes } : {}),
  };
}

/**
 * Text-ish content types get decoded rather than saved: `text/*` plus the
 * calendar media types iCalendar clients use.
 */
export function isTextualContentType(contentType: string): boolean {
  const type = contentType.toLowerCase().split(";")[0]?.trim() ?? "";
  return type.startsWith("text/") || type === "application/calendar" || type.startsWith("application/calendar+");
}

export type DecodedContent = { bytes: Buffer } | { error: string };

/** Strict-enough base64 decode; rejects shapes Node would silently truncate. */
export function decodeContentBytes(value: string): DecodedContent {
  const cleaned = value.replace(/\s+/g, "");
  if (cleaned === "" || cleaned.length % 4 === 1 || !/^[A-Za-z0-9+/]*={0,2}$/.test(cleaned)) {
    return { error: "ContentBytes is not valid base64" };
  }
  return { bytes: Buffer.from(cleaned, "base64") };
}

/**
 * Basename-only, traversal-free, ≤100-character attachment filename.
 * Control characters and separators are replaced, leading dots dropped.
 */
export function sanitizeAttachmentName(raw: string | undefined): string {
  if (typeof raw !== "string") return "attachment";
  // Treat both separators as path separators so a Windows-style name cannot
  // smuggle ".." through as ordinary characters.
  const base = raw.replace(/\\/g, "/").split("/").pop() ?? "";
  let name = base.replace(/[\u0000-\u001f\u007f/\\]/g, "_").trim();
  name = name.replace(/^\.+/, "");
  if (name === "") return "attachment";
  if (name.length > 100) {
    const dot = name.lastIndexOf(".");
    const extension = dot > 0 && name.length - dot <= 20 ? name.slice(dot) : "";
    const stem = name.slice(0, name.length - extension.length);
    name = `${stem.slice(0, 100 - extension.length)}${extension}`;
  }
  return name;
}

/** Ids are used inside the filename; keep them path-safe and short enough to leave room for the name. */
function safeSegment(value: string, fallback: string): string {
  const cleaned = value.replace(/[^A-Za-z0-9_-]/g, "_").slice(0, 24);
  return cleaned === "" ? fallback : cleaned;
}

/**
 * Write decoded bytes under `<cacheDir>/attachments` (0700), file 0600.
 * The name is sanitised by the caller via `sanitizeAttachmentName`; ids are
 * reduced to path-safe segments here. A random suffix keeps concurrent saves
 * of the same attachment distinct. The file is never opened or executed.
 */
export async function saveAttachment(
  cacheDir: string,
  messageId: string,
  attachmentId: string,
  name: string,
  bytes: Buffer,
): Promise<string> {
  const dir = resolve(cacheDir, "attachments");
  await mkdir(dir, { recursive: true, mode: 0o700 });
  await chmod(dir, 0o700);

  const safeName = sanitizeAttachmentName(name);
  const dot = safeName.lastIndexOf(".");
  const extension = dot > 0 ? safeName.slice(dot) : "";
  const stem = dot > 0 ? safeName.slice(0, dot) : safeName;
  const prefix = `${safeSegment(messageId, "message")}-${safeSegment(attachmentId, "attachment")}-`;

  for (let attempt = 0; attempt < 4; attempt += 1) {
    const suffix = randomBytes(4).toString("hex");
    const filename = `${prefix}${stem.slice(0, Math.max(1, 100 - prefix.length - extension.length - 9))}-${suffix}${extension}`;
    const path = resolve(dir, filename);
    if (!path.startsWith(`${dir}${sep}`)) {
      throw new Error("refusing to write outside the attachment cache directory");
    }
    try {
      await writeFile(path, bytes, { flag: "wx", mode: 0o600 });
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "EEXIST") continue;
      throw error;
    }
    await chmod(path, 0o600);
    return path;
  }
  throw new Error("could not allocate a unique attachment filename");
}
