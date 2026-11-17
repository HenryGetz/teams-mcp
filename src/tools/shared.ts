/**
 * Shared plumbing for the action-dispatch tools: result envelopes, op
 * extraction, and the single unknown-op error shape.
 */
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import type { ZodType } from "zod";
import { errorText, toJsonText } from "../format.js";
import { config, isValidTimeZone } from "../config.js";
import { ApiError } from "../http/client.js";
import { AuthRequiredError } from "../auth/store.js";
import { SnapshotError } from "../snapshot/calendar.js";

export interface ToolModule {
  /** Tool name, matching the `op`-dispatched surface in docs/design.md. */
  name: string;
  title: string;
  description: string;
  /** Flat argument shape; every op shares it to keep tool-schema tokens low. */
  shape: Record<string, ZodType>;
  handle(args: Record<string, unknown>): Promise<CallToolResult>;
}

export function textResult(value: unknown, space = 0): CallToolResult {
  return { content: [{ type: "text", text: toJsonText(value, space) }] };
}

export function errorResult(code: string, message: string, hint?: string): CallToolResult {
  return { content: [{ type: "text", text: errorText(code, message, hint) }], isError: true };
}

/**
 * Unknown or missing `op` for an op-dispatched tool. Always an error result so
 * an agent can never mistake it for data, or for a feature arriving later.
 */
export function unknownOp(tool: string, op: string, ops: readonly string[]): CallToolResult {
  return errorResult(
    "invalid_argument",
    op ? `unknown op "${op}" for ${tool}` : `${tool} requires op`,
    `available ops: ${ops.join(", ")}`,
  );
}

export function opOf(args: Record<string, unknown>, ops: readonly string[]): string {
  const op = typeof args.op === "string" ? args.op : "";
  if (!op) return "";
  return ops.includes(op) ? op : "";
}

export function stringArg(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed === "" ? undefined : trimmed;
}

export function numberArg(args: Record<string, unknown>, key: string): number | undefined {
  const value = args[key];
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

/**
 * A page size clamped to [1, max]. The MCP schema already enforces this, but a
 * negative or fractional value reaching the handler directly must not become
 * `$top=-5` on the provider request.
 */
export function pageSize(
  args: Record<string, unknown>,
  key: string,
  fallback: number,
  max: number,
): number {
  const value = numberArg(args, key);
  return value === undefined ? fallback : Math.min(Math.max(1, Math.trunc(value)), max);
}

export function booleanArg(args: Record<string, unknown>, key: string): boolean | undefined {
  const value = args[key];
  return typeof value === "boolean" ? value : undefined;
}

/**
 * Resolve the caller's `timezone` argument (or the configured default) and verify
 * it is an IANA zone. Rendering dates through an unknown zone throws a
 * `RangeError` from `Intl`; that must be a tool error, not a crash.
 */
export function timezoneOf(
  args: Record<string, unknown>,
  fallback: string,
): { timezone: string } | { result: CallToolResult } {
  const supplied = typeof args.timezone === "string" ? args.timezone.trim() : "";
  const timezone = supplied === "" ? fallback : supplied;
  if (!isValidTimeZone(timezone)) {
    return {
      result: errorResult(
        "invalid_argument",
        `"${timezone}" is not an IANA time zone`,
        "use a zone like America/Chicago or UTC",
      ),
    };
  }
  return { timezone };
}

export function arrayArg(args: Record<string, unknown>, key: string): string[] | undefined {
  const value = args[key];
  if (!Array.isArray(value)) return undefined;
  const items = value.filter((item): item is string => typeof item === "string" && item.trim() !== "");
  return items.length > 0 ? items : undefined;
}

/**
 * Map a known failure onto a tool error. Returns undefined for unexpected
 * errors, which the MCP runtime converts into an error result itself.
 */
export function toolError(error: unknown): CallToolResult | undefined {
  if (error instanceof ApiError) {
    const hint =
      error.hint ??
      (/^(ErrorInvalidIdMalformed|ErrorInvalidId)$/.test(error.code)
        ? "ids come from mail.list or calendar.list/get; copy the id verbatim"
        : undefined);
    return errorResult(error.code, error.message, hint);
  }
  if (error instanceof AuthRequiredError) return errorResult(error.code, error.message, error.hint);
  if (error instanceof SnapshotError) return errorResult(error.code, error.message, error.hint);
  return undefined;
}

/** Opaque cursor = base64url of a provider nextLink. */
export function encodeCursor(nextLink: string): string {
  return Buffer.from(nextLink, "utf8").toString("base64url");
}

export function decodeCursor(cursor: string | undefined): string | undefined {
  if (!cursor) return undefined;
  try {
    const decoded = Buffer.from(cursor, "base64url").toString("utf8");
    return /^https?:\/\//i.test(decoded) ? decoded : undefined;
  } catch {
    return undefined;
  }
}

export interface CursorLookup {
  /** The provider nextLink, when a cursor was supplied and decodable. */
  link?: string;
  /** Error to return when the caller supplied something that is not a cursor. */
  result?: CallToolResult;
}

/**
 * Read `cursor` strictly: absent means page one, but a supplied value that does
 * not decode is an error rather than a silent restart from the first page.
 */
export function resolveCursor(args: Record<string, unknown>, key = "cursor"): CursorLookup {
  const supplied = stringArg(args, key);
  if (!supplied) return {};
  const link = decodeCursor(supplied);
  if (!link) {
    return {
      result: errorResult(
        "invalid_cursor",
        `"${key}" is not a teams-mcp cursor`,
        "omit it to start from the first page, or pass a next_cursor from a prior call",
      ),
    };
  }
  return { link };
}

/**
 * Recipients rejected by `invite_allowlist`. An empty list means no
 * restriction; entries match a full address or a domain suffix
 * (`@example.com`). Case-insensitive.
 */
export function disallowedRecipients(addresses: string[]): string[] {
  const allowlist = config().inviteAllowlist;
  if (allowlist.length === 0) return [];
  return addresses.filter((address) => {
    const lower = address.trim().toLowerCase();
    return !allowlist.some((entry) => {
      const rule = entry.trim().toLowerCase();
      return rule.startsWith("@") ? lower.endsWith(rule) : lower === rule;
    });
  });
}
