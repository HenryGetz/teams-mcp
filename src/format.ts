/**
 * Output formatting for tool results.
 *
 * Rules enforced here (see docs/design.md):
 *  - never emit raw Graph JSON envelopes;
 *  - drop null/undefined instead of spelling them out;
 *  - truncate long strings;
 *  - stable key order so repeated calls diff cleanly;
 *  - ISO 8601 timestamps everywhere.
 */

export const DEFAULT_PREVIEW_LENGTH = 300;

export type Json = null | boolean | number | string | Json[] | { [key: string]: Json };

/** Recursively remove null/undefined object members (array holes are filtered of nulls). */
export function omitNullish<T>(value: T): T {
  if (Array.isArray(value)) {
    const mapped = value
      .filter((item) => item !== null && item !== undefined)
      .map((item) => omitNullish(item));
    return mapped as unknown as T;
  }
  if (value !== null && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
      if (item === null || item === undefined) continue;
      out[key] = omitNullish(item);
    }
    return out as T;
  }
  return value;
}

/** Collapse all whitespace runs to single spaces, then cut to `max` characters. */
export function truncate(value: string | null | undefined, max = DEFAULT_PREVIEW_LENGTH): string {
  if (value === null || value === undefined) return "";
  const flat = value.replace(/\s+/g, " ").trim();
  if (max <= 0) return "";
  if (flat.length <= max) return flat;
  return `${flat.slice(0, Math.max(0, max - 1)).trimEnd()}…`;
}

/**
 * Cut to `max` characters without flattening the text.
 *
 * Body reads keep their line structure — a reply whose paragraphs were squeezed
 * onto one line is not the mail that was sent. (`truncate` is for one-line
 * previews and errors.)
 */
export function cutText(value: string | null | undefined, max: number): string {
  if (value === null || value === undefined) return "";
  if (max <= 0) return "";
  const text = value.trim();
  if (text.length <= max) return text;
  return `${text.slice(0, Math.max(0, max - 1)).trimEnd()}…`;
}

/** ISO 8601 string for a Date, epoch ms or already-ISO string; undefined when unusable. */
export function isoTime(value: Date | number | string | null | undefined): string | undefined {
  if (value === null || value === undefined) return undefined;
  if (typeof value === "number") {
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? undefined : d.toISOString();
  }
  if (typeof value === "string") {
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? undefined : d.toISOString();
  }
  return Number.isNaN(value.getTime()) ? undefined : value.toISOString();
}

/** JSON with recursively sorted keys, so output is byte-stable for identical data. */
export function stableStringify(value: unknown, space?: number): string {
  return JSON.stringify(sortKeys(value), null, space);
}

function sortKeys(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (value !== null && typeof value === "object") {
    const source = value as Record<string, unknown>;
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(source).sort()) out[key] = sortKeys(source[key]);
    return out;
  }
  return value;
}

/**
 * Text payload for a tool result: compact, null-free, key-stable JSON.
 * Strings pass through untouched (already human-readable).
 */
export function toJsonText(value: unknown, space = 0): string {
  if (typeof value === "string") return value;
  return stableStringify(omitNullish(value), space || undefined);
}

/**
 * Project a Graph-ish record onto compact keys.
 * Spec values are either a source key or a mapper. Missing/null fields vanish.
 */
export function pick<T extends Record<string, unknown>>(
  source: T,
  spec: Record<string, keyof T | ((value: T) => unknown)>,
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [shortKey, selector] of Object.entries(spec)) {
    const raw =
      typeof selector === "function" ? selector(source) : (source[selector] as unknown);
    if (raw === null || raw === undefined) continue;
    out[shortKey] = raw;
  }
  return out;
}

/** One-line error text: `code: message (hint)`, newlines flattened. */
export function errorText(code: string, message: string, hint?: string): string {
  const head = `${code}: ${truncate(message, 500) || "unknown error"}`;
  return hint ? `${head} (${truncate(hint, 200)})` : head;
}

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
  ndash: "–",
  mdash: "—",
  hellip: "…",
  rsquo: "’",
  lsquo: "‘",
  ldquo: "“",
  rdquo: "”",
};

function decodeEntities(text: string): string {
  return text
    .replace(/&#x([0-9a-f]+);/gi, (whole, hex: string) => {
      const code = Number.parseInt(hex, 16);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    })
    .replace(/&#(\d+);/g, (whole, dec: string) => {
      const code = Number(dec);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    })
    .replace(/&([a-z]+);/gi, (whole, name: string) => NAMED_ENTITIES[name.toLowerCase()] ?? whole);
}

/**
 * Plain-text rendering of an HTML mail body.
 *
 * Outlook stores replies and forwards as HTML, so the old regex strip returned
 * entity soup (`&quot;`, `&amp;`), dropped any text that looked like a tag
 * (`<angle>`), and flattened every newline — the caller could not read the mail
 * it had just sent. Block-level tags become newlines, entities are decoded, and
 * only intra-line whitespace is collapsed.
 */
export function htmlToText(html: string): string {
  return decodeEntities(
    html
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<hr\b[^>]*>/gi, "\n")
      .replace(/<li\b[^>]*>/gi, "\n- ")
      .replace(/<\/(p|div|tr|h[1-6]|blockquote|table|ul|ol)>/gi, "\n")
      .replace(/<[^>]+>/g, ""),
  )
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t\u00a0]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** Derive a compact preview from an HTML or text body. */
export function previewFromBody(body: string | null | undefined, max = DEFAULT_PREVIEW_LENGTH): string {
  if (!body) return "";
  return truncate(htmlToText(body), max);
}
