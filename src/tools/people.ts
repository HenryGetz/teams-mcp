/**
 * `people` tool: directory lookups through Microsoft Graph (`/users`).
 *
 * Graph is used ONLY for directory search in this tenant (app assignment for
 * mail/calendar is blocked). Results are cached in-process for
 * `cache.directory_ttl_minutes`.
 *
 * Ambiguity is data, not a decision: every candidate is returned, with
 * `ambiguous: true` and a `needs_disambiguation` note when a name matches more
 * than one person. Nothing is picked silently. Configured aliases resolve
 * exactly, bypassing search.
 */
import { z } from "zod";
import { config, resolveAlias } from "../config.js";
import {
  getUser,
  graphClient,
  searchDirectory,
  type DirectoryUser,
  type GraphClient,
} from "../graph/client.js";
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

const OPS = ["search", "resolve"] as const;

interface PersonRow {
  id?: string;
  name?: string;
  email?: string;
  title?: string;
  department?: string;
}

interface CacheEntry {
  rows: PersonRow[];
  /** The page filled `$top`, so more directory rows may exist. */
  truncated: boolean;
  at: number;
}

const searchCache = new Map<string, CacheEntry>();

/** Distinct-query cache entries retained before the oldest is evicted. */
const SEARCH_CACHE_LIMIT = 100;

/**
 * `resolve` fans out one GET per input; an unbounded array would open hundreds
 * of sockets and invite provider throttling. Plenty for a picker-sized batch.
 */
const MAX_RESOLVE_INPUTS = 50;

function personRow(user: DirectoryUser): PersonRow {
  return {
    ...(user.id ? { id: user.id } : {}),
    ...(user.displayName ? { name: user.displayName } : {}),
    ...(user.mail || user.userPrincipalName ? { email: user.mail ?? user.userPrincipalName } : {}),
    ...(user.jobTitle ? { title: user.jobTitle } : {}),
    ...(user.department ? { department: user.department } : {}),
  };
}

function isEmptyRow(row: PersonRow): boolean {
  return !row.id && !row.name && !row.email;
}

async function lookup(client: GraphClient, value: string): Promise<PersonRow | undefined> {
  try {
    const user = await getUser(client, value);
    const row = personRow(user);
    return isEmptyRow(row) ? undefined : row;
  } catch (error) {
    // A missing mailbox/id is data, not a failure: report `found: false`.
    if (
      typeof error === "object" &&
      error !== null &&
      "status" in error &&
      (error as { status?: number }).status === 404
    ) {
      return undefined;
    }
    throw error;
  }
}

export const peopleTool: ToolModule = {
  name: "people",
  title: "People",
  description:
    "Directory (Graph /users). op=search (query, top<=25) matches displayName/mail/userPrincipalName; configured aliases resolve exactly; multiple matches are returned in full with ambiguous=true + needs_disambiguation, never silently picked | op=resolve (ids[] and/or emails[], max 50 per call) returns one row per input, found=false when absent. Rows: id, name, email, title, department. Cached in-process for the configured TTL; refresh=true bypasses.",
  shape: {
    op: z.enum(OPS),
    query: z.string().optional().describe("search: name, email, partial display name or configured alias"),
    ids: z.array(z.string()).optional().describe("resolve: Graph user ids"),
    emails: z.array(z.string()).optional().describe("resolve: addresses/UPNs (aliases allowed)"),
    top: z.number().int().min(1).max(25).optional().describe("search: max rows, default 10"),
    refresh: z.boolean().optional().describe("bypass the in-process cache"),
  },
  async handle(args) {
    const op = opOf(args, OPS);
    if (!op) return unknownOp("people", String(args.op ?? ""), OPS);
    const cfg = config();
    const client = graphClient();

    try {
      if (op === "search") {
        const query = stringArg(args, "query");
        if (!query) return errorResult("invalid_argument", "search requires query");
        const top = pageSize(args, "top", 10, cfg.maxResults.people);

        const aliasTarget = cfg.aliases[query.toLowerCase()];
        if (aliasTarget) {
          const row = await lookup(client, aliasTarget);
          return textResult({
            query,
            alias: true,
            resolved_to: aliasTarget,
            people: row ? [row] : [],
            ...(row ? {} : { found: false }),
          });
        }

        const key = `${query.toLowerCase()}|${top}`;
        const cached = searchCache.get(key);
        if (!booleanArg(args, "refresh") && cached && Date.now() - cached.at < cfg.cache.directoryTtlSeconds * 1000) {
          return textResult({
            query,
            count: cached.rows.length,
            people: cached.rows,
            cached: true,
            ...(cached.truncated ? { truncated: true, note: capNote(top, cfg.maxResults.people) } : {}),
            ...ambiguity(cached.rows),
          });
        }
        const payload = await searchDirectory(client, query, top);
        const rows = (payload.value ?? []).map(personRow).filter((row) => !isEmptyRow(row));
        const truncated = rows.length >= top;
        searchCache.set(key, { rows, truncated, at: Date.now() });
        while (searchCache.size > SEARCH_CACHE_LIMIT) {
          let oldestKey: string | undefined;
          let oldestAt = Infinity;
          for (const [candidate, entry] of searchCache) {
            if (entry.at < oldestAt) {
              oldestAt = entry.at;
              oldestKey = candidate;
            }
          }
          if (oldestKey === undefined) break;
          searchCache.delete(oldestKey);
        }
        return textResult({
          query,
          count: rows.length,
          people: rows,
          // `$top` is a cap, not a promise that nothing else matched.
          ...(truncated ? { truncated: true, note: capNote(top, cfg.maxResults.people) } : {}),
          ...ambiguity(rows),
        });
      }

      const ids = (arrayArg(args, "ids") ?? []).map((value) => resolveAlias(value));
      const emails = (arrayArg(args, "emails") ?? []).map((value) => resolveAlias(value));
      if (ids.length === 0 && emails.length === 0) {
        return errorResult("invalid_argument", "resolve requires ids[] and/or emails[]");
      }
      const inputs = [...ids, ...emails];
      if (inputs.length > MAX_RESOLVE_INPUTS) {
        return errorResult(
          "invalid_argument",
          `resolve accepts at most ${MAX_RESOLVE_INPUTS} ids/emails per call (got ${inputs.length})`,
          "split the list across calls",
        );
      }
      const results = await Promise.all(
        inputs.map(async (input) => {
          const row = await lookup(client, input);
          return row ? { input, found: true, ...row } : { input, found: false };
        }),
      );
      return textResult({ people: results });
    } catch (error) {
      const mapped = toolError(error);
      if (mapped) return mapped;
      throw error;
    }
  },
};

/** Why a full page may be incomplete, phrased for the caller. */
function capNote(top: number, max: number): string {
  return `requested at most ${top} directory rows; refine the query or raise top (max ${max}) to see more`;
}

/** Ambiguity marker for multi-match searches. */
function ambiguity(rows: PersonRow[]): Record<string, unknown> {
  if (rows.length <= 1) return {};
  return {
    ambiguous: true,
    needs_disambiguation: "multiple directory matches; choose one by id or email before acting",
  };
}

export function clearPeopleCache(): void {
  searchCache.clear();
}
