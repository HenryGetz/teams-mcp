/**
 * Recipient/attendee resolution shared by the mail and calendar write paths.
 *
 * An input containing `@` is an address; a configured alias is taken as-is.
 * Anything else is a human name, which is looked up in the tenant directory
 * (Graph `/users`). Resolution is deliberately conservative:
 *
 *  - exactly one directory match -> that person's address;
 *  - zero matches -> a `recipient_not_found` error, never a guessed address;
 *  - several matches -> `ambiguous_recipient` listing every candidate, so the
 *    caller chooses instead of the server picking one.
 *
 * Resolution happens before any allowlist check or write, so a bad name can
 * never turn into a message to an unintended mailbox.
 */
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { config } from "../config.js";
import { graphClient, searchDirectory } from "../graph/client.js";

export interface RecipientCandidate {
  id?: string;
  name?: string;
  email?: string;
}

export interface RecipientFailure {
  input: string;
  code: "recipient_not_found" | "ambiguous_recipient" | "invalid_recipient";
  message: string;
  candidates?: RecipientCandidate[];
  needs_disambiguation?: true;
}

export type RecipientResolution =
  | { ok: true; addresses: string[] }
  | { ok: false; result: CallToolResult };

/** Directory rows that carry a usable address, de-duplicated case-insensitively. */
async function directoryMatches(name: string): Promise<RecipientCandidate[]> {
  const payload = await searchDirectory(graphClient(), name, config().maxResults.people);
  const seen = new Set<string>();
  const matches: RecipientCandidate[] = [];
  for (const user of payload.value ?? []) {
    const email = user.mail ?? user.userPrincipalName;
    if (!email) continue;
    const key = email.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    matches.push({
      ...(user.id ? { id: user.id } : {}),
      ...(user.displayName ? { name: user.displayName } : {}),
      email,
    });
  }
  return matches;
}

/**
 * Bare addr-spec from a recipient string.
 *
 * Callers paste `Joel <jd@example.com>` or `"jd@example.com"`; sending the whole
 * string as the REST `Address` misparses, and the allowlist compares the wrong
 * thing. Extract the angle-bracket address when present, else unquote.
 */
export function addrSpec(value: string): string {
  const angled = /<\s*([^<>@\s]+@[^<>@\s]+)\s*>/.exec(value);
  if (angled?.[1]) return angled[1];
  return value
    .trim()
    .replace(/^["']+|["']+$/g, "")
    .trim();
}

/**
 * One bare addr-spec: exactly one `@`, no list separators, display parts,
 * whitespace or control characters.
 *
 * `addrSpec()` unwraps `Name <addr>` but deliberately leaves garbage like
 * `a@b@c` alone. That garbage must not reach a write: a comma or CRLF inside a
 * single `to[]` element could smuggle a second recipient past an allowlist that
 * only compares the whole string (`attacker@evil.com, victim@allowed.com` ends
 * with `@allowed.com`).
 */
export function isSingleAddress(value: string): boolean {
  return (
    /^[^\s@,;<>()[\]\\"]+@[^\s@,;<>()[\]\\"]+$/.test(value) &&
    !/[\u0000-\u001f\u007f-\u009f]/.test(value)
  );
}

function failureResult(failures: RecipientFailure[]): CallToolResult {
  const ambiguous = failures.some((failure) => failure.code === "ambiguous_recipient");
  const body =
    failures.length === 1
      ? {
          error: failures[0]?.code,
          input: failures[0]?.input,
          message: failures[0]?.message,
          ...(failures[0]?.candidates ? { candidates: failures[0].candidates } : {}),
          ...(ambiguous ? { needs_disambiguation: true } : {}),
        }
      : { error: "recipient_resolution_failed", failures, ...(ambiguous ? { needs_disambiguation: true } : {}) };
  return { content: [{ type: "text", text: JSON.stringify(body) }], isError: true };
}

/**
 * Resolve a recipient/attendee list. Order is preserved so a combined
 * to/cc/bcc list can be split by the caller's original lengths.
 */
export async function resolveRecipients(values: string[]): Promise<RecipientResolution> {
  const aliases = config().aliases;
  const addresses: string[] = [];
  const failures: RecipientFailure[] = [];
  // One directory search per distinct name, even when it appears in to/cc/bcc.
  const lookups = new Map<string, Promise<RecipientCandidate[]>>();

  for (const raw of values) {
    const value = raw.trim();
    if (!value) continue;
    const target = (aliases[value.toLowerCase()] ?? value).trim();
    if (target.includes("@")) {
      const address = addrSpec(target);
      if (!isSingleAddress(address)) {
        failures.push({
          input: value,
          code: "invalid_recipient",
          message: `"${value}" is not a single email address; pass one address per element`,
        });
        continue;
      }
      addresses.push(address);
      continue;
    }
    const key = target.toLowerCase();
    let pending = lookups.get(key);
    if (!pending) {
      pending = directoryMatches(target);
      lookups.set(key, pending);
    }
    const matches = await pending;
    if (matches.length === 1 && matches[0]?.email) {
      addresses.push(matches[0].email);
      continue;
    }
    if (matches.length === 0) {
      failures.push({
        input: value,
        code: "recipient_not_found",
        message: `no directory match for "${value}"; pass a full email address`,
      });
      continue;
    }
    failures.push({
      input: value,
      code: "ambiguous_recipient",
      message: `${matches.length} directory matches for "${value}"; choose one email address`,
      candidates: matches,
      needs_disambiguation: true,
    });
  }

  if (failures.length > 0) return { ok: false, result: failureResult(failures) };
  return { ok: true, addresses };
}
