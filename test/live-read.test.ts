/**
 * Live read-only controls for the two contract fixes that a mock cannot prove:
 *
 *  1. `mail.list` with `from` must not produce Exchange `InefficientFilter`.
 *  2. A fake mailbox must be reported as `unknown`, never as free.
 *
 * Opt-in only (`TEAMS_MCP_LIVE_TEST=1`); every call is a read. The token cache is
 * COPIED into a temp config dir, so a refresh (if the token is stale) writes the
 * copy rather than the canonical `~/.config/teams-mcp/token.json`. Read-only mode
 * is forced on, so a write could not slip through even if the test were wrong.
 */
import { copyFile, mkdtemp, rm } from "node:fs/promises";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { clearSnapshotCache } from "../src/snapshot/calendar.js";
import { availabilityTool } from "../src/tools/availability.js";
import { mailTool } from "../src/tools/mail.js";

const ENABLED = process.env.TEAMS_MCP_LIVE_TEST === "1";
const CANONICAL_TOKEN = join(homedir(), ".config", "teams-mcp", "token.json");
const FAKE = "teams-mcp-no-such-mailbox-9f3a@outlook.com";

const saved = { ...process.env };
let dir: string | undefined;
let self: string | undefined;

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

describe.skipIf(!ENABLED)("live read-only controls", () => {
  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), "teams-mcp-live-read-"));
    // A copy, never the canonical file: a token refresh writes the copy.
    await copyFile(CANONICAL_TOKEN, join(dir, "token.json"));
    process.env.TEAMS_MCP_CONFIG_DIR = dir;
    process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
    process.env.TEAMS_MCP_SNAPSHOT_DIR = join(dir, "cache", "snapshots");
    process.env.TEAMS_MCP_READ_ONLY = "1";
    delete process.env.TEAMS_MCP_OUTLOOK_BASE_URL;
    delete process.env.TEAMS_MCP_GRAPH_BASE_URL;
    resetConfigCache();
    clearSnapshotCache();
  });

  afterAll(async () => {
    process.env = { ...saved };
    resetConfigCache();
    clearSnapshotCache();
    if (dir) await rm(dir, { recursive: true, force: true });
  });

  it("mail.list with from is accepted by Exchange and ordered newest first", async () => {
    const first = await mailTool.handle({ op: "list", top: 5 });
    expect(first.isError, textOf(first)).toBeUndefined();
    const rows = (payload(first).messages ?? []) as Array<Record<string, unknown>>;
    if (rows.length === 0) return; // empty inbox: nothing to filter by
    const sender = String(rows[0]?.from ?? "");
    expect(sender).toContain("@");

    const filtered = await mailTool.handle({
      op: "list",
      from: sender,
      top: 5,
    });
    expect(filtered.isError, textOf(filtered)).toBeUndefined();
    expect(textOf(filtered)).not.toContain("InefficientFilter");
    const filteredRows = (payload(filtered).messages ?? []) as Array<Record<string, unknown>>;
    expect(filteredRows.length).toBeGreaterThan(0);
    for (const row of filteredRows) {
      expect(String(row.from).toLowerCase()).toBe(sender.toLowerCase());
    }
    const received = filteredRows.map((row) => String(row.received ?? ""));
    expect(received).toEqual([...received].sort((a, b) => b.localeCompare(a)));

    // Regression: a small `top` on the from path used to return the OLDEST
    // matches. The newest of top=1 must equal the newest of the wider page.
    const newestOnly = await mailTool.handle({ op: "list", from: sender, top: 1 });
    expect(newestOnly.isError, textOf(newestOnly)).toBeUndefined();
    const newestRows = (payload(newestOnly).messages ?? []) as Array<Record<string, unknown>>;
    expect(newestRows).toHaveLength(1);
    expect(String(newestRows[0]?.received ?? "")).toBe(received[0]);
    expect(Number(payload(newestOnly).scanned_pages)).toBeGreaterThanOrEqual(1);
  });

  it("a fake mailbox is unknown, never free", async () => {
    const status = await availabilityTool.handle({ op: "freebusy" });
    const statusBody = payload(status);
    const own = (statusBody.attendees ?? []) as Array<Record<string, unknown>>;
    self = typeof own[0]?.email === "string" ? own[0].email : undefined;
    expect(self).toBeDefined();

    const result = await availabilityTool.handle({
      op: "freebusy",
      attendees: [FAKE, self as string],
    });
    expect(result.isError, textOf(result)).toBeUndefined();
    const out = payload(result);
    const unknown = (out.unknown ?? []) as string[];
    expect(unknown.map((email) => email.toLowerCase())).toContain(FAKE.toLowerCase());
    expect(out.warning).toContain("unknown_mailbox");
    const view = out.availability_view as Record<string, string>;
    expect(Object.keys(view).map((key) => key.toLowerCase())).not.toContain(FAKE.toLowerCase());
    const fakeRow = (out.attendees as Array<Record<string, unknown>>).find(
      (row) => String(row.email).toLowerCase() === FAKE.toLowerCase(),
    );
    expect(fakeRow).toMatchObject({ unknown: true, busy: [] });
  });

  it("suggest excludes the fake mailbox instead of ranking it free", async () => {
    if (!self) return;
    const result = await availabilityTool.handle({
      op: "suggest",
      attendees: [self, FAKE],
      duration_min: 30,
      max_candidates: 3,
    });
    expect(result.isError, textOf(result)).toBeUndefined();
    const out = payload(result);
    const excluded = ((out.excluded_unknown_attendees ?? []) as string[]).map((email) =>
      email.toLowerCase(),
    );
    expect(excluded).toContain(FAKE.toLowerCase());
    expect(out.warning).toContain("excluded from ranking");
    const candidates = (out.candidates ?? []) as Array<Record<string, unknown>>;
    expect(JSON.stringify(candidates).toLowerCase()).not.toContain(FAKE.toLowerCase());
  });
});
