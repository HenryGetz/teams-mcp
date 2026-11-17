#!/usr/bin/env node
/**
 * Live write probe through the real MCP server (stdio, as an agent would call it).
 *
 *   LIVE_WRITES=1 node scripts/dev/live-writes.mjs
 *
 * Safety contract:
 *  - reads the canonical token cache READ-ONLY; aborts if the outlook token is
 *    stale instead of refreshing (the auth worker owns refresh/rotation);
 *  - creates exactly one calendar event (+2h, 30 min, no attendees), then
 *    updates, reads and deletes it;
 *  - sends exactly one email, to the signed-in user only;
 *  - never invites anyone, never touches a pre-existing event or message, no
 *    bulk operations, and never calls unsubscribe_apply.
 */
import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const MARKER = "AI TEST — teams-mcp — write probe";
const SKEW_MS = 5 * 60 * 1000;

function tokenCachePath() {
  const dir = process.env.TEAMS_MCP_CONFIG_DIR ?? join(homedir(), ".config", "teams-mcp");
  return join(dir, "token.json");
}

/** The signed-in address, from the token cache (never a hard-coded mailbox). */
function selfAddress() {
  try {
    const doc = JSON.parse(readFileSync(tokenCachePath(), "utf8"));
    const upn = doc?.account?.upn;
    if (typeof upn === "string" && upn.includes("@")) return upn;
  } catch {
    // fall through to the explicit override
  }
  return process.env.TEAMS_MCP_SELF ?? "you@example.com";
}

function assertFreshOutlookToken() {
  const path = tokenCachePath();
  let doc;
  try {
    doc = JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    throw new Error(`cannot read token cache at ${path}: ${error.message}`);
  }
  const entry = doc?.tokens?.outlook;
  if (!entry?.access_token || !entry?.expires_at) {
    throw new Error(`no outlook token in ${path}; sign in first (this script never logs in)`);
  }
  const remaining = Date.parse(entry.expires_at) - Date.now();
  if (remaining < SKEW_MS) {
    throw new Error(
      `outlook token is stale (${Math.round(remaining / 1000)}s left, need >300s). ` +
        "Aborting without refreshing; run the auth worker's token refresh.",
    );
  }
  return { path, expires_at: entry.expires_at, remaining_s: Math.round(remaining / 1000) };
}

function textOf(result) {
  const first = result?.content?.[0];
  if (!first || first.type !== "text") throw new Error("tool returned no text content");
  return first.text;
}

function parse(result, label) {
  if (result?.isError) throw new Error(`${label} failed: ${textOf(result)}`);
  return JSON.parse(textOf(result));
}

function step(name, detail) {
  console.log(`\n## ${name}`);
  if (detail !== undefined) console.log(typeof detail === "string" ? detail : JSON.stringify(detail, null, 2));
}

async function main() {
  if (process.env.LIVE_WRITES !== "1") {
    console.error("refusing to run: set LIVE_WRITES=1 (this script creates and deletes real objects)");
    process.exit(2);
  }
  const token = assertFreshOutlookToken();
  console.log(`token cache ok: ${token.path} (outlook expires ${token.expires_at}, ${token.remaining_s}s left)`);

  const transport = new StdioClientTransport({
    command: process.execPath,
    args: ["dist/index.js"],
    cwd: process.cwd(),
    env: { ...process.env },
    stderr: "inherit",
  });
  const client = new Client({ name: "teams-mcp-live-writes", version: "0.1.0" });
  await client.connect(transport);

  const summary = { calendar: {}, mail: {}, unsubscribe: {} };

  try {
    // ---------------------------------------------------------------- calendar
    const start = new Date(Date.now() + 2 * 60 * 60 * 1000);
    start.setSeconds(0, 0);
    const end = new Date(start.getTime() + 30 * 60 * 1000);

    const created = parse(
      await client.callTool({
        name: "calendar",
        arguments: {
          op: "create",
          subject: MARKER,
          start: start.toISOString(),
          end: end.toISOString(),
        },
      }),
      "calendar.create",
    );
    if (!created.id) throw new Error("calendar.create returned no id");
    step("calendar.create", { id: created.id, subject: created.subject, local: created.local });
    summary.calendar.created = created.id;

    const updatedSubject = `${MARKER} updated`;
    const updated = parse(
      await client.callTool({
        name: "calendar",
        arguments: {
          op: "update",
          event_id: created.id,
          expected_subject: MARKER,
          subject: updatedSubject,
        },
      }),
      "calendar.update",
    );
    if (updated.subject !== updatedSubject) {
      throw new Error(`calendar.update subject did not change (got "${updated.subject}")`);
    }
    step("calendar.update", { id: updated.id, subject: updated.subject });
    summary.calendar.updated = updated.id;

    const fetched = parse(
      await client.callTool({ name: "calendar", arguments: { op: "get", event_id: created.id } }),
      "calendar.get",
    );
    if (fetched.subject !== updatedSubject) throw new Error("calendar.get did not show the updated subject");
    step("calendar.get", { id: fetched.id, subject: fetched.subject, local: fetched.local });

    // Guard check: a wrong expected_subject must be refused (no delete).
    const guarded = await client.callTool({
      name: "calendar",
      arguments: { op: "delete", event_id: created.id, expected_subject: "definitely not the subject" },
    });
    if (!guarded.isError) throw new Error("calendar.delete guard did not refuse a wrong subject");
    step("calendar.delete guard", textOf(guarded));

    const deleted = parse(
      await client.callTool({
        name: "calendar",
        arguments: { op: "delete", event_id: created.id, expected_subject: updatedSubject },
      }),
      "calendar.delete",
    );
    if (deleted.deleted !== true) throw new Error("calendar.delete did not report deleted:true");
    step("calendar.delete", deleted);
    summary.calendar.deleted = created.id;

    const listed = parse(
      await client.callTool({
        name: "calendar",
        arguments: {
          op: "list",
          start: new Date(Date.now() - 60 * 60 * 1000).toISOString(),
          end: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
        },
      }),
      "calendar.list",
    );
    const leftover = (listed.events ?? []).filter((event) => event.subject === updatedSubject);
    if (leftover.length > 0) throw new Error(`deleted event still listed: ${leftover.map((e) => e.id).join(", ")}`);
    step("calendar.list after delete", { count: listed.count, leftover: leftover.length });

    // -------------------------------------------------------------------- mail
    const self = selfAddress();
    const sent = parse(
      await client.callTool({
        name: "mail",
        arguments: {
          op: "send",
          to: [self],
          subject: MARKER,
          body: "Automated write probe. Safe to delete.",
        },
      }),
      "mail.send",
    );
    if (sent.ok !== true || sent.recipient_count !== 1) throw new Error("mail.send did not report ok with 1 recipient");
    step("mail.send", sent);
    summary.mail.sent_to = self;

    let found;
    for (let attempt = 0; attempt < 6 && !found; attempt += 1) {
      if (attempt > 0) await new Promise((resolve) => setTimeout(resolve, 4000));
      const search = parse(
        await client.callTool({ name: "mail", arguments: { op: "search", query: MARKER, top: 5 } }),
        "mail.search",
      );
      found = (search.messages ?? []).find((message) => message.subject === MARKER);
    }
    if (!found?.id) throw new Error("sent message did not appear in mail.search after ~20s");
    step("mail.search", { id: found.id, subject: found.subject, from: found.from });
    summary.mail.id = found.id;

    await parse(
      await client.callTool({ name: "mail", arguments: { op: "mark", message_id: found.id, mark: "unread" } }),
      "mail.mark unread",
    );
    let after = parse(
      await client.callTool({ name: "mail", arguments: { op: "get", message_id: found.id } }),
      "mail.get after unread",
    );
    if (after.unread !== true) throw new Error("mail.mark unread was not reflected");
    step("mail.mark unread", { id: found.id, unread: after.unread });

    await parse(
      await client.callTool({ name: "mail", arguments: { op: "mark", message_id: found.id, mark: "read" } }),
      "mail.mark read",
    );
    after = parse(
      await client.callTool({ name: "mail", arguments: { op: "get", message_id: found.id } }),
      "mail.get after read",
    );
    if (after.unread !== false) throw new Error("mail.mark read was not reflected");
    step("mail.mark read", { id: found.id, unread: after.unread });

    const moved = parse(
      await client.callTool({
        name: "mail",
        arguments: { op: "move", message_id: found.id, destination: "archive" },
      }),
      "mail.move",
    );
    if (!moved.id || moved.id === found.id) throw new Error("mail.move did not return a new id");
    step("mail.move", moved);
    summary.mail.moved_id = moved.id;

    const archived = parse(
      await client.callTool({ name: "mail", arguments: { op: "get", message_id: moved.id } }),
      "mail.get after move",
    );
    if (archived.subject !== MARKER) throw new Error("archived message subject changed unexpectedly");
    step("mail.get after move", { id: archived.id, subject: archived.subject });
    summary.mail.verified = true;

    // -------------------------------------------------------------- unsubscribe
    const inbox = parse(
      await client.callTool({ name: "mail", arguments: { op: "list", folder: "inbox", top: 8 } }),
      "mail.list for unsubscribe",
    );
    let preview;
    for (const message of inbox.messages ?? []) {
      const result = await client.callTool({
        name: "mail",
        arguments: { op: "unsubscribe_preview", message_id: message.id },
      });
      if (result.isError) continue;
      const body = JSON.parse(textOf(result));
      if (body.unsubscribe) {
        preview = { message_id: body.message_id, subject: body.subject, unsubscribe: body.unsubscribe };
        break;
      }
    }
    step("unsubscribe_preview", preview ?? "no message in the latest 8 carried List-Unsubscribe");
    summary.unsubscribe.preview = preview ?? null;
    console.log("\nunsubscribe_apply was intentionally NOT called.");
  } finally {
    await client.close();
  }

  console.log("\n## summary");
  console.log(JSON.stringify(summary, null, 2));
}

main().catch((error) => {
  console.error(`\nLIVE WRITE PROBE FAILED: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
});
