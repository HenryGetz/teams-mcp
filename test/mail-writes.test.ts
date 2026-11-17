import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { setOneClickPoster, setOneClickResolver, setSentLookupDelays } from "../src/outlook/mail.js";
import { resetWriteLimiter } from "../src/safety.js";
import { mailTool } from "../src/tools/mail.js";
import type { FetchLike } from "../src/http/client.js";

const BASE = "https://outlook.office.com/api/v2.0";

interface Call {
  method: string;
  url: string;
  body: unknown;
}

function harness(responses: Array<() => Response>) {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({
      method: init?.method ?? "GET",
      url,
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    });
    const factory = responses[Math.min(index, responses.length - 1)];
    index += 1;
    if (!factory) throw new Error("no response queued");
    return factory();
  };
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
  return calls;
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(body === undefined ? null : JSON.stringify(body), { status });
const empty = (status: number): (() => Response) => () => new Response(null, { status });

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-mail-writes-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  resetConfigCache();
  resetWriteLimiter();
  // Deterministic: no real Sent Items backoff in unit tests.
  setSentLookupDelays([0, 0]);
});

afterEach(async () => {
  setOutlookClient(undefined);
  setOneClickPoster(undefined);
  setOneClickResolver(undefined);
  setSentLookupDelays(undefined);
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("mail.send", () => {
  it("posts PascalCase sendmail and reports no fabricated id", async () => {
    const calls = harness([empty(202)]);
    const result = await mailTool.handle({
      op: "send",
      to: ["Self@Example.com"],
      cc: ["cc@example.com"],
      subject: "AI TEST — write probe",
      body: "hello",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[0]?.url).toBe(`${BASE}/me/sendmail`);
    expect(calls[0]?.method).toBe("POST");
    const body = calls[0]?.body as Record<string, unknown>;
    const message = body.Message as Record<string, unknown>;
    expect(message.Subject).toBe("AI TEST — write probe");
    expect(message.Body).toEqual({ ContentType: "Text", Content: "hello" });
    expect(message.ToRecipients).toEqual([{ EmailAddress: { Address: "Self@Example.com" } }]);
    expect(message.CcRecipients).toEqual([{ EmailAddress: { Address: "cc@example.com" } }]);
    expect(body.SaveToSentItems).toBe(true);

    const out = payload(result);
    expect(out).toMatchObject({
      ok: true,
      op: "send",
      recipient_count: 2,
      recipients: { to: ["Self@Example.com"], cc: ["cc@example.com"] },
    });
    // Empty lists are omitted, not returned as [].
    expect((out.recipients as Record<string, unknown>).bcc).toBeUndefined();
    // The 202 has no id and Sent Items returned nothing: a note, never a guess.
    expect(out.id).toBeUndefined();
    expect(String(out.note)).toContain("no single copy");
    // One 202 plus the initial Sent Items lookup and its two zero-delay retries.
    expect(calls).toHaveLength(4);
    expect(calls.slice(1).every((call) => call.url.includes("/me/mailfolders/sentitems/messages?"))).toBe(true);
  });

  it("refuses a missing recipient/body without calling the API", async () => {
    const calls = harness([empty(202)]);
    const result = await mailTool.handle({ op: "send", subject: "x", body: "y" });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("enforces invite_allowlist", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@example.com"]\n');
    resetConfigCache();
    const calls = harness([empty(202)]);
    const result = await mailTool.handle({
      op: "send",
      to: ["outsider@other.test"],
      subject: "x",
      body: "y",
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain(
      "recipient_not_allowed",
    );
  });
});

describe("mail.reply / mail.forward", () => {
  it("posts the reply comment to the reply action", async () => {
    const calls = harness([
      json({
        Id: "msg-1",
        ConversationId: "conv-1",
        Subject: "Quarterly numbers",
        From: { EmailAddress: { Address: "sender@example.com" } },
        ToRecipients: [{ EmailAddress: { Address: "me@example.com" } }],
      }),
      empty(202),
      json({ value: [] }),
      json({ value: [] }),
    ]);
    const result = await mailTool.handle({ op: "reply", message_id: "msg-1", body: "thanks" });
    expect(result.isError).toBeUndefined();
    expect(calls.map((call) => `${call.method} ${call.url}`).slice(0, 2)).toEqual([
      `GET ${BASE}/me/messages/msg-1?$select=Id%2CConversationId%2CSubject%2CFrom%2CToRecipients%2CCcRecipients`,
      `POST ${BASE}/me/messages/msg-1/reply`,
    ]);
    // One lookup plus the two (zero-delay) retries from the test schedule.
    expect(calls).toHaveLength(5);
    for (const call of calls.slice(2)) {
      expect(call.url).toContain(`${BASE}/me/mailfolders/sentitems/messages?`);
      expect(call.url).toContain("ConversationId+eq+%27conv-1%27");
      expect(call.url).toContain("ReceivedDateTime+ge+");
    }
    expect(calls[1]?.body).toEqual({ Comment: "thanks" });
    const out = payload(result);
    expect(out).toMatchObject({
      ok: true,
      op: "reply",
      recipient_count: 1,
      recipients: { to: ["sender@example.com"] },
    });
    // No unambiguous sent copy: a note, never a guessed id.
    expect(out.id).toBeUndefined();
    expect(String(out.note)).toContain("Sent Items");
  });

  it("posts comment + ToRecipients for forward", async () => {
    const calls = harness([
      json({ Id: "msg-2", ConversationId: "conv-2", Subject: "Agenda" }),
      empty(202),
      json({ value: [] }),
      json({ value: [] }),
    ]);
    const result = await mailTool.handle({
      op: "forward",
      message_id: "msg-2",
      to: ["a@example.com"],
      body: "fyi",
    });
    expect(result.isError).toBeUndefined();
    const forwardCall = calls.find((call) => call.url.endsWith("/me/messages/msg-2/forward"));
    expect(forwardCall?.body).toEqual({
      Comment: "fyi",
      ToRecipients: [{ EmailAddress: { Address: "a@example.com" } }],
    });
    expect(payload(result)).toMatchObject({
      ok: true,
      op: "forward",
      recipient_count: 1,
      recipients: { to: ["a@example.com"] },
    });
  });

  it("requires message_id and body for reply", async () => {
    const calls = harness([empty(202)]);
    const result = await mailTool.handle({ op: "reply", message_id: "msg-1" });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe("mail.mark / mail.move", () => {
  it("patches only the requested read state", async () => {
    const calls = harness([json({ Id: "m1", IsRead: false })]);
    const result = await mailTool.handle({ op: "mark", message_id: "m1", mark: "unread" });
    expect(result.isError).toBeUndefined();
    expect(calls[0]?.method).toBe("PATCH");
    expect(calls[0]?.url).toBe(`${BASE}/me/messages/m1`);
    expect(calls[0]?.body).toEqual({ IsRead: false });
    expect(payload(result)).toMatchObject({ ok: true, mark: "unread", id: "m1" });
  });

  it("patches the flag state", async () => {
    const calls = harness([json({ Id: "m1" })]);
    await mailTool.handle({ op: "mark", message_id: "m1", mark: "flag" });
    expect(calls[0]?.body).toEqual({ Flag: { FlagStatus: "Flagged" } });
  });

  it("moves and returns the NEW message id", async () => {
    const calls = harness([json({ Id: "moved-id" })]);
    const result = await mailTool.handle({
      op: "move",
      message_id: "old-id",
      destination: "archive",
    });
    expect(result.isError).toBeUndefined();
    expect(calls[0]?.url).toBe(`${BASE}/me/messages/old-id/move`);
    expect(calls[0]?.body).toEqual({ DestinationId: "archive" });
    expect(payload(result)).toEqual({
      ok: true,
      op: "move",
      id: "moved-id",
      from_id: "old-id",
      destination: "archive",
    });
  });

  it("surfaces a REST error as a tool error", async () => {
    harness([
      json({ error: { code: "ErrorInvalidIdMalformed", message: "bad id" } }, 400),
    ]);
    const result = await mailTool.handle({ op: "mark", message_id: "bad", mark: "read" });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain(
      "ErrorInvalidIdMalformed",
    );
  });
});

describe("mail.list cursor + op validation", () => {
  it("rejects a malformed cursor instead of silently starting over", async () => {
    const calls = harness([json({ value: [] })]);
    const result = await mailTool.handle({ op: "list", cursor: "not-a-cursor!!" });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain("invalid_cursor");
    expect(calls).toHaveLength(0);
  });

  it("rejects an unknown op with the valid set", async () => {
    const calls = harness([json({ value: [] })]);
    const result = await mailTool.handle({ op: "destroy" });
    expect(result.isError).toBe(true);
    const text = result.content[0]?.type === "text" ? result.content[0].text : "";
    expect(text).toContain("invalid_argument");
    expect(text).toContain("available ops: list");
    expect(calls).toHaveLength(0);
  });
});

describe("mail.list row controls", () => {
  const rowPayload = (overrides: Record<string, unknown> = {}) => ({
    Id: "m1",
    ConversationId: "conv-1",
    Subject: "Newsletter",
    From: { EmailAddress: { Address: "news@example.com" } },
    ReceivedDateTime: "2026-11-16T21:00:00Z",
    IsRead: true,
    BodyPreview: "preview text",
    WebLink: "https://outlook.office365.com/owa/?ItemID=m1",
    InternetMessageHeaders: [
      { Name: "List-Unsubscribe", Value: "<https://lists.example.com/u/1>, <mailto:leave@example.com>" },
      { Name: "List-Unsubscribe-Post", Value: "List-Unsubscribe=One-Click" },
    ],
    ...overrides,
  });

  it("exposes conversation_id and omits web_link/unsub by default", async () => {
    const calls = harness([json({ value: [rowPayload()] })]);
    const out = payload(await mailTool.handle({ op: "list", top: 1 }));
    const row = (out.messages as Array<Record<string, unknown>>)[0];
    expect(row?.conversation_id).toBe("conv-1");
    expect(row?.web_link).toBeUndefined();
    expect(row?.unsub).toBeUndefined();
    expect(calls[0]?.url).not.toContain("InternetMessageHeaders");
  });

  it("include_unsubscribe adds only flags by default, targets on request", async () => {
    const calls = harness([json({ value: [rowPayload()] })]);
    const light = payload(
      await mailTool.handle({ op: "list", top: 1, include_unsubscribe: true, include_web_link: true }),
    );
    const lightRow = (light.messages as Array<Record<string, unknown>>)[0];
    expect(lightRow?.web_link).toContain("outlook.office365.com");
    expect(lightRow?.unsub).toEqual({ has: true, one_click: true });
    expect(calls[0]?.url).toContain("InternetMessageHeaders");

    const verbose = payload(
      await mailTool.handle({
        op: "list",
        top: 1,
        include_unsubscribe: true,
        include_unsubscribe_targets: true,
      }),
    );
    expect((verbose.messages as Array<Record<string, unknown>>)[0]?.unsub).toEqual({
      has: true,
      one_click: true,
      url: "https://lists.example.com/u/1",
      mailto: "mailto:leave@example.com",
    });
  });

  it("omit lists unsub only when include_unsubscribe is set", async () => {
    harness([json({ value: [rowPayload()] })]);
    const out = payload(await mailTool.handle({ op: "list", top: 1, include_unsubscribe_targets: true }));
    expect((out.messages as Array<Record<string, unknown>>)[0]?.unsub).toBeUndefined();
  });

  it("defaults body_chars to 4000", async () => {
    harness([json({ Id: "m1", Subject: "Long", Body: { ContentType: "Text", Content: "x".repeat(6000) } })]);
    const out = payload(await mailTool.handle({ op: "get", message_id: "m1" }));
    expect((out.body as string).length).toBeLessThanOrEqual(4000);
    expect(out.body_truncated).toBe(true);
  });

  it("include_body_preview=false omits preview entirely", async () => {
    harness([json({ value: [rowPayload()] })]);
    const out = payload(await mailTool.handle({ op: "list", top: 1, include_body_preview: false }));
    const row = (out.messages as Array<Record<string, unknown>>)[0];
    expect(row?.preview).toBeUndefined();
  });

  it("body_chars caps a long get body", async () => {
    harness([json({ Id: "m1", Subject: "Long", Body: { ContentType: "Text", Content: "x".repeat(500) } })]);
    const out = payload(await mailTool.handle({ op: "get", message_id: "m1", body_chars: 120 }));
    expect((out.body as string).length).toBeLessThanOrEqual(120);
    expect(out.body_truncated).toBe(true);
  });
});

describe("mail.thread", () => {
  it("REG: does not send $orderby (live InefficientFilter) and sorts locally", async () => {
    const calls = harness([
      json({
        value: [
          { Id: "m2", ConversationId: "conv-1", Subject: "Later", ReceivedDateTime: "2026-11-16T18:00:00Z" },
          { Id: "m1", ConversationId: "conv-1", Subject: "Earlier", ReceivedDateTime: "2026-11-16T16:00:00Z" },
        ],
      }),
    ]);
    const result = await mailTool.handle({ op: "thread", thread_id: "conv-1" });
    expect(result.isError).toBeUndefined();
    expect(decodeURIComponent(calls[0]?.url ?? "")).toContain("$filter=ConversationId");
    expect(calls[0]?.url).not.toContain("%24orderby");
    expect(calls[0]?.url).not.toContain("$orderby");
    const rows = (payload(result).messages as Array<Record<string, unknown>>) ?? [];
    expect(rows.map((row) => row.id)).toEqual(["m1", "m2"]);
    expect(rows[0]?.conversation_id).toBe("conv-1");
  });
});

describe("unsubscribe", () => {
  const headers = [
    { Name: "List-Unsubscribe", Value: "<https://lists.example.com/u/1>, <mailto:leave@example.com>" },
    { Name: "List-Unsubscribe-Post", Value: "List-Unsubscribe=One-Click" },
  ];
  const message = (extra: Record<string, unknown> = {}) =>
    json({
      Id: "bulk-1",
      Subject: "Newsletter",
      From: { EmailAddress: { Address: "news@example.com" } },
      InternetMessageHeaders: headers,
      ...extra,
    });

  // Opt-in is off by default (see test/safety.test.ts); this group covers the
  // enabled contract, including the sender binding.
  beforeEach(() => {
    process.env.TEAMS_MCP_UNSUBSCRIBE_ENABLED = "1";
    resetConfigCache();
  });

  afterEach(() => {
    delete process.env.TEAMS_MCP_UNSUBSCRIBE_ENABLED;
    resetConfigCache();
  });

  it("previews targets without sending anything", async () => {
    const calls = harness([message()]);
    const result = await mailTool.handle({ op: "unsubscribe_preview", message_id: "bulk-1" });
    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
    const out = payload(result);
    expect(out.unsubscribe).toMatchObject({
      url: "https://lists.example.com/u/1",
      mailto: "mailto:leave@example.com",
      oneClick: true,
    });
  });

  it("refuses without confirm and reports the preview", async () => {
    const calls = harness([message()]);
    const result = await mailTool.handle({ op: "unsubscribe_apply", message_id: "bulk-1" });
    const out = payload(result);
    expect(out.refused).toBe(true);
    expect(out.reason).toContain("confirm=true");
    expect(out.preview).toBeDefined();
    expect(calls).toHaveLength(1);
  });

  it("posts the one-click form exactly once and sends the mailto target", async () => {
    setOneClickResolver(async () => ["93.184.216.34"]);
    const posted: Array<{ target: string; body: string }> = [];
    setOneClickPoster(async (target, body) => {
      posted.push({ target, body });
      return { status: 200 };
    });
    const calls = harness([message(), empty(202)]);
    const result = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-1",
      confirm: true,
      expected_sender: "News@Example.com",
    });
    expect(posted).toEqual([{ target: "https://lists.example.com/u/1", body: "List-Unsubscribe=One-Click" }]);
    const sendmail = calls.find((call) => call.url.endsWith("/me/sendmail"));
    expect(sendmail?.body).toMatchObject({
      Message: { Subject: "unsubscribe", ToRecipients: [{ EmailAddress: { Address: "leave@example.com" } }] },
    });
    const out = payload(result);
    expect(out.ok).toBe(true);
    expect(out.results).toEqual([
      { target: "https://lists.example.com/u/1", type: "https", status: "posted", http_status: 200 },
      { target: "leave@example.com", type: "mailto", status: "sent" },
    ]);
  });

  it("never posts to a non-https target", async () => {
    setOneClickPoster(async () => {
      throw new Error("must not be called");
    });
    const calls = harness([
      json({
        Id: "bulk-2",
        Subject: "Old",
        From: { EmailAddress: { Address: "news@example.com" } },
        InternetMessageHeaders: [
          { Name: "List-Unsubscribe", Value: "<http://insecure.example.com/u/1>" },
        ],
      }),
    ]);
    const result = await mailTool.handle({
      op: "unsubscribe_apply",
      message_id: "bulk-2",
      confirm: true,
      expected_sender: "news@example.com",
    });
    const out = payload(result);
    expect(out.ok).toBe(false);
    expect(out.results).toEqual([
      { target: "http://insecure.example.com/u/1", type: "https", status: "refused_insecure_target" },
    ]);
    expect(calls).toHaveLength(1);
  });

  it("reports a message without List-Unsubscribe", async () => {
    harness([json({ Id: "plain-1", Subject: "Plain" })]);
    const result = await mailTool.handle({ op: "unsubscribe_apply", message_id: "plain-1", confirm: true });
    expect(result.isError).toBe(true);
  });
});

describe("reply/forward id lookup", () => {
  const original = (overrides: Record<string, unknown> = {}) =>
    json({
      Id: "msg-9",
      ConversationId: "conv-9",
      Subject: "Quarterly numbers",
      From: { EmailAddress: { Address: "sender@example.com" } },
      ToRecipients: [{ EmailAddress: { Address: "me@example.com" } }],
      ...overrides,
    });
  const sentRows = (...rows: Array<Record<string, unknown>>) => json({ value: rows });

  it("returns the sent id when exactly one copy matches", async () => {
    harness([
      original(),
      empty(202),
      sentRows({ Id: "sent-1", Subject: "RE: Quarterly numbers", ReceivedDateTime: "2026-11-16T12:00:00Z" }),
    ]);
    const result = await mailTool.handle({ op: "reply", message_id: "msg-9", body: "thanks" });
    expect(result.isError).toBeUndefined();
    expect(payload(result)).toEqual({
      ok: true,
      op: "reply",
      recipient_count: 1,
      recipients: { to: ["sender@example.com"] },
      id: "sent-1",
    });
  });

  it("refuses to guess when two copies match", async () => {
    harness([
      original(),
      empty(202),
      sentRows(
        { Id: "sent-1", Subject: "RE: Quarterly numbers" },
        { Id: "sent-2", Subject: "RE:  Quarterly   numbers" },
      ),
    ]);
    const result = await mailTool.handle({ op: "reply", message_id: "msg-9", body: "thanks" });
    const out = payload(result);
    expect(out.id).toBeUndefined();
    expect(String(out.note)).toContain("2 copies");
  });

  it("never picks a different subject in the same conversation", async () => {
    harness([
      original(),
      empty(202),
      sentRows({ Id: "sent-other", Subject: "RE: Something else" }),
      sentRows({ Id: "sent-other", Subject: "RE: Something else" }),
    ]);
    const result = await mailTool.handle({ op: "reply", message_id: "msg-9", body: "thanks" });
    const out = payload(result);
    expect(out.id).toBeUndefined();
    expect(String(out.note)).toContain("Sent Items");
  });

  it("still finds the copy when the provider collapses a repeated prefix", async () => {
    harness([
      original({ Subject: "RE: Quarterly numbers" }),
      empty(202),
      sentRows({ Id: "sent-collapsed", Subject: "RE: Quarterly numbers" }),
    ]);
    const result = await mailTool.handle({ op: "reply", message_id: "msg-9", body: "again" });
    expect(payload(result)).toMatchObject({ id: "sent-collapsed" });
  });

  it("returns the forwarded copy's id", async () => {
    harness([
      json({ Id: "msg-9", ConversationId: "conv-9", Subject: "Quarterly numbers" }),
      empty(202),
      sentRows({ Id: "sent-fw", Subject: "FW: Quarterly numbers" }),
    ]);
    const result = await mailTool.handle({
      op: "forward",
      message_id: "msg-9",
      to: ["a@example.com"],
      body: "fyi",
    });
    expect(payload(result)).toMatchObject({ ok: true, op: "forward", recipient_count: 1, id: "sent-fw" });
  });
});

describe("mail.send sent-id lookup", () => {
  const sentRows = (...rows: Array<Record<string, unknown>>) => json({ value: rows });

  it("returns the id of the single subject match", async () => {
    const calls = harness([
      empty(202),
      sentRows({ Id: "sent-42", Subject: "Weekly digest", ReceivedDateTime: "2026-11-17T12:00:00Z" }),
    ]);
    const out = payload(
      await mailTool.handle({ op: "send", to: ["a@example.com"], subject: "Weekly digest", body: "x" }),
    );
    expect(out).toMatchObject({ ok: true, op: "send", recipient_count: 1, id: "sent-42" });
    expect(out.note).toBeUndefined();
    // A plain send has no conversation yet: the lookup is subject + time only.
    const lookup = new URL(calls[1]?.url ?? "");
    expect(lookup.pathname).toContain("/me/mailfolders/sentitems/messages");
    expect(lookup.searchParams.get("$filter")).toContain("ReceivedDateTime ge");
    expect(lookup.searchParams.get("$filter")).not.toContain("ConversationId");
  });

  it("returns a note, not an id, when two copies match", async () => {
    harness([
      empty(202),
      sentRows(
        { Id: "sent-a", Subject: "Weekly digest" },
        { Id: "sent-b", Subject: "weekly   DIGEST" },
      ),
    ]);
    const out = payload(
      await mailTool.handle({ op: "send", to: ["a@example.com"], subject: "Weekly digest", body: "x" }),
    );
    expect(out.id).toBeUndefined();
    expect(String(out.note)).toContain("2 copies");
  });

  it("returns a note when the copy is not indexed yet", async () => {
    harness([empty(202), sentRows()]);
    const out = payload(
      await mailTool.handle({ op: "send", to: ["a@example.com"], subject: "Weekly digest", body: "x" }),
    );
    expect(out.id).toBeUndefined();
    expect(String(out.note)).toContain("Sent Items");
  });
});

describe("mail.draft", () => {
  it("posts a REST Message to /me/messages and returns its id", async () => {
    const calls = harness([json({ Id: "draft-1", WebLink: "https://outlook.office365.com/owa/?ItemID=draft-1" })]);
    const out = payload(
      await mailTool.handle({
        op: "draft",
        to: ["self@example.com"],
        cc: ["cc@example.com"],
        subject: "AI TEST — draft",
        body: "hello",
      }),
    );
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("POST");
    expect(calls[0]?.url).toBe(`${BASE}/me/messages`);
    expect(calls[0]?.body).toEqual({
      Subject: "AI TEST — draft",
      Body: { ContentType: "Text", Content: "hello" },
      ToRecipients: [{ EmailAddress: { Address: "self@example.com" } }],
      CcRecipients: [{ EmailAddress: { Address: "cc@example.com" } }],
    });
    expect(out).toMatchObject({
      ok: true,
      op: "draft",
      recipient_count: 2,
      id: "draft-1",
      web_link: "https://outlook.office365.com/owa/?ItemID=draft-1",
    });
  });

  it("requires to[], subject and body without a request", async () => {
    const calls = harness([json({ Id: "draft-1" })]);
    const result = await mailTool.handle({ op: "draft", to: ["a@example.com"], body: "x" });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("enforces invite_allowlist before creating anything", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@example.com"]\n');
    resetConfigCache();
    const calls = harness([json({ Id: "draft-1" })]);
    const result = await mailTool.handle({
      op: "draft",
      to: ["outsider@other.test"],
      subject: "x",
      body: "y",
    });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("is refused offline in read-only mode", async () => {
    process.env.TEAMS_MCP_READ_ONLY = "1";
    resetConfigCache();
    const calls = harness([json({ Id: "draft-1" })]);
    const result = await mailTool.handle({
      op: "draft",
      to: ["a@example.com"],
      subject: "x",
      body: "y",
    });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain("read_only");
    expect(calls).toHaveLength(0);
  });

  it("spends one rate token and journals the write", async () => {
    process.env.TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE = "1";
    resetConfigCache();
    resetWriteLimiter();
    const calls = harness([json({ Id: "draft-1" })]);
    const first = await mailTool.handle({
      op: "draft",
      to: ["a@example.com"],
      subject: "x",
      body: "y",
    });
    expect(first.isError).toBeUndefined();
    const second = await mailTool.handle({
      op: "draft",
      to: ["a@example.com"],
      subject: "x",
      body: "y",
    });
    expect(second.isError).toBe(true);
    expect(second.content[0]?.type === "text" ? second.content[0].text : "").toContain("rate_limited");
    expect(calls).toHaveLength(1);
    const audit = await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8");
    expect(audit).toContain('"op":"draft"');
    expect(audit).toContain('"outcome":"ok"');
  });
});

describe("mail.send_draft", () => {
  const draftRow = (overrides: Record<string, unknown> = {}) => ({
    Id: "draft-1",
    ConversationId: "conv-1",
    Subject: "AI TEST — draft",
    ToRecipients: [{ EmailAddress: { Address: "a@example.com" } }],
    CcRecipients: [{ EmailAddress: { Address: "cc@example.com" } }],
    ...overrides,
  });

  it("re-reads the draft, then posts the send action", async () => {
    const calls = harness([
      json(draftRow()),
      empty(202),
      json({ value: [{ Id: "sent-7", Subject: "AI TEST — draft" }] }),
    ]);
    const out = payload(await mailTool.handle({ op: "send_draft", draft_id: "draft-1", confirm: true }));
    expect(out).toMatchObject({
      ok: true,
      op: "send_draft",
      draft_id: "draft-1",
      recipient_count: 2,
      recipients: { to: ["a@example.com"], cc: ["cc@example.com"] },
      id: "sent-7",
    });
    expect(decodeURIComponent(calls[0]?.url ?? "")).toContain(
      "$select=Id,ConversationId,Subject,ToRecipients,CcRecipients,BccRecipients",
    );
    expect(calls[1]?.method).toBe("POST");
    expect(calls[1]?.url).toBe(`${BASE}/me/messages/draft-1/send`);
  });

  it("requires confirm=true and sends nothing otherwise", async () => {
    const calls = harness([json(draftRow())]);
    const result = await mailTool.handle({ op: "send_draft", draft_id: "draft-1" });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain(
      "confirmation_required",
    );
    expect(calls).toHaveLength(0);
  });

  it("re-applies invite_allowlist to the draft as it is now", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@example.com"]\n');
    resetConfigCache();
    const calls = harness([
      json(
        draftRow({
          CcRecipients: [{ EmailAddress: { Address: "outsider@other.test" } }],
        }),
      ),
    ]);
    const result = await mailTool.handle({ op: "send_draft", draft_id: "draft-1", confirm: true });
    expect(result.isError).toBe(true);
    const text = result.content[0]?.type === "text" ? result.content[0].text : "";
    expect(text).toContain("recipient_not_allowed");
    expect(text).toContain("outsider@other.test");
    expect(calls.map((call) => call.url)).toEqual([
      expect.stringContaining(`${BASE}/me/messages/draft-1?`),
    ]);
  });

  it("refuses a draft with no recipients", async () => {
    const calls = harness([json(draftRow({ ToRecipients: [], CcRecipients: [] }))]);
    const result = await mailTool.handle({ op: "send_draft", draft_id: "draft-1", confirm: true });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(1);
  });

  it("is refused in read-only mode with no request at all", async () => {
    process.env.TEAMS_MCP_READ_ONLY = "1";
    resetConfigCache();
    const calls = harness([json(draftRow())]);
    const result = await mailTool.handle({ op: "send_draft", draft_id: "draft-1", confirm: true });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe("mail.folders", () => {
  it("reads one page of folders with counts", async () => {
    const calls = harness([
      json({
        value: [
          { Id: "inbox", DisplayName: "Inbox", TotalItemCount: 42, UnreadItemCount: 7, ChildFolderCount: 2 },
          { Id: "sentitems", DisplayName: "Sent Items", TotalItemCount: 9, UnreadItemCount: 0, ChildFolderCount: 0 },
        ],
      }),
    ]);
    const out = payload(await mailTool.handle({ op: "folders" }));
    expect(calls[0]?.method).toBe("GET");
    const query = new URL(calls[0]?.url ?? "").searchParams;
    expect(query.get("$top")).toBe("50");
    expect(query.get("$select")).toBe(
      "Id,DisplayName,TotalItemCount,UnreadItemCount,ChildFolderCount",
    );
    expect(out).toEqual({
      count: 2,
      truncated: false,
      folders: [
        { id: "inbox", name: "Inbox", total: 42, unread: 7, child_folders: 2 },
        { id: "sentitems", name: "Sent Items", total: 9, unread: 0, child_folders: 0 },
      ],
    });
  });

  it("flags a full page as truncated", async () => {
    harness([
      json({
        value: Array.from({ length: 50 }, (_, index) => ({
          Id: `f${index}`,
          DisplayName: `Folder ${index}`,
          TotalItemCount: 0,
          UnreadItemCount: 0,
          ChildFolderCount: 0,
        })),
      }),
    ]);
    const out = payload(await mailTool.handle({ op: "folders" }));
    expect(out.count).toBe(50);
    expect(out.truncated).toBe(true);
  });
});

describe("mail.reply reply_all", () => {
  const original = (overrides: Record<string, unknown> = {}) =>
    json({
      Id: "msg-1",
      ConversationId: "conv-1",
      Subject: "Planning",
      From: { EmailAddress: { Address: "sender@example.com" } },
      ToRecipients: [{ EmailAddress: { Address: "me@example.com" } }],
      CcRecipients: [{ EmailAddress: { Address: "cc@example.com" } }],
      ...overrides,
    });

  it("uses the replyall action and reports the derived recipient count", async () => {
    const calls = harness([
      original(),
      empty(202),
      json({ value: [{ Id: "sent-ra", Subject: "RE: Planning" }] }),
    ]);
    const out = payload(
      await mailTool.handle({ op: "reply", message_id: "msg-1", body: "ok", reply_all: true }),
    );
    expect(out).toMatchObject({
      ok: true,
      op: "reply",
      reply_all: true,
      recipient_count: 3,
      recipients: { to: ["sender@example.com", "me@example.com"], cc: ["cc@example.com"] },
      id: "sent-ra",
    });
    const send = calls.find((call) => call.url.endsWith("/replyall"));
    expect(send?.body).toEqual({ Comment: "ok" });
    expect(calls.some((call) => call.url.endsWith("/reply"))).toBe(false);
  });

  it("refuses before sending when a derived recipient is outside the allowlist", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@mycorp.test"]\n');
    resetConfigCache();
    const calls = harness([
      original({
        From: { EmailAddress: { Address: "sender@mycorp.test" } },
        ToRecipients: [{ EmailAddress: { Address: "me@mycorp.test" } }],
        CcRecipients: [{ EmailAddress: { Address: "outsider@other.test" } }],
      }),
    ]);
    const result = await mailTool.handle({
      op: "reply",
      message_id: "msg-1",
      body: "ok",
      reply_all: true,
    });
    expect(result.isError).toBe(true);
    const text = result.content[0]?.type === "text" ? result.content[0].text : "";
    expect(text).toContain("recipient_not_allowed");
    expect(text).toContain("outsider@other.test");
    expect(calls.some((call) => call.method === "POST")).toBe(false);
  });

  it("allows reply_all when every derived recipient is allowlisted", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@mycorp.test"]\n');
    resetConfigCache();
    const calls = harness([
      original({
        From: { EmailAddress: { Address: "sender@mycorp.test" } },
        ToRecipients: [{ EmailAddress: { Address: "me@mycorp.test" } }],
        CcRecipients: [{ EmailAddress: { Address: "cc@mycorp.test" } }],
      }),
      empty(202),
      json({ value: [] }),
    ]);
    const out = payload(
      await mailTool.handle({ op: "reply", message_id: "msg-1", body: "ok", reply_all: true }),
    );
    expect(out.reply_all).toBe(true);
    expect(calls.some((call) => call.url.endsWith("/replyall"))).toBe(true);
  });
});
