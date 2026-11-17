/**
 * `teams` tool: op dispatch, token routing, write gates and error mapping.
 *
 * Every case injects both Graph clients (`setGraphClient`, `setTeamsWebClient`)
 * so the assertions are about the wire the tool would have produced — a refused
 * call is asserted as `calls.length === 0`, not merely as an error result.
 */
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { GraphClient, setGraphClient } from "../src/graph/client.js";
import { TeamsWebClient, setTeamsWebClient } from "../src/teams/client.js";
import { recordCreatedTarget, readTeamsState } from "../src/teams/state.js";
import { setSetting } from "../src/settings.js";
import { resetWriteLimiter } from "../src/safety.js";
import { writeTokenCache } from "../src/auth/store.js";
import { writeTeamsWebCache } from "../src/auth/teams-web.js";
import { teamsTool } from "../src/tools/teams.js";
import type { FetchLike } from "../src/http/client.js";

interface Call {
  method: string;
  url: string;
  body: unknown;
}

function harness(responses: Array<() => Response>, token: string) {
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
    if (!factory) throw new Error(`no response queued for ${url}`);
    return factory();
  };
  return { calls, fetchImpl, token };
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status });

function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

function payload(result: CallToolResult): Record<string, unknown> {
  return JSON.parse(textOf(result)) as Record<string, unknown>;
}

const empty = (status: number): (() => Response) => () => new Response(null, { status });

const readState = () => readTeamsState();

const saved = { ...process.env };
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-tool-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  delete process.env.TEAMS_MCP_READ_ONLY;
  resetConfigCache();
  resetWriteLimiter();
});

afterEach(async () => {
  setGraphClient(undefined);
  setTeamsWebClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  resetWriteLimiter();
  await rm(dir, { recursive: true, force: true });
});

function owaGraph(responses: Array<() => Response>): Call[] {
  const h = harness(responses, "owa-token");
  setGraphClient(
    new GraphClient({ tokenProvider: async () => ({ token: h.token }), fetchImpl: h.fetchImpl, maxAttempts: 1 }),
  );
  return h.calls;
}

function webGraph(responses: Array<() => Response>): Call[] {
  const h = harness(responses, "web-token");
  setTeamsWebClient(
    new TeamsWebClient({ tokenProvider: async () => ({ token: h.token }), fetchImpl: h.fetchImpl, maxAttempts: 1 }),
  );
  return h.calls;
}

/** Seed the OWA cache so `selfUserId()` can answer without a network call. */
async function seedSelf(id = "self-id") {
  await writeTokenCache({
    version: 1,
    account: { upn: "user@example.com", id },
    tokens: {
      outlook: { access_token: "cache-token", expires_at: new Date(Date.now() + 3_600_000).toISOString() },
    },
    source: "refresh",
    updated_at: new Date().toISOString(),
  });
}

describe("reads", () => {
  it("lists chats with a minimal select and an opaque cursor", async () => {
    const calls = owaGraph([
      json({
        value: [{ id: "19:chat", topic: "Team 2A", chatType: "group", lastUpdatedDateTime: "2026-11-16T10:00:00Z" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/chats?$top=25&$skiptoken=abc",
      }),
    ]);
    const result = await teamsTool.handle({ op: "list", kind: "chats" });
    expect(result.isError).toBeUndefined();
    const body = payload(result);
    expect(body.count).toBe(1);
    expect((body.chats as Array<Record<string, unknown>>)[0]).toEqual({
      id: "19:chat",
      name: "Team 2A",
      chat_type: "group",
      last_updated: "2026-11-16T10:00:00Z",
    });
    expect(typeof body.next_cursor).toBe("string");
    expect(calls[0]?.url).toContain("$select=id,topic,chatType,lastUpdatedDateTime");
  });

  it("includes the last message preview when Graph exposes one", async () => {
    owaGraph([
      json({
        value: [
          {
            id: "19:preview",
            topic: "AI TEST — teams-mcp — automated test",
            chatType: "group",
            lastUpdatedDateTime: "2026-11-16T10:00:00Z",
            lastMessagePreview: {
              id: "1730000000456",
              createdDateTime: "2026-11-16T09:59:00Z",
              body: { contentType: "html", content: "<p>hello <b>team</b></p>" },
            },
          },
        ],
      }),
    ]);
    const row = (payload(await teamsTool.handle({ op: "list", kind: "chats" })).chats as Array<
      Record<string, unknown>
    >)[0];
    expect(row).toMatchObject({
      last_message_preview: "hello team",
      last_message_at: "2026-11-16T09:59:00Z",
    });
  });

  it("lists channels only with a team id", async () => {
    const calls = owaGraph([json({ value: [{ id: "19:ch", displayName: "General", membershipType: "standard" }] })]);
    const missing = await teamsTool.handle({ op: "list", kind: "channels" });
    expect(missing.isError).toBe(true);
    expect(textOf(missing)).toContain("team_id");
    expect(calls).toHaveLength(0);

    const result = await teamsTool.handle({ op: "list", kind: "channels", team_id: "t1" });
    expect(result.isError).toBeUndefined();
    expect(payload(result).channels).toEqual([{ id: "19:ch", name: "General", membership: "standard" }]);
    expect(calls[0]?.url).toContain("/teams/t1/channels");
  });

  it("queries teams by name across pages, case-insensitively", async () => {
    const calls = owaGraph([
      json({
        value: [
          { id: "t1", displayName: "Team Alpha" },
          { id: "t2", displayName: "Team Beta" },
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/joinedTeams?$skiptoken=a",
      }),
      json({ value: [{ id: "t3", displayName: "Robotics" }] }),
    ]);
    const body = payload(await teamsTool.handle({ op: "list", kind: "teams", query: "robotics" }));
    expect(body.count).toBe(1);
    expect((body.teams as Array<Record<string, unknown>>)[0]).toMatchObject({ id: "t3", name: "Robotics" });
    expect(body.has_more).toBe(false);
    expect(body.truncated).toBe(false);
    expect(body.scanned_pages).toBe(2);
    expect(calls).toHaveLength(2);
  });

  it("queries one team's channels by name, case-insensitively", async () => {
    const channels = {
      value: [
        { id: "c1", displayName: "General 2026", membershipType: "standard" },
        { id: "c2", displayName: "General 2025", membershipType: "standard" },
        { id: "c3", displayName: "Robotics Announcements", membershipType: "standard" },
      ],
    };
    owaGraph([json(channels), json(channels)]);
    const body = payload(
      await teamsTool.handle({ op: "list", kind: "channels", team_id: "t1", query: "general 2025" }),
    );
    expect(body.count).toBe(1);
    expect((body.channels as Array<Record<string, unknown>>)[0]).toMatchObject({
      id: "c2",
      name: "General 2025",
    });
    expect(body.truncated).toBe(false);

    const none = payload(
      await teamsTool.handle({ op: "list", kind: "channels", team_id: "t1", query: "does-not-exist" }),
    );
    expect(none.count).toBe(0);
    expect(none.has_more).toBe(false);
    expect(none.truncated).toBe(false);
  });

  it("reads channel messages with the Teams-web token, strips system rows and converts HTML", async () => {
    const web = webGraph([
      json({
        value: [
          {
            id: "m1",
            messageType: "message",
            createdDateTime: "2026-11-16T10:00:00Z",
            body: { contentType: "html", content: "<p>Hello <b>team</b></p>" },
            from: { user: { displayName: "Joel" } },
            webUrl: "https://teams.microsoft.com/l/message/1",
          },
          {
            id: "m2",
            messageType: "unknownFutureValue",
            body: { contentType: "html", content: "<systemEventMessage/>" },
          },
        ],
      }),
    ]);
    const owa = owaGraph([]);
    const result = await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" });
    expect(result.isError).toBeUndefined();
    const body = payload(result);
    expect(body.count).toBe(1);
    expect((body.messages as Array<Record<string, unknown>>)[0]).toMatchObject({
      id: "m1",
      from: "Joel",
      text: "Hello team",
    });
    expect(body.truncated).toBe(false);
    expect(web).toHaveLength(1);
    expect(owa).toHaveLength(0);
  });

  it("keeps paging past system rows until it has top real messages", async () => {
    const web = webGraph([
      json({
        value: [
          { id: "s1", messageType: "unknownFutureValue", body: { content: "<systemEventMessage/>" } },
          { id: "m1", messageType: "message", body: { content: "one" } },
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages?$top=2&$skiptoken=a",
      }),
      json({ value: [{ id: "m2", messageType: "message", body: { content: "two" } }] }),
    ]);
    const result = await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1", top: 2 });
    const body = payload(result);
    expect(body.count).toBe(2);
    expect((body.messages as Array<Record<string, unknown>>).map((row) => row.id)).toEqual(["m1", "m2"]);
    expect(body.truncated).toBe(false);
    expect(web).toHaveLength(2);
  });

  it("renders system events only when include_system is true", async () => {
    webGraph([
      json({
        value: [
          {
            id: "s1",
            messageType: "unknownFutureValue",
            body: { contentType: "html", content: "<systemEventMessage/>" },
            eventDetail: {
              "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
              members: [{ id: "aad-1", displayName: "Test Student" }],
            },
          },
        ],
      }),
    ]);
    const result = await teamsTool.handle({
      op: "read",
      team_id: "t1",
      channel_id: "c1",
      include_system: true,
    });
    const row = (payload(result).messages as Array<Record<string, unknown>>)[0];
    expect(row).toMatchObject({ id: "s1", system: true, event: "member added" });
    expect(String(row?.text)).toContain("Test Student");
  });

  it("filters chats by chat_type server-side across pages", async () => {
    const calls = owaGraph([
      json({
        value: [{ id: "19:a", topic: "A", chatType: "group" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/chats?$top=25&$filter=chatType%20eq%20'group'&$skiptoken=a",
      }),
      json({ value: [{ id: "19:b", topic: "B", chatType: "group" }] }),
    ]);
    const result = await teamsTool.handle({ op: "list", kind: "chats", chat_type: "group", top: 25 });
    const body = payload(result);
    expect(body.count).toBe(2);
    expect((body.chats as Array<Record<string, unknown>>).map((row) => row.id)).toEqual(["19:a", "19:b"]);
    expect(body.truncated).toBe(false);
    expect(calls[0]?.url).toContain("$filter=chatType eq 'group'");
    expect(calls).toHaveLength(2);
  });

  it("queries chats by topic or member name across pages", async () => {
    const calls = owaGraph([
      json({
        value: [
          { id: "19:a", topic: "March planning sync", chatType: "group", members: [] },
          {
            id: "19:b",
            chatType: "oneOnOne",
            members: [{ displayName: "Test Student", email: "joel@example.com" }],
          },
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/chats?$top=5&$expand=members&$skiptoken=a",
      }),
      json({ value: [{ id: "19:c", topic: "Nothing relevant", members: [{ displayName: "Someone" }] }] }),
    ]);
    const result = await teamsTool.handle({ op: "list", kind: "chats", query: "student", top: 5 });
    const body = payload(result);
    expect(body.count).toBe(1);
    expect((body.chats as Array<Record<string, unknown>>)[0]).toMatchObject({
      id: "19:b",
      members: ["Test Student"],
    });
    expect(body.truncated).toBe(false);
    expect(calls[0]?.url).toContain("$expand=members");
    expect(calls).toHaveLength(2);
  });

  it("maps 404/410 to a not_found error", async () => {
    webGraph([json({ error: { code: "Gone", message: "UnknownError" } }, 410)]);
    const result = await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c-gone" });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("not_found");
  });

  it("maps a Graph 403 to missing_scope naming the permission", async () => {
    webGraph([
      json(
        {
          error: {
            code: "Forbidden",
            message:
              "Missing scope permissions on the request. API requires one of 'ChannelMessage.Read.All'. Scopes on the request 'Chat.Read'",
          },
        },
        403,
      ),
    ]);
    const result = await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("missing_scope");
    expect(textOf(result)).toContain("ChannelMessage.Read.All");
  });

  it("keeps a membership-bound 403 as forbidden, not missing_scope", async () => {
    webGraph([json({ error: { code: "Forbidden", message: "UnknownError" } }, 403)]);
    const result = await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c-private" });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("forbidden");
    expect(textOf(result)).toContain("private channel");
  });

  it("follows an empty first page on channel messages", async () => {
    const web = webGraph([
      json({
        value: [],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages?$top=1&$skiptoken=a",
      }),
      json({ value: [{ id: "m1", body: { contentType: "text", content: "first" } }] }),
    ]);
    const result = await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1", top: 1 });
    expect(result.isError).toBeUndefined();
    expect(payload(result).count).toBe(1);
    expect(web).toHaveLength(2);
  });

  it("reads replies on the Teams-web token", async () => {
    const web = webGraph([json({ value: [{ id: "r1", body: { contentType: "text", content: "reply" } }] })]);
    const result = await teamsTool.handle({
      op: "replies",
      team_id: "t1",
      channel_id: "c1",
      message_id: "m1",
    });
    expect(result.isError).toBeUndefined();
    expect(payload(result).message_id).toBe("m1");
    expect(web[0]?.url).toContain("/messages/m1/replies");
  });

  it("rejects a cursor minted for another host", async () => {
    owaGraph([]);
    const cursor = Buffer.from("https://evil.example.com/steal", "utf8").toString("base64url");
    const result = await teamsTool.handle({ op: "list", kind: "chats", cursor });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("invalid_cursor");
  });
});

describe("write gates", () => {
  it("requires confirm before any request", async () => {
    const calls = owaGraph([]);
    const result = await teamsTool.handle({ op: "send", chat_id: "19:chat", text: "hi" });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("confirmation_required");
    expect(calls).toHaveLength(0);
  });

  it("refuses a pre-existing target unless the operator allowlists it", async () => {
    const calls = owaGraph([json({ id: "m-new" })]);
    const refused = await teamsTool.handle({
      op: "send",
      chat_id: "19:existing",
      text: "hi",
      confirm: true,
    });
    expect(refused.isError).toBe(true);
    expect(textOf(refused)).toContain("target_not_allowed");
    expect(textOf(refused)).not.toContain("allow_existing");
    expect(calls).toHaveLength(0);

    // Only an operator setting (config.yaml/cli) opens the target.
    await setSetting("teams.allowed_targets", ["chat:19:existing"]);
    resetConfigCache();

    const sent = await teamsTool.handle({
      op: "send",
      chat_id: "19:existing",
      text: "hi",
      confirm: true,
    });
    expect(sent.isError).toBeUndefined();
    expect(payload(sent).message_id).toBe("m-new");
    expect(calls[0]?.method).toBe("POST");
    expect(calls[0]?.url).toContain("/chats/19%3Aexisting/messages");
    expect(calls[0]?.body).toEqual({ body: { contentType: "text", content: "hi" } });
  });

  it("scopes a channel allowlist entry to its team", async () => {
    await setSetting("teams.allowed_targets", ["channel:t1/19:shared"]);
    resetConfigCache();
    const web = webGraph([json({ id: "m1" })]);
    const refused = await teamsTool.handle({
      op: "send",
      team_id: "t2",
      channel_id: "19:shared",
      text: "hi",
      confirm: true,
    });
    expect(textOf(refused)).toContain("target_not_allowed");
    expect(web).toHaveLength(0);

    const sent = await teamsTool.handle({
      op: "send",
      team_id: "t1",
      channel_id: "19:shared",
      text: "hi",
      confirm: true,
    });
    expect(sent.isError).toBeUndefined();
    expect(web).toHaveLength(1);
  });

  it("journals confirmation and target policy refusals", async () => {
    const calls = owaGraph([json({ id: "m1" })]);
    await teamsTool.handle({ op: "send", chat_id: "19:existing", text: "x" });
    await teamsTool.handle({ op: "send", chat_id: "19:existing", text: "x", confirm: true });
    expect(calls).toHaveLength(0);
    const audit = await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8");
    expect(audit).toContain('"outcome":"confirmation_required"');
    expect(audit).toContain('"outcome":"target_not_allowed"');
  });

  it("sends to a target this server created with no config entry", async () => {
    await recordCreatedTarget({
      surface: "chat",
      id: "19:ours",
      name: "test",
      created_at: new Date().toISOString(),
    });
    const calls = owaGraph([json({ id: "m1" })]);
    const result = await teamsTool.handle({ op: "send", chat_id: "19:ours", text: "hello", confirm: true });
    expect(result.isError).toBeUndefined();
    expect(payload(result)).toMatchObject({ ok: true, surface: "chat", target_id: "19:ours", message_id: "m1" });
    expect(calls).toHaveLength(1);
  });

  it("sends to a channel with the Teams-web token", async () => {
    await recordCreatedTarget({
      surface: "channel",
      id: "19:ch",
      team_id: "t1",
      name: "test",
      created_at: new Date().toISOString(),
    });
    const web = webGraph([json({ id: "m2" })]);
    const result = await teamsTool.handle({
      op: "send",
      team_id: "t1",
      channel_id: "19:ch",
      text: "hello",
      confirm: true,
    });
    expect(result.isError).toBeUndefined();
    expect(payload(result).surface).toBe("channel");
    expect(web[0]?.url).toContain("/teams/t1/channels/19%3Ach/messages");
  });

  it("enforces the shared write rate limit before a second request", async () => {
    await recordCreatedTarget({
      surface: "chat",
      id: "19:ours",
      created_at: new Date().toISOString(),
    });
    process.env.TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE = "1";
    resetConfigCache();
    resetWriteLimiter();
    const calls = owaGraph([json({ id: "m1" })]);
    const first = await teamsTool.handle({ op: "send", chat_id: "19:ours", text: "one", confirm: true });
    const second = await teamsTool.handle({ op: "send", chat_id: "19:ours", text: "two", confirm: true });
    expect(first.isError).toBeUndefined();
    expect(second.isError).toBe(true);
    expect(textOf(second)).toContain("rate_limited");
    expect(calls).toHaveLength(1);
    const audit = await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8");
    expect(audit).toContain('"outcome":"rate_limited"');
  });

  it("rejects text over the 20000-character cap before any request or token spend", async () => {
    await recordCreatedTarget({
      surface: "chat",
      id: "19:ours",
      created_at: new Date().toISOString(),
    });
    process.env.TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE = "1";
    resetConfigCache();
    resetWriteLimiter();
    const calls = owaGraph([json({ id: "m1" })]);

    const tooLong = await teamsTool.handle({
      op: "send",
      chat_id: "19:ours",
      text: "x".repeat(20_001),
      confirm: true,
    });
    expect(tooLong.isError).toBe(true);
    expect(textOf(tooLong)).toContain("invalid_argument");
    expect(textOf(tooLong)).toContain("20000");
    expect(calls).toHaveLength(0);
    const audit = await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8");
    expect(audit).toContain('"outcome":"invalid_argument"');

    // The refusal must not consume the single rate token.
    const atLimit = await teamsTool.handle({
      op: "send",
      chat_id: "19:ours",
      text: "x".repeat(20_000),
      confirm: true,
    });
    expect(atLimit.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
  });

  it("audits the send and consumes a rate-limit token", async () => {
    await recordCreatedTarget({
      surface: "chat",
      id: "19:ours",
      created_at: new Date().toISOString(),
    });
    owaGraph([json({ id: "m1" })]);
    await teamsTool.handle({ op: "send", chat_id: "19:ours", text: "hello", confirm: true });
    const audit = await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8");
    expect(audit).toContain('"tool":"teams"');
    expect(audit).toContain('"outcome":"ok"');
  });
});

describe("delete", () => {
  async function recordChannel(id = "19:ours", teamId = "t1", name = "AI TEST — teams-mcp") {
    await recordCreatedTarget({
      surface: "channel",
      id,
      team_id: teamId,
      name,
      created_at: new Date().toISOString(),
    });
  }

  it("requires confirm and journals the refusal", async () => {
    await recordChannel();
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "delete",
      team_id: "t1",
      channel_id: "19:ours",
      expected_name: "AI TEST — teams-mcp",
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("confirmation_required");
    expect(calls).toHaveLength(0);
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"confirmation_required"',
    );
  });

  it("refuses a channel this server did not create", async () => {
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "delete",
      team_id: "t1",
      channel_id: "19:existing",
      expected_name: "x",
      confirm: true,
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("target_not_created");
    expect(calls).toHaveLength(0);
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"target_not_created"',
    );
  });

  it("compares expected_name before deleting", async () => {
    await recordChannel();
    const calls = owaGraph([json({ id: "19:ours", displayName: "Something else", membershipType: "private" })]);
    const result = await teamsTool.handle({
      op: "delete",
      team_id: "t1",
      channel_id: "19:ours",
      expected_name: "AI TEST — teams-mcp",
      confirm: true,
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("name_mismatch");
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(await readState()).toHaveLength(1);
  });

  it("deletes an owned channel and prunes the ledger", async () => {
    await recordChannel();
    const calls = owaGraph([
      json({ id: "19:ours", displayName: "AI TEST — teams-mcp", membershipType: "private" }),
      empty(204),
    ]);
    const result = await teamsTool.handle({
      op: "delete",
      team_id: "t1",
      channel_id: "19:ours",
      expected_name: "AI TEST — teams-mcp",
      confirm: true,
    });
    expect(result.isError).toBeUndefined();
    expect(payload(result)).toMatchObject({ deleted: true, team_id: "t1", channel_id: "19:ours" });
    expect(calls.map((call) => call.method)).toEqual(["GET", "DELETE"]);
    expect(await readState()).toHaveLength(0);
  });

  it("treats an already-absent channel as success and prunes the ledger", async () => {
    await recordChannel();
    owaGraph([json({ error: { code: "NotFound", message: "gone" } }, 410)]);
    const result = await teamsTool.handle({
      op: "delete",
      team_id: "t1",
      channel_id: "19:ours",
      expected_name: "AI TEST — teams-mcp",
      confirm: true,
    });
    expect(result.isError).toBeUndefined();
    expect(payload(result)).toMatchObject({ deleted: false, already_absent: true });
    expect(await readState()).toHaveLength(0);
    // The ledger prune is a state mutation; it must leave a journal line.
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"already_absent"',
    );
  });
});

describe("create", () => {
  it("refuses chat creation instead of silently granting every member ownership", async () => {
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "chat",
      topic: "AI TEST — teams-mcp",
      members: ["joel@example.com"],
      confirm: true,
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("invalid_argument");
    expect(textOf(result)).toContain("owner or guest");
    expect(textOf(result)).toContain("Teams");
    expect(calls).toHaveLength(0);
    expect(await readTeamsState()).toEqual([]);
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"invalid_argument"',
    );
  });

  it("requires confirm before creating a channel", async () => {
    const calls = owaGraph([]);
    const noConfirm = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "AI TEST — teams-mcp",
      private: true,
    });
    expect(textOf(noConfirm)).toContain("confirmation_required");
    expect(calls).toHaveLength(0);
  });

  it("refuses a channel unless the caller is a team owner", async () => {
    await seedSelf("self-id");
    const calls = owaGraph([json({ value: [{ id: "someone-else" }] })]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "AI TEST — teams-mcp",
      private: true,
      confirm: true,
    });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("not_owner");
    expect(calls).toHaveLength(1);
  });

  it("creates a private channel when the caller owns the team", async () => {
    await seedSelf("self-id");
    const calls = owaGraph([
      json({ value: [{ id: "self-id" }] }),
      json({ id: "19:chnew", displayName: "AI TEST — teams-mcp", membershipType: "private" }),
    ]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "AI TEST — teams-mcp",
      private: true,
      confirm: true,
    });
    expect(result.isError).toBeUndefined();
    expect(payload(result)).toMatchObject({ ok: true, kind: "channel", membership: "private" });
    const body = calls[1]?.body as Record<string, unknown>;
    expect(body.membershipType).toBe("private");
    expect((body.members as Array<Record<string, unknown>>)[0]?.roles).toEqual(["owner"]);
    expect(await readTeamsState()).toEqual([
      expect.objectContaining({ surface: "channel", id: "19:chnew", team_id: "t1" }),
    ]);
  });

  it("only creates private channels", async () => {
    await seedSelf("self-id");
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "x",
      confirm: true,
    });
    expect(textOf(result)).toContain("private_required");
    expect(calls).toHaveLength(0);
  });
});

describe("pagination cursors", () => {
  const page = (count: number, prefix: string) => ({
    value: Array.from({ length: count }, (_, index) => ({
      id: `${prefix}${index}`,
      messageType: "message",
      body: { contentType: "text", content: `body ${index}` },
    })),
  });

  it("continues a chat read after the first page is cut mid-page", async () => {
    const body = page(40, "m");
    const calls = owaGraph([json(body), json(body)]);
    const first = payload(await teamsTool.handle({ op: "read", chat_id: "19:abc", top: 10 }));
    expect(first.count).toBe(10);
    // The requested 10 rows were all returned, so the answer was not cut short;
    // the cursor is still offered for older history.
    expect(first.truncated).toBe(false);
    expect(first.has_more).toBe(true);
    const cursor = String(first.next_cursor);

    const second = await teamsTool.handle({ op: "read", chat_id: "19:abc", top: 10, cursor });
    expect(second.isError).toBeUndefined();
    const rows = payload(second).messages as Array<Record<string, unknown>>;
    expect(rows.map((row) => row.id)).toEqual(["m10", "m11", "m12", "m13", "m14", "m15", "m16", "m17", "m18", "m19"]);
    expect(calls).toHaveLength(2);
    // The cursor must hold an absolute URL, not the relative path of page one.
    const decoded = JSON.parse(Buffer.from(cursor, "base64url").toString("utf8")) as { u: string; s: number };
    expect(decoded.u).toMatch(/^https:\/\/graph\.microsoft\.com\/v1\.0\/chats\//);
    expect(decoded.s).toBe(10);
  });

  it("continues a channel read and a replies read across a cut page", async () => {
    const body = page(30, "r");
    webGraph([json(body), json(body), json(body), json(body)]);
    const channel = payload(
      await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1", top: 5 }),
    );
    const channelNext = await teamsTool.handle({
      op: "read",
      team_id: "t1",
      channel_id: "c1",
      top: 5,
      cursor: String(channel.next_cursor),
    });
    expect(channelNext.isError).toBeUndefined();
    expect((payload(channelNext).messages as Array<Record<string, unknown>>).map((row) => row.id)).toEqual([
      "r5",
      "r6",
      "r7",
      "r8",
      "r9",
    ]);

    const replies = payload(
      await teamsTool.handle({ op: "replies", team_id: "t1", channel_id: "c1", message_id: "m1", top: 5 }),
    );
    const repliesNext = await teamsTool.handle({
      op: "replies",
      team_id: "t1",
      channel_id: "c1",
      message_id: "m1",
      top: 5,
      cursor: String(replies.next_cursor),
    });
    expect(repliesNext.isError).toBeUndefined();
    expect((payload(repliesNext).messages as Array<Record<string, unknown>>).map((row) => row.id)).toEqual([
      "r5",
      "r6",
      "r7",
      "r8",
      "r9",
    ]);
  });

  it("continues a channel list when a team has more channels than top", async () => {
    const rows = Array.from({ length: 30 }, (_, index) => ({
      id: `c${index}`,
      displayName: `Channel ${index}`,
      membershipType: "standard",
    }));
    const calls = owaGraph([json({ value: rows }), json({ value: rows })]);
    const first = payload(await teamsTool.handle({ op: "list", kind: "channels", team_id: "t1", top: 25 }));
    expect(first.count).toBe(25);
    expect(first.truncated).toBe(false);
    expect(first.has_more).toBe(true);

    const second = await teamsTool.handle({
      op: "list",
      kind: "channels",
      team_id: "t1",
      top: 25,
      cursor: String(first.next_cursor),
    });
    expect(second.isError).toBeUndefined();
    expect((payload(second).channels as Array<Record<string, unknown>>).map((row) => row.id)).toEqual([
      "c25",
      "c26",
      "c27",
      "c28",
      "c29",
    ]);
    expect(calls).toHaveLength(2);
  });

  it("clamps a directly supplied top to the 25-row cap", async () => {
    const body = {
      value: Array.from({ length: 40 }, (_, index) => ({
        id: `m${index}`,
        messageType: "message",
        body: { content: `body ${index}` },
      })),
    };
    const calls = owaGraph([json(body)]);
    const page = payload(await teamsTool.handle({ op: "read", chat_id: "19:abc", top: 999 }));
    expect(page.count).toBe(25);
    expect(calls[0]?.url).toContain("$top=50");
  });

  it("rejects a cursor replayed against another op, target or filter", async () => {
    const calls = owaGraph([json(page(40, "m"))]);
    const first = payload(await teamsTool.handle({ op: "read", chat_id: "19:abc", top: 10 }));
    const cursor = String(first.next_cursor);

    // Same op, different target.
    const otherChat = await teamsTool.handle({ op: "read", chat_id: "19:other", top: 10, cursor });
    expect(textOf(otherChat)).toContain("invalid_cursor");

    // Same op, different filter: include_system changes which rows the offset counts.
    const withSystem = await teamsTool.handle({
      op: "read",
      chat_id: "19:abc",
      top: 10,
      cursor,
      include_system: true,
    });
    expect(textOf(withSystem)).toContain("invalid_cursor");

    // Cross-op: a chat read cursor cannot be replayed on replies.
    const webCalls = webGraph([json({ value: [] })]);
    const replies = await teamsTool.handle({
      op: "replies",
      team_id: "t1",
      channel_id: "19:abc",
      message_id: "m1",
      cursor,
    });
    expect(textOf(replies)).toContain("invalid_cursor");
    expect(webCalls).toHaveLength(0);

    // Cross-op: nor on a chat list (filtered or not).
    const list = await teamsTool.handle({ op: "list", kind: "chats", cursor });
    expect(textOf(list)).toContain("invalid_cursor");
    const filteredList = await teamsTool.handle({ op: "list", kind: "chats", query: "robotics", cursor });
    expect(textOf(filteredList)).toContain("invalid_cursor");

    // Only the first page ever reached a transport.
    expect(calls).toHaveLength(1);
  });

  it("continues an unfiltered chat list and rejects that cursor under a filter", async () => {
    const calls = owaGraph([
      json({
        value: [
          { id: "19:a", topic: "A", chatType: "group" },
          { id: "19:b", topic: "B", chatType: "oneOnOne" },
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/chats?$top=25&$skiptoken=a",
      }),
      json({ value: [{ id: "19:c", topic: "C", chatType: "group" }] }),
    ]);
    const first = payload(await teamsTool.handle({ op: "list", kind: "chats", top: 5 }));
    expect(first.count).toBe(2);
    expect((first.chats as Array<Record<string, unknown>>).map((row) => row.id)).toEqual(["19:a", "19:b"]);
    expect(first.has_more).toBe(true);
    // Unfiltered list fetches one page only, so fewer than `top` with more pages
    // is a cut-short answer.
    expect(first.truncated).toBe(true);
    const cursor = String(first.next_cursor);

    // The same cursor with a query must be refused: the old offset skipped rows
    // the filter would have matched.
    for (const filter of [{ query: "a" }, { chat_type: "group" }]) {
      const refused = await teamsTool.handle({ op: "list", kind: "chats", top: 5, cursor, ...filter });
      expect(textOf(refused)).toContain("invalid_cursor");
    }
    expect(calls).toHaveLength(1);

    const second = await teamsTool.handle({ op: "list", kind: "chats", top: 5, cursor });
    expect(second.isError).toBeUndefined();
    expect((payload(second).chats as Array<Record<string, unknown>>)[0]?.id).toBe("19:c");
    expect(payload(second).has_more).toBe(false);
    expect(calls).toHaveLength(2);
  });

  it("rejects an unbound legacy URL cursor", async () => {
    owaGraph([]);
    const legacy = Buffer.from(
      "https://graph.microsoft.com/v1.0/me/chats?$top=25&$skiptoken=abc",
      "utf8",
    ).toString("base64url");
    const result = await teamsTool.handle({ op: "list", kind: "chats", cursor: legacy });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("invalid_cursor");
  });

  it("rejects a channel-list cursor used on another team or with a query", async () => {
    const rows = Array.from({ length: 30 }, (_, index) => ({
      id: `c${index}`,
      displayName: `Channel ${index}`,
      membershipType: "standard",
    }));
    const calls = owaGraph([json({ value: rows })]);
    const first = payload(await teamsTool.handle({ op: "list", kind: "channels", team_id: "t1", top: 25 }));
    const cursor = String(first.next_cursor);

    const otherTeam = await teamsTool.handle({
      op: "list",
      kind: "channels",
      team_id: "t2",
      top: 25,
      cursor,
    });
    expect(textOf(otherTeam)).toContain("invalid_cursor");

    const filtered = await teamsTool.handle({
      op: "list",
      kind: "channels",
      team_id: "t1",
      top: 25,
      query: "general",
      cursor,
    });
    expect(textOf(filtered)).toContain("invalid_cursor");
    expect(calls).toHaveLength(1);
  });

  it("continues a filtered list on the same query even when its case changed", async () => {
    const calls = owaGraph([
      json({
        value: [{ id: "19:a", topic: "Robotics planning", chatType: "group" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/chats?$top=50&$expand=members&$skiptoken=a",
      }),
      json({ value: [{ id: "19:b", topic: "Robotics recap", chatType: "group" }] }),
    ]);
    const first = payload(await teamsTool.handle({ op: "list", kind: "chats", query: "robotics", top: 1 }));
    expect(first.count).toBe(1);
    expect(first.has_more).toBe(true);
    // The requested one row was returned: paging is not truncation.
    expect(first.truncated).toBe(false);

    const second = await teamsTool.handle({
      op: "list",
      kind: "chats",
      query: "Robotics",
      top: 1,
      cursor: String(first.next_cursor),
    });
    expect(second.isError).toBeUndefined();
    expect((payload(second).chats as Array<Record<string, unknown>>)[0]?.id).toBe("19:b");
    expect(calls).toHaveLength(2);
  });

  it("marks a scan-cap stop as truncated and still hands back a bound cursor", async () => {
    const responses = [
      ...Array.from({ length: 15 }, (_, index) =>
        json({
          value: [{ id: `19:x${index}`, topic: "Nothing relevant", chatType: "group" }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/me/chats?$top=50&$expand=members&$skiptoken=p${index}`,
        }),
      ),
      json({ value: [{ id: "19:last", topic: "Nothing relevant", chatType: "group" }] }),
    ];
    const calls = owaGraph(responses);
    const first = payload(await teamsTool.handle({ op: "list", kind: "chats", query: "robotics", top: 5 }));
    expect(first.count).toBe(0);
    expect(first.has_more).toBe(true);
    expect(first.truncated).toBe(true);
    expect(calls).toHaveLength(15);

    const second = await teamsTool.handle({
      op: "list",
      kind: "chats",
      query: "robotics",
      top: 5,
      cursor: String(first.next_cursor),
    });
    expect(second.isError).toBeUndefined();
    // The last queued page has no nextLink, so the scan is complete.
    expect(payload(second).has_more).toBe(false);
    expect(payload(second).truncated).toBe(false);
  });

  it("continues a joined-teams list the same way", async () => {
    const rows = Array.from({ length: 28 }, (_, index) => ({ id: `t${index}`, displayName: `Team ${index}` }));
    owaGraph([json({ value: rows }), json({ value: rows })]);
    const first = payload(await teamsTool.handle({ op: "list", kind: "teams", top: 25 }));
    const second = await teamsTool.handle({
      op: "list",
      kind: "teams",
      top: 25,
      cursor: String(first.next_cursor),
    });
    expect(second.isError).toBeUndefined();
    expect(payload(second).count).toBe(3);
  });
});

describe("regressions", () => {
  it("never returns a chat of another type when chat_type was requested", async () => {
    owaGraph([
      json({
        value: [
          { id: "19:a", topic: "A", chatType: "group" },
          { id: "19:m", topic: "Meeting", chatType: "meeting" },
        ],
      }),
    ]);
    const result = await teamsTool.handle({ op: "list", kind: "chats", chat_type: "group", top: 25 });
    const body = payload(result);
    expect(body.count).toBe(1);
    expect((body.chats as Array<Record<string, unknown>>)[0]?.id).toBe("19:a");
  });

  it("treats a non-boolean confirm as missing", async () => {
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "send",
      chat_id: "19:existing",
      text: "hi",
      confirm: "true",
    });
    expect(textOf(result)).toContain("confirmation_required");
    expect(calls).toHaveLength(0);
  });

  it("does not hide a user message that merely starts with system-event markup", async () => {
    webGraph([
      json({
        value: [
          {
            id: "u1",
            messageType: "message",
            body: { contentType: "text", content: "<systemEventMessage/> is the payload shape" },
            from: { user: { displayName: "Joel" } },
          },
        ],
      }),
    ]);
    const body = payload(await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" }));
    expect(body.count).toBe(1);
    const row = (body.messages as Array<Record<string, unknown>>)[0];
    expect(row?.system).toBeUndefined();
    expect(String(row?.text)).toContain("payload shape");
  });

  it("caps create members at the declared maximum before any request", async () => {
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "AI TEST — teams-mcp",
      private: true,
      members: Array.from({ length: 21 }, (_, index) => `u${index}@example.com`),
      confirm: true,
    });
    expect(textOf(result)).toContain("invalid_argument");
    expect(calls).toHaveLength(0);
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"invalid_argument"',
    );
  });

  it("dedupes channel members resolved as a UPN and as an AAD id", async () => {
    await seedSelf("self-id");
    const joelId = "33333333-3333-4333-8333-333333333333";
    const calls = owaGraph([
      json({ value: [{ id: "self-id" }] }),
      json({ id: joelId }),
      json({ id: "19:chnew", displayName: "dedupe", membershipType: "private" }),
    ]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "dedupe",
      private: true,
      members: ["joel@example.com", joelId],
      confirm: true,
    });
    expect(result.isError).toBeUndefined();
    const members = (calls[2]?.body as Record<string, unknown>).members as Array<Record<string, unknown>>;
    expect(members).toHaveLength(2);
    expect(members.map((member) => member.roles)).toEqual([["owner"], []]);
  });

  it("caps message text at 500 by default, at 1000 when text_chars asks, and strips HTML", async () => {
    const message = {
      id: "m1",
      messageType: "message",
      body: {
        contentType: "html",
        content: `<p>${"a".repeat(5000)}</p><script>alert(1)</script><img src=x onerror=alert(2)>`,
      },
    };
    webGraph([json({ value: [message] })]);
    const row = (payload(await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" }))
      .messages as Array<Record<string, unknown>>)[0];
    const text = String(row?.text);
    expect(text.length).toBe(500);
    expect(text.endsWith("…")).toBe(true);
    expect(text).not.toContain("<script");
    expect(text).not.toContain("onerror");

    webGraph([json({ value: [message] })]);
    const wider = (payload(
      await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1", text_chars: 1000 }),
    ).messages as Array<Record<string, unknown>>)[0];
    expect(String(wider?.text).length).toBe(1000);
  });

  it("honours the documented systemEventMessage enum as a system row", async () => {
    webGraph([
      json({
        value: [
          { id: "s1", messageType: "systemEventMessage", body: { content: "<systemEventMessage/>" } },
          { id: "m1", messageType: "message", body: { content: "real" } },
        ],
      }),
    ]);
    const body = payload(await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" }));
    expect(body.count).toBe(1);
    expect((body.messages as Array<Record<string, unknown>>)[0]?.id).toBe("m1");
  });

  it("audits a channel create that is missing its name", async () => {
    const calls = owaGraph([]);
    const result = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      topic: "topic but no name",
      private: true,
      confirm: true,
    });
    expect(textOf(result)).toContain("invalid_argument");
    expect(calls).toHaveLength(0);
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"invalid_argument"',
    );
  });

  it("drops malformed upstream rows instead of failing the page", async () => {
    webGraph([json({ value: [null, 1, "x", [], { id: "m1", messageType: "message", body: { content: "ok" } }] })]);
    const body = payload(await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" }));
    expect(body.count).toBe(1);
  });

  it("returns an empty page for a system-only channel, and the event when asked", async () => {
    webGraph([
      json({ value: [{ id: "s1", messageType: "unknownFutureValue", body: { content: "<systemEventMessage/>" } }] }),
      json({ value: [{ id: "s1", messageType: "unknownFutureValue", body: { content: "<systemEventMessage/>" } }] }),
    ]);
    const hidden = payload(await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1" }));
    expect(hidden.count).toBe(0);
    expect(hidden.truncated).toBe(false);

    const shown = payload(
      await teamsTool.handle({ op: "read", team_id: "t1", channel_id: "c1", include_system: true }),
    );
    expect(shown.count).toBe(1);
    expect((shown.messages as Array<Record<string, unknown>>)[0]?.system).toBe(true);
  });

  it("audits and refunds the rate token when a channel create is refused after the guard", async () => {
    await seedSelf("self-id");
    process.env.TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE = "1";
    resetConfigCache();
    resetWriteLimiter();
    const calls = owaGraph([json({ value: [{ id: "someone-else" }] }), json({ id: "m1" })]);
    const refused = await teamsTool.handle({
      op: "create",
      kind: "channel",
      team_id: "t1",
      name: "AI TEST — teams-mcp",
      private: true,
      confirm: true,
    });
    expect(textOf(refused)).toContain("not_owner");

    await recordCreatedTarget({
      surface: "chat",
      id: "19:ours",
      created_at: new Date().toISOString(),
    });
    const sent = await teamsTool.handle({ op: "send", chat_id: "19:ours", text: "hi", confirm: true });
    // The refusal refunded its token, so the next write is not rate limited.
    expect(sent.isError).toBeUndefined();
    expect(calls).toHaveLength(2);
    expect(await readFile(join(dir, "cache", "write-audit.jsonl"), "utf8")).toContain(
      '"outcome":"not_owner"',
    );
  });
});

describe("status", () => {
  it("reports a stale cache without touching the network", async () => {
    await writeTeamsWebCache({
      version: 1,
      client_id: "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
      scope: "https://graph.microsoft.com/.default",
      access_token: "stale-token",
      expires_at: new Date(Date.now() - 60_000).toISOString(),
      source: "refresh",
      updated_at: new Date().toISOString(),
    });
    const calls = owaGraph([]);
    const body = payload(await teamsTool.handle({ op: "status" }));
    expect(body.teams_web).toMatchObject({ authenticated: true, fresh: false, refreshable: false });
    expect(String(body.channel_history_token)).toContain("stale");
    expect(calls).toHaveLength(0);
  });

  it("is offline and reports the cache path", async () => {
    const result = await teamsTool.handle({ op: "status" });
    expect(result.isError).toBeUndefined();
    const body = payload(result);
    expect(body.teams_web).toMatchObject({ authenticated: false, client_id: "5e3ce6c0-2b1f-4285-8d4b-75ee78787346" });
  });
});
