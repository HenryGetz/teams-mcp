/**
 * Truncation/pagination regressions.
 *
 * Before this suite a `top`-limited thread or recurrence read returned a page
 * with no `next_cursor` and no `truncated` flag, so a model could not tell a
 * 1-message answer from the first message of a 40-message conversation.
 */
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient, type FetchLike as OutlookFetch } from "../src/outlook/client.js";
import { GraphClient, setGraphClient } from "../src/graph/client.js";
import { clearPeopleCache, peopleTool } from "../src/tools/people.js";
import { mailTool } from "../src/tools/mail.js";
import { calendarTool } from "../src/tools/calendar.js";

const OUTLOOK = "https://outlook.office.com/api/v2.0";

interface Call {
  method: string;
  url: string;
}

function outlookHarness(responses: Array<() => Response>): Call[] {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: OutlookFetch = async (url, init) => {
    calls.push({ method: init?.method ?? "GET", url });
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

function graphHarness(responses: Array<() => Response>): Call[] {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: OutlookFetch = async (url, init) => {
    calls.push({ method: init?.method ?? "GET", url });
    const factory = responses[Math.min(index, responses.length - 1)];
    index += 1;
    if (!factory) throw new Error("no response queued");
    return factory();
  };
  setGraphClient(new GraphClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }));
  return calls;
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status });

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

/** Error results are `code: message` text, not JSON. */
function textOf(result: CallToolResult): string {
  const first = result.content[0];
  return first && first.type === "text" ? first.text : "";
}

function message(id: string, overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    Id: id,
    ConversationId: "conv-1",
    Subject: "AI TEST — teams-mcp — thread",
    From: { EmailAddress: { Address: "student@example.edu" } },
    ReceivedDateTime: "2026-11-16T12:00:00Z",
    IsRead: true,
    BodyPreview: "hello",
    ...overrides,
  };
}

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-pagination-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_TIMEZONE = "America/Chicago";
  resetConfigCache();
  clearPeopleCache();
});

afterEach(async () => {
  setOutlookClient(undefined);
  setGraphClient(undefined);
  clearPeopleCache();
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("mail.thread", () => {
  const nextLink = `${OUTLOOK}/me/messages?$filter=ConversationId&$top=1&$skiptoken=page2`;

  it("says a top=1 page is truncated and hands back a cursor", async () => {
    const calls = outlookHarness([
      json({ value: [message("m-1")], "@odata.nextLink": nextLink }),
      json({ value: [message("m-2", { ReceivedDateTime: "2026-11-16T13:00:00Z" })] }),
    ]);
    const first = payload(
      await mailTool.handle({ op: "thread", thread_id: "conv-1", top: 1 }),
    );
    expect(first.count).toBe(1);
    expect(first.truncated).toBe(true);
    expect(typeof first.next_cursor).toBe("string");
    expect(calls[0]?.url).toContain("%24top=1");

    const cursor = String(first.next_cursor);
    const second = payload(
      await mailTool.handle({ op: "thread", thread_id: "conv-1", top: 1, cursor }),
    );
    expect(calls[1]?.url).toBe(nextLink);
    expect(second.messages).toMatchObject([{ id: "m-2" }]);
    expect(second.truncated).toBe(false);
    expect(second.next_cursor).toBeUndefined();
  });

  it("rejects a cursor pointing at another host", async () => {
    const calls = outlookHarness([json({ value: [] })]);
    const foreign = Buffer.from("https://evil.example.com/me/messages", "utf8").toString("base64url");
    const result = await mailTool.handle({ op: "thread", thread_id: "conv-1", cursor: foreign });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("invalid_cursor");
    expect(calls).toHaveLength(0);
  });

  it("rejects a cursor that is not a teams-mcp cursor", async () => {
    const calls = outlookHarness([json({ value: [] })]);
    const result = await mailTool.handle({ op: "thread", thread_id: "conv-1", cursor: "not-a-cursor" });
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("invalid_cursor");
    expect(calls).toHaveLength(0);
  });
});

describe("mail.list and mail.search", () => {
  it("flags a truncated list page", async () => {
    outlookHarness([
      json({
        value: [message("m-1")],
        "@odata.nextLink": `${OUTLOOK}/me/mailfolders/inbox/messages?$top=1&$skiptoken=x`,
      }),
    ]);
    const out = payload(await mailTool.handle({ op: "list", top: 1 }));
    expect(out.count).toBe(1);
    expect(out.truncated).toBe(true);
    expect(typeof out.next_cursor).toBe("string");
  });

  it("follows a search cursor and reports the last page as complete", async () => {
    const calls = outlookHarness([
      json({ value: [message("s-1")], "@odata.nextLink": `${OUTLOOK}/me/messages?$search=x&$top=1&$skiptoken=p2` }),
      json({ value: [message("s-2")] }),
    ]);
    const first = payload(await mailTool.handle({ op: "search", query: "AI TEST", top: 1 }));
    expect(first.truncated).toBe(true);
    const second = payload(
      await mailTool.handle({ op: "search", query: "AI TEST", top: 1, cursor: String(first.next_cursor) }),
    );
    expect(second.count).toBe(1);
    expect(second.truncated).toBe(false);
    expect(calls[1]?.url).toContain("skiptoken=p2");
  });
});

describe("calendar.instances", () => {
  const nextLink = `${OUTLOOK}/me/events/series-1/instances?$top=1&$skiptoken=p2`;

  it("reports the next page instead of silently returning one occurrence", async () => {
    const calls = outlookHarness([
      json({ value: [{ Id: "inst-1", Subject: "Weekly", Start: { DateTime: "2026-11-02T16:00:00.0000000", TimeZone: "UTC" }, End: { DateTime: "2026-11-02T16:30:00.0000000", TimeZone: "UTC" } }], "@odata.nextLink": nextLink }),
      json({ value: [{ Id: "inst-2", Subject: "Weekly", Start: { DateTime: "2026-11-09T16:00:00.0000000", TimeZone: "UTC" }, End: { DateTime: "2026-11-09T16:30:00.0000000", TimeZone: "UTC" } }] }),
    ]);
    const first = payload(
      await calendarTool.handle({
        op: "instances",
        event_id: "series-1",
        start: "2026-11-01T00:00:00Z",
        end: "2026-12-01T00:00:00Z",
        top: 1,
      }),
    );
    expect(first.count).toBe(1);
    expect(first.truncated).toBe(true);
    const second = payload(
      await calendarTool.handle({
        op: "instances",
        event_id: "series-1",
        cursor: String(first.next_cursor),
      }),
    );
    expect(calls[1]?.url).toBe(nextLink);
    expect(second.count).toBe(1);
    expect(second.truncated).toBe(false);
    expect((second.instances as Array<Record<string, unknown>>)[0]?.id).toBe("inst-2");
  });

  it("rejects a cursor from another API host", async () => {
    const calls = outlookHarness([json({ value: [] })]);
    const foreign = Buffer.from("https://graph.microsoft.com/v1.0/me/events", "utf8").toString("base64url");
    const result = await calendarTool.handle({ op: "instances", event_id: "e", cursor: foreign });
    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe("calendar.list", () => {
  it("flags truncation on the first page and on a cursor page", async () => {
    outlookHarness([
      json({
        value: [{ Id: "e-1", Subject: "One", Start: { DateTime: "2026-11-16T16:00:00.0000000", TimeZone: "UTC" }, End: { DateTime: "2026-11-16T16:30:00.0000000", TimeZone: "UTC" } }],
        "@odata.nextLink": `${OUTLOOK}/me/calendarview?$top=1&$skiptoken=p2`,
      }),
      json({
        value: [{ Id: "e-2", Subject: "Two", Start: { DateTime: "2026-11-17T16:00:00.0000000", TimeZone: "UTC" }, End: { DateTime: "2026-11-17T16:30:00.0000000", TimeZone: "UTC" } }],
      }),
    ]);
    const first = payload(
      await calendarTool.handle({
        op: "list",
        start: "2026-11-16",
        end: "2026-11-20",
        top: 1,
      }),
    );
    expect(first.truncated).toBe(true);
    expect(typeof first.next_cursor).toBe("string");
    // A bare date is local midnight, not UTC midnight.
    expect(first.range).toMatchObject({ start: "2026-11-16 00:00", timezone: "America/Chicago" });
    const second = payload(
      await calendarTool.handle({ op: "list", cursor: String(first.next_cursor) }),
    );
    expect(second.truncated).toBe(false);
  });
});

describe("people.search", () => {
  const person = (index: number): Record<string, unknown> => ({
    id: `id-${index}`,
    displayName: `Student ${index}`,
    mail: `student${index}@example.edu`,
    userPrincipalName: `student${index}@example.edu`,
  });

  it("says a full page may be incomplete, including from cache", async () => {
    graphHarness([json({ value: [person(1), person(2)] })]);
    const first = payload(await peopleTool.handle({ op: "search", query: "student", top: 2 }));
    expect(first.count).toBe(2);
    expect(first.truncated).toBe(true);
    expect(String(first.note)).toContain("at most 2");

    const cached = payload(await peopleTool.handle({ op: "search", query: "student", top: 2 }));
    expect(cached.cached).toBe(true);
    expect(cached.truncated).toBe(true);
  });

  it("does not flag a short page", async () => {
    graphHarness([json({ value: [person(1)] })]);
    const out = payload(await peopleTool.handle({ op: "search", query: "student", top: 5 }));
    expect(out.truncated).toBeUndefined();
  });
});
