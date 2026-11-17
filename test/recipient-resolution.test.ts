import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { GraphClient, setGraphClient } from "../src/graph/client.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { setSentLookupDelays } from "../src/outlook/mail.js";
import { calendarTool } from "../src/tools/calendar.js";
import { mailTool } from "../src/tools/mail.js";
import type { FetchLike } from "../src/http/client.js";

const OUTLOOK = "https://outlook.office.com/api/v2.0";

interface Call {
  api: "outlook" | "graph";
  method: string;
  url: string;
  body: unknown;
}

function harness(responses: { graph?: Array<() => Response>; outlook?: Array<() => Response> }) {
  const calls: Call[] = [];
  let graphIndex = 0;
  let outlookIndex = 0;
  const graphFetch: FetchLike = async (url, init) => {
    calls.push({
      api: "graph",
      method: init?.method ?? "GET",
      url,
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    });
    const factory = responses.graph?.[graphIndex];
    graphIndex += 1;
    if (!factory) throw new Error(`unexpected graph call ${url}`);
    return factory();
  };
  const outlookFetch: FetchLike = async (url, init) => {
    calls.push({
      api: "outlook",
      method: init?.method ?? "GET",
      url,
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    });
    const factory = responses.outlook?.[outlookIndex];
    outlookIndex += 1;
    if (!factory) throw new Error(`unexpected outlook call ${url}`);
    return factory();
  };
  setGraphClient(new GraphClient({ tokenProvider: async () => ({ token: "g" }), fetchImpl: graphFetch, maxAttempts: 1 }));
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl: outlookFetch, maxAttempts: 1 }),
  );
  return calls;
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });
const empty = (status: number): (() => Response) => () => new Response(null, { status });

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

const joel = {
  id: "id-joel",
  displayName: "Student, Test",
  mail: "joel@example.edu",
  userPrincipalName: "joel@example.edu",
};
const joelWork = {
  id: "id-joel-2",
  displayName: "Student, Test (SW)",
  mail: "joel.alt@example.edu",
  userPrincipalName: "joel.alt@example.edu",
};

function event(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    Id: "evt-1",
    Subject: "AI TEST — teams-mcp — resolution",
    Start: { DateTime: "2026-11-18T16:00:00.0000000", TimeZone: "UTC" },
    End: { DateTime: "2026-11-18T16:30:00.0000000", TimeZone: "UTC" },
    IsAllDay: false,
    ShowAs: "Busy",
    Attendees: [],
    ...overrides,
  };
}

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-recipients-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  resetConfigCache();
  // No real Sent Items backoff while asserting call order.
  setSentLookupDelays([0, 0]);
});

afterEach(async () => {
  setGraphClient(undefined);
  setOutlookClient(undefined);
  setSentLookupDelays(undefined);
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("recipient resolution on mail.send", () => {
  it("refuses an ambiguous name with every candidate and sends nothing", async () => {
    const calls = harness({ graph: [json({ value: [joel, joelWork] })] });
    const result = await mailTool.handle({
      op: "send",
      to: ["Test Student"],
      subject: "x",
      body: "y",
    });
    expect(result.isError).toBe(true);
    const out = payload(result);
    expect(out).toMatchObject({
      error: "ambiguous_recipient",
      input: "Test Student",
      needs_disambiguation: true,
    });
    expect(out.candidates).toEqual([
      { id: "id-joel", name: "Student, Test", email: "joel@example.edu" },
      { id: "id-joel-2", name: "Student, Test (SW)", email: "joel.alt@example.edu" },
    ]);
    expect(calls.some((call) => call.api === "outlook")).toBe(false);
  });

  it("uses the single directory match for to, cc and bcc", async () => {
    const calls = harness({
      graph: [json({ value: [joel] })],
      // 202 + the zero-delay Sent Items id lookups.
      outlook: [empty(202), json({ value: [] }), json({ value: [] }), json({ value: [] })],
    });
    const result = await mailTool.handle({
      op: "send",
      to: ["Test Student"],
      cc: ["Test Student"],
      bcc: ["someone@example.com"],
      subject: "x",
      body: "y",
    });
    expect(result.isError).toBeUndefined();
    const sendmail = calls.find((call) => call.url.endsWith("/me/sendmail"));
    expect(sendmail?.body).toMatchObject({
      Message: {
        ToRecipients: [{ EmailAddress: { Address: "joel@example.edu" } }],
        CcRecipients: [{ EmailAddress: { Address: "joel@example.edu" } }],
        BccRecipients: [{ EmailAddress: { Address: "someone@example.com" } }],
      },
    });
    expect(payload(result)).toMatchObject({ ok: true, op: "send", recipient_count: 3 });
  });

  it("reports a name with no directory match instead of guessing", async () => {
    const calls = harness({ graph: [json({ value: [] })] });
    const result = await mailTool.handle({ op: "send", to: ["Nobody Here"], subject: "x", body: "y" });
    expect(result.isError).toBe(true);
    expect(payload(result)).toMatchObject({
      error: "recipient_not_found",
      input: "Nobody Here",
    });
    expect(calls.some((call) => call.api === "outlook")).toBe(false);
  });

  it("allowlist-checks the resolved address", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@example.com"]\n');
    resetConfigCache();
    const calls = harness({ graph: [json({ value: [joel] })] });
    const result = await mailTool.handle({ op: "send", to: ["Test Student"], subject: "x", body: "y" });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain(
      "recipient_not_allowed",
    );
    expect(calls.some((call) => call.api === "outlook")).toBe(false);
  });

  it("prefers a configured alias over a directory lookup", async () => {
    await writeFile(join(dir, "config.yaml"), "aliases: {joel: joel@example.edu}\n");
    resetConfigCache();
    const calls = harness({
      outlook: [empty(202), json({ value: [] }), json({ value: [] }), json({ value: [] })],
    });
    const result = await mailTool.handle({ op: "send", to: ["Joel"], subject: "x", body: "y" });
    expect(result.isError).toBeUndefined();
    expect(calls.filter((call) => call.api === "graph")).toHaveLength(0);
    expect(calls.find((call) => call.url.endsWith("/me/sendmail"))?.body).toMatchObject({
      Message: { ToRecipients: [{ EmailAddress: { Address: "joel@example.edu" } }] },
    });
  });
});

describe("recipient resolution on mail.forward", () => {
  it("resolves the forward target", async () => {
    const calls = harness({
      graph: [json({ value: [joel] })],
      outlook: [
        json({ Id: "m1", ConversationId: "conv-m1", Subject: "Agenda" }),
        empty(202),
        json({ value: [] }),
        json({ value: [] }),
        json({ value: [] }),
      ],
    });
    const result = await mailTool.handle({ op: "forward", message_id: "m1", to: ["Test Student"] });
    expect(result.isError).toBeUndefined();
    expect(calls.find((call) => call.url.endsWith("/forward"))?.body).toEqual({
      ToRecipients: [{ EmailAddress: { Address: "joel@example.edu" } }],
    });
  });

  it("stops on ambiguity before calling Outlook", async () => {
    const calls = harness({ graph: [json({ value: [joel, joelWork] })] });
    const result = await mailTool.handle({ op: "forward", message_id: "m1", to: ["Test Student"] });
    expect(result.isError).toBe(true);
    expect(calls.some((call) => call.api === "outlook")).toBe(false);
  });
});

describe("recipient resolution on calendar", () => {
  it("creates an event with a directory-resolved attendee", async () => {
    const calls = harness({
      graph: [json({ value: [joel] })],
      outlook: [json({ value: [] }), json(event({ Subject: "AI TEST — teams-mcp — resolved" }), 201)],
    });
    const result = await calendarTool.handle({
      op: "create",
      subject: "AI TEST — teams-mcp — resolved",
      start: "2026-11-18T16:00:00Z",
      attendees: ["Test Student"],
    });
    expect(result.isError).toBeUndefined();
    const post = calls.find((call) => call.api === "outlook" && call.url === `${OUTLOOK}/me/events`);
    expect(post?.body).toMatchObject({
      Attendees: [{ EmailAddress: { Address: "joel@example.edu" }, Type: "Required" }],
    });
  });

  it("refuses an ambiguous attendee without creating an event", async () => {
    const calls = harness({ graph: [json({ value: [joel, joelWork] })] });
    const result = await calendarTool.handle({
      op: "create",
      subject: "Ambiguous",
      start: "2026-11-18T16:00:00Z",
      attendees: ["Test Student"],
    });
    expect(result.isError).toBe(true);
    expect(payload(result)).toMatchObject({ error: "ambiguous_recipient", needs_disambiguation: true });
    expect(calls.some((call) => call.api === "outlook")).toBe(false);
  });

  it("REGRESSION: update applies the invite_allowlist to new attendees", async () => {
    await writeFile(join(dir, "config.yaml"), 'invite_allowlist: ["@example.com"]\n');
    resetConfigCache();
    const calls = harness({ outlook: [json(event())] });
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "AI TEST — teams-mcp — resolution",
      attendees: ["outsider@other.test"],
    });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain(
      "recipient_not_allowed",
    );
    // Only the guard GET; no PATCH.
    expect(calls.map((call) => call.method)).toEqual(["GET"]);
  });

  it("resolves update attendees and patches the address", async () => {
    const calls = harness({
      graph: [json({ value: [joel] })],
      outlook: [json(event()), json(event())],
    });
    const result = await calendarTool.handle({
      op: "update",
      event_id: "evt-1",
      expected_subject: "AI TEST — teams-mcp — resolution",
      attendees: ["Test Student"],
    });
    expect(result.isError).toBeUndefined();
    const patch = calls.find((call) => call.method === "PATCH");
    expect(patch?.body).toEqual({
      Attendees: [{ EmailAddress: { Address: "joel@example.edu" }, Type: "Required" }],
    });
  });
});
