/**
 * Common-use mail primitives: local-time rows, compact rows, and thread bodies.
 *
 * An agent use-case trial surfaced a wrong weekday because mail rows carried UTC
 * only; `brief` rows and `thread{include_body}` exist so
 * triage and thread summaries do not need one `get` per message.
 */
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { mailTool } from "../src/tools/mail.js";
import type { FetchLike } from "../src/http/client.js";

const TZ = "America/Chicago";

interface Call {
  method: string;
  url: string;
}

function harness(handler: (url: string) => unknown) {
  const calls: Call[] = [];
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({ method: init?.method ?? "GET", url });
    return new Response(JSON.stringify(handler(url)), { status: 200 });
  };
  setOutlookClient(
    new OutlookClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
  return calls;
}

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

const message = (id: string, overrides: Record<string, unknown> = {}) => ({
  Id: id,
  ConversationId: "conv-1",
  Subject: `subject ${id}`,
  From: { EmailAddress: { Address: "sender@example.com" } },
  ReceivedDateTime: "2026-11-12T00:13:20Z",
  IsRead: false,
  BodyPreview: "preview text",
  ...overrides,
});

const rowsOf = (result: CallToolResult): Array<Record<string, unknown>> =>
  (payload(result).messages ?? []) as Array<Record<string, unknown>>;

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-mail-rows-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  process.env.TEAMS_MCP_TIMEZONE = TZ;
  resetConfigCache();
});

afterEach(async () => {
  setOutlookClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("received_local", () => {
  it("renders the configured zone alongside the UTC instant", async () => {
    harness(() => ({ value: [message("m1")] }));
    const [row] = rowsOf(await mailTool.handle({ op: "list", top: 1 }));
    // 2026-11-12T00:13:20Z == Wed 2026-11-11 18:13 CST — the eval's off-by-one.
    expect(row?.received).toBe("2026-11-12T00:13:20.000Z");
    expect(row?.received_local).toBe("2026-11-11 18:13");
  });

  it("renders UTC unchanged when the zone is UTC", async () => {
    process.env.TEAMS_MCP_TIMEZONE = "UTC";
    resetConfigCache();
    harness(() => ({ value: [message("m1")] }));
    const [row] = rowsOf(await mailTool.handle({ op: "list", top: 1 }));
    expect(row?.received_local).toBe("2026-11-12 00:13");
  });

  it("tracks the offset across DST transitions", async () => {
    harness(() => ({
      value: [
        // Spring forward 2027-03-14: 02:00 CST -> 03:00 CDT.
        message("before", { ReceivedDateTime: "2027-03-14T07:59:00Z" }),
        message("after", { ReceivedDateTime: "2027-03-14T08:00:00Z" }),
        // Summer (CDT, -05:00) and winter (CST, -06:00).
        message("summer", { ReceivedDateTime: "2026-07-01T12:00:00Z" }),
        message("winter", { ReceivedDateTime: "2026-12-01T12:00:00Z" }),
      ],
    }));
    const rows = rowsOf(await mailTool.handle({ op: "list", top: 4 }));
    expect(rows.map((row) => row.received_local)).toEqual([
      "2027-03-14 01:59",
      "2027-03-14 03:00",
      "2026-07-01 07:00",
      "2026-12-01 06:00",
    ]);
  });

  it("is present on get and thread rows too", async () => {
    harness(() => ({
      Id: "m1",
      ConversationId: "conv-1",
      Subject: "s",
      ReceivedDateTime: "2026-11-12T00:13:20Z",
      IsRead: false,
      Body: { ContentType: "Text", Content: "hi" },
    }));
    const got = payload(await mailTool.handle({ op: "get", message_id: "m1" }));
    expect(got.received_local).toBe("2026-11-11 18:13");

    harness(() => ({ value: [message("m1")] }));
    const [threadRow] = rowsOf(await mailTool.handle({ op: "thread", thread_id: "conv-1" }));
    expect(threadRow?.received_local).toBe("2026-11-11 18:13");
    expect(threadRow?.unread).toBe(true);
  });
});

describe("brief rows", () => {
  it("keeps only id, subject, from, received, received_local, unread", async () => {
    harness(() => ({
      value: [
        message("m1", {
          BodyPreview: "a long preview that brief rows must not carry",
          HasAttachments: false,
          Flag: { FlagStatus: "Flagged" },
          WebLink: "https://outlook.office365.com/owa/?ItemID=x",
        }),
      ],
    }));
    const out = payload(await mailTool.handle({ op: "list", top: 1, brief: true }));
    const [row] = out.messages as Array<Record<string, unknown>>;
    expect(Object.keys(row ?? {}).sort()).toEqual([
      "from",
      "id",
      "received",
      "received_local",
      "subject",
      "unread",
    ]);
    expect(row?.unread).toBe(true);
  });

  it("adds has_attachments only when true", async () => {
    harness(() => ({ value: [message("m1", { HasAttachments: true })] }));
    const [row] = rowsOf(await mailTool.handle({ op: "list", top: 1, brief: true }));
    expect(row?.has_attachments).toBe(true);
    expect(Object.keys(row ?? {}).sort()).toEqual([
      "from",
      "has_attachments",
      "id",
      "received",
      "received_local",
      "subject",
      "unread",
    ]);
  });

  it("applies to search as well", async () => {
    harness(() => ({ value: [message("m1", { HasAttachments: true })] }));
    const [row] = rowsOf(await mailTool.handle({ op: "search", query: "invoice", brief: true }));
    expect(Object.keys(row ?? {}).sort()).toEqual([
      "from",
      "has_attachments",
      "id",
      "received",
      "received_local",
      "subject",
      "unread",
    ]);
  });

  it("is substantially smaller than a full row", async () => {
    harness(() => ({
      value: [
        message("AAMk".padEnd(140, "x"), {
          BodyPreview: "p".repeat(300),
          Flag: { FlagStatus: "Flagged" },
          WebLink: `https://outlook.office365.com/owa/?ItemID=${"z".repeat(120)}`,
        }),
      ],
    }));
    const full = rowsOf(await mailTool.handle({ op: "list", top: 1, include_web_link: true }))[0];
    const brief = rowsOf(await mailTool.handle({ op: "list", top: 1, brief: true }))[0];
    const fullBytes = JSON.stringify(full).length;
    const briefBytes = JSON.stringify(brief).length;
    expect(briefBytes).toBeLessThan(fullBytes * 0.6);
  });

  it("is off by default: a normal row keeps preview, ids and flag", async () => {
    harness(() => ({
      value: [message("m1", { Flag: { FlagStatus: "Flagged" }, HasAttachments: true })],
    }));
    const [row] = rowsOf(await mailTool.handle({ op: "list", top: 1 }));
    expect(row?.preview).toBe("preview text");
    expect(row?.conversation_id).toBe("conv-1");
    expect(row?.flag).toBe(true);
    expect(row?.has_attachments).toBe(true);
  });
});

describe("preview_chars", () => {
  it("caps the preview at the requested length", async () => {
    harness(() => ({ value: [message("m1", { BodyPreview: "x".repeat(300) })] }));
    const [row] = rowsOf(await mailTool.handle({ op: "list", top: 1, preview_chars: 40 }));
    const preview = String(row?.preview ?? "");
    expect(preview.length).toBeLessThanOrEqual(40);
    expect(preview.endsWith("…")).toBe(true);
  });

  it("clamps an oversized request to 1000 characters", async () => {
    harness(() => ({ value: [message("m1", { BodyPreview: "y".repeat(1_200) })] }));
    const [row] = rowsOf(await mailTool.handle({ op: "list", top: 1, preview_chars: 5_000 }));
    expect(String(row?.preview ?? "").length).toBeLessThanOrEqual(1_000);
  });

  it("include_body_preview=false still wins: no preview at all", async () => {
    harness(() => ({ value: [message("m1", { BodyPreview: "x".repeat(300) })] }));
    const [row] = rowsOf(
      await mailTool.handle({
        op: "list",
        top: 1,
        preview_chars: 500,
        include_body_preview: false,
      }),
    );
    expect(row?.preview).toBeUndefined();
  });

  it("reaches search rows", async () => {
    harness(() => ({ value: [message("m1", { BodyPreview: "z".repeat(300) })] }));
    const [row] = rowsOf(await mailTool.handle({ op: "search", query: "q", preview_chars: 50 }));
    expect(String(row?.preview ?? "").length).toBeLessThanOrEqual(50);
  });
});

describe("thread bodies", () => {
  const threadFixture = (overrides: Record<string, unknown> = {}) => ({
    value: [
      message("m1", {
        ReceivedDateTime: "2026-11-12T00:13:20Z",
        IsRead: true,
        Body: { ContentType: "Text", Content: "first body" },
        ...overrides,
      }),
    ],
  });

  it("does not request or return bodies by default", async () => {
    const calls = harness(() => threadFixture());
    const [row] = rowsOf(await mailTool.handle({ op: "thread", thread_id: "conv-1" }));
    const select = new URL(calls[0]?.url ?? "").searchParams.get("$select") ?? "";
    expect(select.split(",")).not.toContain("Body");
    expect(row?.body).toBeUndefined();
    expect(row?.unread).toBe(false);
  });

  it("inlines a body when include_body is true", async () => {
    const calls = harness(() => threadFixture());
    const [row] = rowsOf(
      await mailTool.handle({ op: "thread", thread_id: "conv-1", include_body: true }),
    );
    const select = new URL(calls[0]?.url ?? "").searchParams.get("$select") ?? "";
    expect(select.split(",")).toContain("Body");
    expect(row?.body).toBe("first body");
    expect(row?.body_truncated).toBeUndefined();
  });

  it("decodes an HTML body to text", async () => {
    harness(() =>
      threadFixture({
        Body: { ContentType: "HTML", Content: "<p>Hello&nbsp;<b>there</b></p>" },
      }),
    );
    const [row] = rowsOf(
      await mailTool.handle({ op: "thread", thread_id: "conv-1", include_body: true }),
    );
    expect(row?.body).toContain("Hello");
    expect(row?.body).toContain("there");
    expect(row?.body).not.toContain("<");
  });

  it("defaults to 500 chars and marks a cut body", async () => {
    harness(() => threadFixture({ Body: { ContentType: "Text", Content: "b".repeat(600) } }));
    const [row] = rowsOf(
      await mailTool.handle({ op: "thread", thread_id: "conv-1", include_body: true }),
    );
    expect(String(row?.body ?? "").length).toBeLessThanOrEqual(500);
    expect(row?.body_truncated).toBe(true);
  });

  it("honours body_chars and clamps it to 2000", async () => {
    harness(() => threadFixture({ Body: { ContentType: "Text", Content: "c".repeat(3_000) } }));
    const short = rowsOf(
      await mailTool.handle({
        op: "thread",
        thread_id: "conv-1",
        include_body: true,
        body_chars: 120,
      }),
    )[0];
    expect(String(short?.body ?? "").length).toBeLessThanOrEqual(120);

    const capped = rowsOf(
      await mailTool.handle({
        op: "thread",
        thread_id: "conv-1",
        include_body: true,
        body_chars: 9_000,
      }),
    )[0];
    expect(String(capped?.body ?? "").length).toBeLessThanOrEqual(2_000);
    expect(capped?.body_truncated).toBe(true);
  });
});

describe("top cap", () => {
  it("clamps to max_results.mail (default 25)", async () => {
    const calls = harness(() => ({ value: [message("m1")] }));
    await mailTool.handle({ op: "list", top: 50 });
    expect(new URL(calls[0]?.url ?? "").searchParams.get("$top")).toBe("25");

    await mailTool.handle({ op: "search", query: "q", top: 50 });
    expect(new URL(calls[1]?.url ?? "").searchParams.get("$top")).toBe("25");
  });

  it("allows 50 when the operator raised max_results.mail", async () => {
    await writeFile(join(dir, "config.yaml"), "max_results:\n  mail: 50\n");
    resetConfigCache();
    const calls = harness(() => ({ value: [message("m1")] }));
    await mailTool.handle({ op: "list", top: 50 });
    expect(new URL(calls[0]?.url ?? "").searchParams.get("$top")).toBe("50");
  });
});
