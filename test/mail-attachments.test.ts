/**
 * Read-only attachment ops on the `mail` tool.
 *
 * The live gap these close: an attachment-only message reported only
 * `has_attachments: true`, so its content could not be read at all. Metadata is
 * one bounded list call; `get_attachment` either decodes text-ish content in the
 * result or saves bytes under the cache directory with 0600 — never executes or
 * opens them, never mutates the mailbox, and never consumes a write token.
 */
import { mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { sanitizeAttachmentName } from "../src/outlook/attachments.js";
import { peekWriteToken, resetWriteLimiter } from "../src/safety.js";
import { mailTool } from "../src/tools/mail.js";
import type { FetchLike } from "../src/http/client.js";

const BASE = "https://outlook.office.com/api/v2.0";

interface Call {
  method: string;
  url: string;
}

function harness(handler: (url: string) => Response | unknown) {
  const calls: Call[] = [];
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({ method: init?.method ?? "GET", url });
    const answer = handler(url);
    if (answer instanceof Response) return answer;
    return new Response(JSON.stringify(answer), { status: 200 });
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

const meta = (overrides: Record<string, unknown> = {}) => ({
  Id: "att-1",
  Name: "report.pdf",
  ContentType: "application/pdf",
  Size: 4,
  IsInline: false,
  ...overrides,
});

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-attachments-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  resetConfigCache();
  resetWriteLimiter();
});

afterEach(async () => {
  setOutlookClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("mail.attachments", () => {
  it("lists metadata without asking for content", async () => {
    const calls = harness(() => ({ value: [meta(), meta({ Id: "att-2", Name: "cal.ics", ContentType: "text/calendar", Size: 9 })] }));
    const result = await mailTool.handle({ op: "attachments", message_id: "msg-1" });

    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    const url = new URL(calls[0]!.url);
    expect(`${url.origin}${url.pathname}`).toBe(`${BASE}/me/messages/msg-1/attachments`);
    expect(url.searchParams.get("$select")).toBe("Id,Name,ContentType,Size,IsInline");
    expect(url.searchParams.get("$top")).toBe("20");
    expect(url.toString()).not.toContain("ContentBytes");

    const out = payload(result);
    expect(out).toMatchObject({ message_id: "msg-1", count: 2, truncated: false });
    expect(out.attachments).toEqual([
      { id: "att-1", name: "report.pdf", content_type: "application/pdf", size: 4, is_inline: false },
      { id: "att-2", name: "cal.ics", content_type: "text/calendar", size: 9, is_inline: false },
    ]);
  });

  it("caps the list at 20 rows and says so", async () => {
    harness(() => ({
      value: Array.from({ length: 25 }, (_, index) => meta({ Id: `att-${index}`, Name: `f${index}.bin` })),
    }));
    const out = payload(await mailTool.handle({ op: "attachments", message_id: "msg-1" }));
    expect(out.count).toBe(20);
    expect((out.attachments as unknown[]).length).toBe(20);
    expect(out.truncated).toBe(true);
    expect(String(out.note)).toContain("get_attachment");
  });
});

describe("mail.get_attachment", () => {
  it("decodes a text attachment and truncates the text at 20 000 chars", async () => {
    const text = `${"x".repeat(19_999)}END`;
    const calls = harness(() => ({
      ...meta({ Name: "note.txt", ContentType: "text/plain", Size: Buffer.byteLength(text) }),
      ContentBytes: Buffer.from(text, "utf8").toString("base64"),
    }));
    const result = await mailTool.handle({
      op: "get_attachment",
      message_id: "msg-1",
      attachment_id: "att-1",
    });

    expect(result.isError).toBeUndefined();
    expect(calls[0]?.url).toBe(`${BASE}/me/messages/msg-1/attachments/att-1`);
    const out = payload(result);
    expect(out).toMatchObject({
      id: "att-1",
      name: "note.txt",
      content_type: "text/plain",
      text_truncated: true,
    });
    expect((out.text as string).length).toBe(20_000);
    expect(out.text).toBe(text.slice(0, 20_000));
    // Text is returned in the result; nothing is written to disk.
    await expect(readdir(join(dir, "cache", "attachments"))).rejects.toThrow();
  });

  it("decodes calendar media types as text", async () => {
    const ics = "BEGIN:VCALENDAR\r\nEND:VCALENDAR";
    harness(() => ({
      ...meta({ Name: "invite.ics", ContentType: "application/calendar+json", Size: 29 }),
      ContentBytes: Buffer.from(ics).toString("base64"),
    }));
    const out = payload(
      await mailTool.handle({ op: "get_attachment", message_id: "msg-1", attachment_id: "att-1" }),
    );
    expect(out.text).toBe(ics);
    expect(out.saved).toBeUndefined();
  });

  it("text:true forces decoding for a textual type outside text/*", async () => {
    const body = JSON.stringify({ hello: "world" });
    harness(() => ({
      ...meta({ Name: "payload.json", ContentType: "application/json", Size: body.length }),
      ContentBytes: Buffer.from(body).toString("base64"),
    }));
    const out = payload(
      await mailTool.handle({
        op: "get_attachment",
        message_id: "msg-1",
        attachment_id: "att-1",
        text: true,
      }),
    );
    expect(out.text).toBe(body);
    expect(out.saved).toBeUndefined();
  });

  it("saves binary bytes with 0600 inside the cache dir and a unique name", async () => {
    const bytes = Buffer.from([0x00, 0x01, 0xfe, 0xff]);
    const calls = harness(() => ({
      ...meta(),
      ContentBytes: bytes.toString("base64"),
    }));
    const args = { op: "get_attachment", message_id: "msg-1", attachment_id: "att-1" } as const;

    const first = payload(await mailTool.handle({ ...args }));
    const second = payload(await mailTool.handle({ ...args }));

    expect(calls).toHaveLength(2);
    expect(first).toMatchObject({ saved: true, size: 4, content_type: "application/pdf" });
    expect(String(first.path).startsWith(join(dir, "cache", "attachments"))).toBe(true);
    expect(String(first.path).endsWith(".pdf")).toBe(true);
    expect(await readFile(String(first.path))).toEqual(bytes);
    expect((await stat(String(first.path))).mode & 0o777).toBe(0o600);
    expect((await stat(join(dir, "cache", "attachments"))).mode & 0o777).toBe(0o700);
    // A random suffix keeps a second read of the same attachment distinct.
    expect(second.path).not.toBe(first.path);
    expect(await readFile(String(second.path))).toEqual(bytes);
  });

  it("strips traversal from the attachment name", async () => {
    harness(() => ({
      ...meta({ Name: "../../etc/passwd", ContentType: "application/octet-stream" }),
      ContentBytes: Buffer.from("secret").toString("base64"),
    }));
    const out = payload(
      await mailTool.handle({ op: "get_attachment", message_id: "msg-1", attachment_id: "att-1" }),
    );
    const path = String(out.path);
    expect(path.startsWith(join(dir, "cache", "attachments"))).toBe(true);
    expect(path.includes("..")).toBe(false);
    expect(path.includes(`${join(dir, "cache", "attachments")}/passwd`)).toBe(false);
    expect(path).toContain("-passwd");
    expect(await readFile(path, "utf8")).toBe("secret");

    expect(sanitizeAttachmentName("..\\..\\evil.exe")).toBe("evil.exe");
    expect(sanitizeAttachmentName("a".repeat(150) + ".txt")).toHaveLength(100);
    expect(sanitizeAttachmentName("")).toBe("attachment");
  });

  it("refuses more than 10 MB decoded without saving", async () => {
    const bytes = Buffer.alloc(10 * 1024 * 1024 + 1, 0x61);
    harness(() => ({
      ...meta({ Name: "huge.bin", ContentType: "application/octet-stream", Size: bytes.length }),
      ContentBytes: bytes.toString("base64"),
    }));
    const out = payload(
      await mailTool.handle({ op: "get_attachment", message_id: "msg-1", attachment_id: "att-1" }),
    );
    expect(out.note).toBe("too_large");
    expect(out.saved).toBeUndefined();
    expect(out.path).toBeUndefined();
    await expect(readdir(join(dir, "cache", "attachments"))).rejects.toThrow();
  });

  it("reports a reference attachment with no content", async () => {
    harness(() => meta({ Name: "linked.docx", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    const out = payload(
      await mailTool.handle({ op: "get_attachment", message_id: "msg-1", attachment_id: "att-1" }),
    );
    expect(out.note).toBe("content_unavailable");
    expect(out.saved).toBeUndefined();
  });

  it("rejects missing ids before any request", async () => {
    const calls = harness(() => ({ value: [] }));
    const noMessage = await mailTool.handle({ op: "attachments" });
    const noAttachment = await mailTool.handle({ op: "get_attachment", message_id: "msg-1" });
    expect(noMessage.isError).toBe(true);
    expect(noAttachment.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it("maps a malformed attachment id from the service to a tool error", async () => {
    harness(
      () =>
        new Response(
          JSON.stringify({ error: { code: "ErrorInvalidIdMalformed", message: "Malformed id." } }),
          { status: 400, headers: { "content-type": "application/json" } },
        ),
    );
    const result = await mailTool.handle({
      op: "get_attachment",
      message_id: "bad",
      attachment_id: "bad",
    });
    expect(result.isError).toBe(true);
    expect(result.content[0]?.type === "text" ? result.content[0].text : "").toContain(
      "ErrorInvalidIdMalformed",
    );
  });

  it("works in read-only mode and consumes no write token", async () => {
    process.env.TEAMS_MCP_READ_ONLY = "1";
    resetConfigCache();
    harness((url) =>
      url.endsWith("/att-1")
        ? { ...meta({ Name: "note.txt", ContentType: "text/plain", Size: 5 }), ContentBytes: Buffer.from("hello").toString("base64") }
        : { value: [meta()] },
    );

    const listed = await mailTool.handle({ op: "attachments", message_id: "msg-1" });
    expect(listed.isError).toBeUndefined();

    for (let index = 0; index < 35; index += 1) {
      const read = await mailTool.handle({
        op: "get_attachment",
        message_id: "msg-1",
        attachment_id: "att-1",
      });
      expect(read.isError).toBeUndefined();
      expect(payload(read).text).toBe("hello");
    }
    // The default limit is 30/minute; 36 attachment reads must leave it intact.
    expect(peekWriteToken().ok).toBe(true);

    const write = await mailTool.handle({ op: "send", to: ["a@example.com"], subject: "x", body: "y" });
    expect(write.isError).toBe(true);
    expect(write.content[0]?.type === "text" ? write.content[0].text : "").toContain("read_only");
  });
});
