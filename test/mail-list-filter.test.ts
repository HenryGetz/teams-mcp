/**
 * `mail.list` with a `from` filter.
 *
 * Two live behaviours constrain this path:
 *
 *  1. `$orderby=ReceivedDateTime desc` together with `From/EmailAddress/Address
 *     eq '…'` is rejected by Exchange (`InefficientFilter`), so no sort key is
 *     sent and the order is applied locally.
 *  2. With that filter and a small `$top`, Exchange returned the TAIL of the
 *     match set — `top=1` answered with the OLDEST mail from the sender. The
 *     page is now scanned in 50-row pages (max 5) and the union is ordered
 *     newest-first before `top` is applied.
 */
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { OutlookClient, setOutlookClient } from "../src/outlook/client.js";
import { mailTool } from "../src/tools/mail.js";
import type { FetchLike } from "../src/http/client.js";

const BASE = "https://outlook.office.com/api/v2.0";

interface Call {
  method: string;
  url: string;
}

const row = (id: string, received: string) => ({
  Id: id,
  ConversationId: `conv-${id}`,
  Subject: id,
  From: { EmailAddress: { Address: "sender@example.com" } },
  ReceivedDateTime: received,
  BodyPreview: "p",
});

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

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-mail-filter-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  resetConfigCache();
});

afterEach(async () => {
  setOutlookClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("mail.list with from", () => {
  it("scans without $orderby and returns the newest top rows", async () => {
    const calls = harness(() => ({
      value: [
        row("older", "2026-11-15T12:00:00Z"),
        row("newest", "2026-11-16T18:00:00Z"),
        row("middle", "2026-11-16T09:00:00Z"),
      ],
    }));
    const result = await mailTool.handle({
      op: "list",
      folder: "inbox",
      from: "sender@example.com",
      top: 3,
    });
    expect(result.isError).toBeUndefined();
    expect(calls).toHaveLength(1);
    const query = new URL(calls[0]?.url ?? "").searchParams;
    expect(query.get("$filter")).toBe("From/EmailAddress/Address eq 'sender@example.com'");
    expect(query.get("$orderby")).toBeNull();
    // The scan page size is independent of the caller's small `top`.
    expect(query.get("$top")).toBe("50");
    const out = payload(result);
    const rows = out.messages as Array<Record<string, unknown>>;
    expect(rows.map((entry) => entry.id)).toEqual(["newest", "middle", "older"]);
    expect(out.scanned_pages).toBe(1);
    expect(out.truncated).toBe(false);
    expect(out.next_cursor).toBeUndefined();
  });

  it("keeps $orderby for a list without from", async () => {
    const calls = harness(() => ({ value: [row("only", "2026-11-16T12:00:00Z")] }));
    const result = await mailTool.handle({ op: "list", top: 1, unread_only: true });
    expect(result.isError).toBeUndefined();
    const query = new URL(calls[0]?.url ?? "").searchParams;
    expect(query.get("$filter")).toBe("IsRead eq false");
    expect(query.get("$orderby")).toBe("ReceivedDateTime desc");
    expect(query.get("$top")).toBe("1");
  });

  it("combines from with unread_only without $orderby", async () => {
    const calls = harness(() => ({ value: [row("only", "2026-11-16T12:00:00Z")] }));
    const result = await mailTool.handle({
      op: "list",
      from: "sender@example.com",
      unread_only: true,
    });
    expect(result.isError).toBeUndefined();
    const query = new URL(calls[0]?.url ?? "").searchParams;
    expect(query.get("$filter")).toBe(
      "IsRead eq false and From/EmailAddress/Address eq 'sender@example.com'",
    );
    expect(query.get("$orderby")).toBeNull();
    expect(calls[0]?.url.startsWith(`${BASE}/me/mailfolders/inbox/messages?`)).toBe(true);
  });

  it("follows nextLink so a tail page cannot hide the newest match", async () => {
    const page2 = `${BASE}/me/mailfolders/inbox/messages?$skip=2&$top=50`;
    const calls = harness((url) =>
      url === page2
        ? {
            value: [
              row("newest", "2026-11-16T18:00:00Z"),
              row("second", "2026-11-16T09:00:00Z"),
            ],
          }
        : {
            value: [
              row("older", "2026-11-15T12:00:00Z"),
              row("oldest", "2025-01-13T12:00:00Z"),
            ],
            "@odata.nextLink": page2,
          },
    );
    const out = payload(
      await mailTool.handle({
        op: "list",
        from: "sender@example.com",
        top: 2,
      }),
    );
    const rows = out.messages as Array<Record<string, unknown>>;
    // The regression: top=2 must be the two NEWEST, not the first scanned page.
    expect(rows.map((entry) => entry.id)).toEqual(["newest", "second"]);
    expect(calls.map((call) => call.url)).toEqual([
      expect.stringContaining(`${BASE}/me/mailfolders/inbox/messages?`),
      page2,
    ]);
    expect(out.scanned_pages).toBe(2);
    expect(out.truncated).toBe(true);
    expect(out.next_cursor).toBeUndefined();
  });

  it("stops at the 5x50 scan cap and hands back a cursor", async () => {
    const seen: string[] = [];
    const calls = harness((url) => {
      seen.push(url);
      const page = seen.length;
      return {
        value: Array.from({ length: 50 }, (_, index) =>
          row(`p${page}-${index}`, `2026-01-${String(page).padStart(2, "0")}T${String(index % 24).padStart(2, "0")}:00:00Z`),
        ),
        "@odata.nextLink": `${BASE}/me/mailfolders/inbox/messages?$skip=${page * 50}`,
      };
    });
    const out = payload(
      await mailTool.handle({
        op: "list",
        from: "sender@example.com",
        top: 1,
      }),
    );
    expect(calls).toHaveLength(5);
    expect(out.scanned_pages).toBe(5);
    expect(out.truncated).toBe(true);
    expect(typeof out.next_cursor).toBe("string");
    // Page 5 holds the latest dates in this fixture, so its newest row wins.
    expect((out.messages as Array<Record<string, unknown>>)[0]?.id).toBe("p5-23");
  });

  it("continues from a scan cursor without restarting page one", async () => {
    const start = `${BASE}/me/mailfolders/inbox/messages?$skip=250`;
    const calls = harness((url) =>
      url === start
        ? { value: [row("tail-newest", "2026-11-01T12:00:00Z")] }
        : { value: [] },
    );
    const out = payload(
      await mailTool.handle({
        op: "list",
        from: "sender@example.com",
        top: 1,
        cursor: Buffer.from(start, "utf8").toString("base64url"),
      }),
    );
    expect(calls.map((call) => call.url)).toEqual([start]);
    expect((out.messages as Array<Record<string, unknown>>)[0]?.id).toBe("tail-newest");
    expect(out.truncated).toBe(false);
  });
});
