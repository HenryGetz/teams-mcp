import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { resetConfigCache } from "../src/config.js";
import { GraphClient, setGraphClient } from "../src/graph/client.js";
import { clearPeopleCache, peopleTool } from "../src/tools/people.js";
import type { FetchLike } from "../src/http/client.js";

const BASE = "https://graph.microsoft.com/v1.0";

interface Call {
  method: string;
  url: string;
}

function harness(responses: Array<() => Response>) {
  const calls: Call[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({ method: init?.method ?? "GET", url });
    const factory = responses[Math.min(index, responses.length - 1)];
    index += 1;
    if (!factory) throw new Error("no response queued");
    return factory();
  };
  setGraphClient(
    new GraphClient({ tokenProvider: async () => ({ token: "t" }), fetchImpl, maxAttempts: 1 }),
  );
  return calls;
}

const json = (body: unknown, status = 200): (() => Response) => () =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

function payload(result: CallToolResult): Record<string, unknown> {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("expected a text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

const joel = {
  id: "id-joel",
  displayName: "Test Student",
  mail: "joel@example.edu",
  userPrincipalName: "joel@example.edu",
  jobTitle: "Student",
  department: "Computer Science",
};
const joelTwin = {
  id: "id-joel-2",
  displayName: "Test Student (alumni)",
  mail: "joel.alumni@example.com",
  userPrincipalName: "joel.alumni@example.com",
  jobTitle: "Alumnus",
};

let dir: string;
const saved = { ...process.env };

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-people-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  resetConfigCache();
  clearPeopleCache();
});

afterEach(async () => {
  setGraphClient(undefined);
  process.env = { ...saved };
  resetConfigCache();
  clearPeopleCache();
  await rm(dir, { recursive: true, force: true });
});

describe("people.search", () => {
  it("returns every candidate and flags ambiguity instead of picking one", async () => {
    const calls = harness([json({ value: [joel, joelTwin] })]);
    const result = await peopleTool.handle({ op: "search", query: "student" });
    expect(result.isError).toBeUndefined();

    const out = payload(result);
    expect(out.ambiguous).toBe(true);
    expect(out.needs_disambiguation).toContain("multiple directory matches");
    expect(out.people).toEqual([
      {
        id: "id-joel",
        name: "Test Student",
        email: "joel@example.edu",
        title: "Student",
        department: "Computer Science",
      },
      {
        id: "id-joel-2",
        name: "Test Student (alumni)",
        email: "joel.alumni@example.com",
        title: "Alumnus",
      },
    ]);

    expect(calls[0]?.url).toContain(`${BASE}/users?`);
    expect(calls[0]?.url).toContain("%24search=");
    expect(calls[0]?.url).toContain("department");
  });

  it("does not flag a single match", async () => {
    harness([json({ value: [joel] })]);
    const out = payload(await peopleTool.handle({ op: "search", query: "Student" }));
    expect(out.ambiguous).toBeUndefined();
    expect(out.needs_disambiguation).toBeUndefined();
    expect(out.people).toEqual([expect.objectContaining({ id: "id-joel" })]);
  });

  it("resolves a configured alias exactly, bypassing search", async () => {
    await writeFile(join(dir, "config.yaml"), "aliases: {joel: joel@example.edu}\n");
    resetConfigCache();
    const calls = harness([json(joel)]);
    const out = payload(await peopleTool.handle({ op: "search", query: "JOEL" }));
    expect(out.alias).toBe(true);
    expect(out.resolved_to).toBe("joel@example.edu");
    expect(out.ambiguous).toBeUndefined();
    expect(calls[0]?.url).toContain(`/users/${encodeURIComponent("joel@example.edu")}?`);
  });

  it("caches repeat searches and allows a refresh", async () => {
    const calls = harness([json({ value: [joel] })]);
    await peopleTool.handle({ op: "search", query: "student" });
    const cached = payload(await peopleTool.handle({ op: "search", query: "student" }));
    expect(cached.cached).toBe(true);
    expect(calls).toHaveLength(1);
    await peopleTool.handle({ op: "search", query: "student", refresh: true });
    expect(calls).toHaveLength(2);
  });
});

describe("people.resolve", () => {
  it("reports found=false for a 404 without failing the batch", async () => {
    harness([
      json(joel),
      json({ error: { code: "Request_ResourceNotFound", message: "not found" } }, 404),
    ]);
    const out = payload(
      await peopleTool.handle({
        op: "resolve",
        emails: ["joel@example.edu", "ghost@example.com"],
      }),
    );
    const rows = out.people as Array<Record<string, unknown>>;
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ input: "joel@example.edu", found: true, id: "id-joel" });
    expect(rows[1]).toEqual({ input: "ghost@example.com", found: false });
  });

  it("requires at least one input", async () => {
    harness([]);
    const result = await peopleTool.handle({ op: "resolve" });
    expect(result.isError).toBe(true);
  });
});
