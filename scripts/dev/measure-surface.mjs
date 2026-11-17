#!/usr/bin/env node
/**
 * Token-surface measurement: what an MCP client pays to load this server, and
 * what four representative calls return.
 *
 *   node scripts/dev/measure-surface.mjs
 *
 * Prints byte counts (UTF-8) and a bytes/4 token estimate. Read-only: every
 * probe call is a list/search/suggest; nothing is written to Outlook.
 * Uses the ambient config/token (set TEAMS_MCP_CONFIG_DIR to measure another).
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const bytes = (value) => Buffer.byteLength(typeof value === "string" ? value : JSON.stringify(value), "utf8");
const tokensOf = (byteCount) => Math.round(byteCount / 4);
const pad = (value, width) => String(value).padEnd(width);

const transport = new StdioClientTransport({
  command: process.execPath,
  args: ["dist/index.js"],
  env: { ...process.env },
  stderr: "inherit",
});
const client = new Client({ name: "teams-mcp-measure", version: "0.1.0" });

const probes = [
  ["mail.list top=10", { name: "mail", arguments: { op: "list", top: 10 } }],
  ["calendar.list week", { name: "calendar", arguments: { op: "list", top: 50 } }],
  ["availability.suggest", { name: "availability", arguments: { op: "suggest" } }],
  ["people.search Student", { name: "people", arguments: { op: "search", query: "Test Student" } }],
];

try {
  await client.connect(transport);
  const { tools } = await client.listTools();
  const { resources } = await client.listResources();

  const total = bytes(tools);
  console.log(`tools/list: ${tools.length} tools, ${total} bytes ≈ ${tokensOf(total)} tokens`);
  console.log(`resources/list: ${resources.length}, ${bytes(resources)} bytes ≈ ${tokensOf(bytes(resources))} tokens`);
  console.log("");
  console.log(`${pad("tool", 12)} ${pad("total", 7)} ${pad("desc", 7)} ${pad("schema", 7)} ops`);
  for (const tool of tools) {
    const opCount = tool.inputSchema?.properties?.op?.enum?.length ?? 0;
    const description = tool.description ?? "";
    const schema = { ...tool, description: "" };
    console.log(
      `${pad(tool.name, 12)} ${pad(bytes(tool), 7)} ${pad(bytes(description), 7)} ${pad(bytes(schema), 7)} ${opCount}`,
    );
  }

  console.log("");
  console.log(`${pad("probe", 22)} ${pad("bytes", 8)} ${pad("tokens", 7)} result`);
  for (const [label, call] of probes) {
    const started = Date.now();
    const result = await client.callTool(call);
    const text = result.content?.map((entry) => (entry.type === "text" ? entry.text : "")).join("") ?? "";
    const error = result.isError ? " ERROR" : "";
    console.log(
      `${pad(label, 22)} ${pad(bytes(text), 8)} ${pad(tokensOf(bytes(text)), 7)} ${Date.now() - started}ms${error}`,
    );
    if (result.isError) console.log(text.slice(0, 200));
  }
} finally {
  await client.close().catch(() => undefined);
}
