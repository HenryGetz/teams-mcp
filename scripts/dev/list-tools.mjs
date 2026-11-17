#!/usr/bin/env node
/**
 * Smoke test: spawn the built server over stdio, initialize, and print the tool
 * and resource inventory. Fails loudly if the handshake or tools/list breaks.
 *
 *   npm run build && npm run smoke
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const transport = new StdioClientTransport({
  command: process.execPath,
  args: ["dist/index.js"],
  env: { ...process.env },
  stderr: "inherit",
});

const client = new Client({ name: "teams-mcp-smoke", version: "0.1.0" });

try {
  await client.connect(transport);
  const info = client.getServerVersion();
  console.log(`initialized: ${info?.name} ${info?.version}`);

  const { tools } = await client.listTools();
  console.log(`tools: ${tools.length}`);
  for (const tool of tools) {
    const op = tool.inputSchema?.properties?.op?.enum ?? [];
    console.log(`  - ${tool.name}: op=[${op.join(", ")}]`);
  }

  const { resources } = await client.listResources();
  console.log(`resources: ${resources.length}`);
  for (const resource of resources) console.log(`  - ${resource.uri} (${resource.mimeType ?? "text/plain"})`);

  const status = await client.callTool({ name: "auth", arguments: { op: "status" } });
  console.log(`auth.status -> ${status.content?.[0]?.text ?? ""}`);

  const unimplemented = await client.callTool({ name: "mail", arguments: { op: "send" } });
  console.log(`mail.send (expected error) -> isError=${unimplemented.isError ?? false}`);

  const readOnly = await client.callTool({ name: "people", arguments: { op: "search" } });
  console.log(`people.search without query -> isError=${readOnly.isError ?? false}`);

  const settings = await client.callTool({ name: "settings", arguments: { op: "path" } });
  console.log(`settings.path -> ${settings.content?.[0]?.text ?? ""}`);
} finally {
  await client.close();
}
