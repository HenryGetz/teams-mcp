#!/usr/bin/env node
/**
 * stdio entrypoint for teams-mcp.
 *
 * stdout carries MCP frames only; every diagnostic goes to stderr.
 */
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { config } from "./config.js";
import { registerResources, registerTools } from "./tools/registry.js";

export const SERVER_NAME = "teams-mcp";
export const SERVER_VERSION = "0.1.0";

function log(message: string): void {
  process.stderr.write(`${message}\n`);
}

export async function main(): Promise<void> {
  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    // Resources are declared too: the snapshot resource exists, and a client that
    // respects the advertised capabilities would otherwise never list it.
    { capabilities: { tools: {}, resources: {} } },
  );

  const tools = registerTools(server);
  const resources = registerResources(server);
  const transport = new StdioServerTransport();
  await server.connect(transport);

  const cfg = config();
  log(
    `${SERVER_NAME} ${SERVER_VERSION} ready — ${tools.length} tools (${tools.join(", ")}); ${resources.length} resource; config ${cfg.configDir}; tz ${cfg.timezone}`,
  );

  let closing = false;
  const shutdown = async (signal: string): Promise<void> => {
    if (closing) return;
    closing = true;
    log(`${signal} received, closing ${SERVER_NAME}`);
    try {
      await server.close();
    } catch (error) {
      log(`shutdown error: ${(error as Error).message}`);
      process.exitCode = 1;
    }
  };

  for (const signal of ["SIGINT", "SIGTERM"] as const) {
    process.once(signal, () => {
      void shutdown(signal).finally(() => process.exit(process.exitCode ?? 0));
    });
  }

  // A client closing the pipe is a normal shutdown.
  process.stdin.once("end", () => {
    void shutdown("stdin end");
  });
}

main().catch((error: unknown) => {
  log(`fatal: ${error instanceof Error ? (error.stack ?? error.message) : String(error)}`);
  process.exit(1);
});
