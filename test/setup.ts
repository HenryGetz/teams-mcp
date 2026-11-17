/**
 * Test isolation defaults.
 *
 * The write-audit journal and unsubscribe counter live under the cache dir. Unit
 * tests must never touch the operator's real `~/.cache/teams-mcp`, so give every
 * worker a private cache dir unless the test file already chose one.
 */
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

process.env.TEAMS_MCP_CACHE_DIR ??= mkdtempSync(join(tmpdir(), "teams-mcp-test-cache-"));
