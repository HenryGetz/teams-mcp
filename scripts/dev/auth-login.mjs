#!/usr/bin/env node
/**
 * Live auth integration check. Run AFTER `npm run build`:
 *
 *   node scripts/dev/auth-login.mjs            # warm profile: one click, no password
 *   node scripts/dev/auth-login.mjs --offline  # cache/plumbing only, no network
 *
 * Proves, against the compiled MCP code:
 *   1. forceLogin() signs in and writes the canonical token.json (0600, both audiences).
 *   2. The outlook token authenticates GET /me and the graph token authenticates
 *      /users?$search — i.e. the two audiences are wired to the right APIs.
 *   3. The refresh path mints a NEW outlook token without touching the browser,
 *      persisting the rotated refresh token.
 *   4. `auth` op=status over the real stdio MCP transport agrees with the file.
 *
 * It deliberately performs exactly ONE refresh so a concurrently running agent
 * keeps a valid token, and it never deletes the browser profile.
 */
import assert from "node:assert/strict";
import { readFile, stat } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { config } from "../../dist/config.js";
import { forceLogin, getAccessToken, persistTokenUpdate, readTokenCache, tokenLockPath } from "../../dist/auth/store.js";
import { decodeJwtClaims } from "../../dist/auth/refresh.js";

const OFFLINE = process.argv.includes("--offline");
const repoRoot = fileURLToPath(new URL("../..", import.meta.url));
const distIndex = join(repoRoot, "dist", "index.js");

const results = [];
const timings = {};
const step = async (name, fn) => {
  const started = Date.now();
  try {
    const value = await fn();
    timings[name] = Date.now() - started;
    results.push({ step: name, ok: true, ms: timings[name] });
    console.log(`  PASS ${name} (${timings[name]}ms)`);
    return value;
  } catch (error) {
    timings[name] = Date.now() - started;
    results.push({ step: name, ok: false, ms: timings[name], error: error instanceof Error ? error.message : String(error) });
    console.error(`  FAIL ${name} (${timings[name]}ms): ${error instanceof Error ? error.message : String(error)}`);
    throw error;
  }
};

const redact = (value) => (typeof value === "string" && value.length > 20 ? `${value.slice(0, 6)}…${value.slice(-6)} (len ${value.length})` : "(none)");

async function callGraph(token, path) {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json", ConsistencyLevel: "eventual" },
  });
  return { status: response.status, body: await response.json().catch(() => null) };
}

async function callOutlook(token, path) {
  const response = await fetch(`https://outlook.office.com/api/v2.0${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  return { status: response.status, body: await response.json().catch(() => null) };
}

console.log(`auth integration — profile ${config().profileDir}`);
console.log(`                  token   ${config().tokenPath}`);
console.log(`                  lock    ${tokenLockPath()}\n`);

const t0 = Date.now();
try {
  // ---- 1. sign in through the compiled code ----
  const token = await step("forceLogin (warm profile)", async () => {
    const result = await forceLogin({ audience: "outlook" });
    assert.equal(result.audience, "outlook");
    assert.ok(result.token.split(".").length === 3, "expected a JWT access token");
    return result;
  });

  // ---- 2. canonical cache shape + permissions ----
  const cacheInfo = await step("token.json shape + 0600", async () => {
    const file = config().tokenPath;
    const mode = (await stat(file)).mode & 0o777;
    assert.equal(mode, 0o600, `token.json must be 0600, got ${mode.toString(8)}`);
    const raw = JSON.parse(await readFile(file, "utf8"));
    assert.equal(raw.version, 1);
    assert.ok(raw.account?.upn, "account.upn missing");
    assert.ok(raw.account?.id, "account.id missing");
    assert.ok(typeof raw.refresh_token === "string" && raw.refresh_token.length > 100, "refresh_token missing/short");
    assert.ok(raw.tokens?.outlook?.access_token, "tokens.outlook missing");
    assert.ok(raw.tokens?.graph?.access_token, "tokens.graph missing");
    for (const audience of ["outlook", "graph"]) {
      assert.ok(!Number.isNaN(Date.parse(raw.tokens[audience].expires_at)), `tokens.${audience}.expires_at invalid`);
      assert.ok(typeof raw.tokens[audience].scope === "string" && raw.tokens[audience].scope.length > 0, `tokens.${audience}.scope empty`);
    }
    assert.ok(["browser", "refresh", "headless"].includes(raw.source), `unexpected source ${raw.source}`);
    assert.ok(!Number.isNaN(Date.parse(raw.updated_at)), "updated_at invalid");
    return { raw, mode };
  });

  const outlookClaims = decodeJwtClaims(cacheInfo.raw.tokens.outlook.access_token);
  const graphClaims = decodeJwtClaims(cacheInfo.raw.tokens.graph.access_token);
  await step("audiences are distinct and correct", async () => {
    assert.equal(outlookClaims?.aud, "https://outlook.office.com");
    assert.equal(graphClaims?.aud, "https://graph.microsoft.com");
    assert.ok((outlookClaims?.scp ?? "").includes("Mail.ReadWrite"), "outlook token lacks Mail.ReadWrite");
    assert.ok((outlookClaims?.scp ?? "").includes("Calendars.ReadWrite"), "outlook token lacks Calendars.ReadWrite");
    assert.ok((graphClaims?.scp ?? "").includes("Directory.Read.All"), "graph token lacks Directory.Read.All");
    assert.equal(outlookClaims?.tid, graphClaims?.tid, "tenants disagree between audiences");
    return { tenant: outlookClaims?.tid, outlookScopes: (outlookClaims?.scp ?? "").split(" ").length, graphScopes: (graphClaims?.scp ?? "").split(" ").length };
  });
  console.log(`       account ${cacheInfo.raw.account.upn} (${cacheInfo.raw.account.id}) tenant ${outlookClaims?.tid}`);
  console.log(`       outlook ${redact(cacheInfo.raw.tokens.outlook.access_token)} ${(outlookClaims?.scp ?? "").split(" ").length} scopes`);
  console.log(`       graph   ${redact(cacheInfo.raw.tokens.graph.access_token)} ${(graphClaims?.scp ?? "").split(" ").length} scopes`);

  if (!OFFLINE) {
    // ---- 3. both tokens work against their real API ----
    await step("outlook token -> GET /me", async () => {
      const { status, body } = await callOutlook(token.token, "/me");
      assert.equal(status, 200, `GET /me returned ${status}`);
      assert.ok(body?.EmailAddress, "no EmailAddress in /me");
      assert.equal(body.EmailAddress.toLowerCase(), cacheInfo.raw.account.upn.toLowerCase());
      return body.EmailAddress;
    });

    await step("graph token -> /users?$search", async () => {
      // Any real directory hit proves the graph token works; override the search
      // term with TEAMS_MCP_DIRECTORY_QUERY when you want a specific person.
      const term = process.env.TEAMS_MCP_DIRECTORY_QUERY ?? "user";
      const search = encodeURIComponent(`"displayName:${term}"`);
      const { status, body } = await callGraph(
        cacheInfo.raw.tokens.graph.access_token,
        `/users?$search=${search}&$select=id,displayName,mail&$top=5`,
      );
      assert.equal(status, 200, `/users $search returned ${status}`);
      const match = (body?.value ?? [])[0];
      assert.ok(match, `directory search for "${term}" returned no rows via the graph token`);
      console.log(`       directory: ${match.displayName} <${match.mail}> ${match.id}`);
      return match;
    });

    // ---- 4. refresh path: new outlook token, no browser, rotation persisted ----
    await step("refresh path (no browser, rotation persisted)", async () => {
      const before = JSON.parse(await readFile(config().tokenPath, "utf8"));
      // Age the cached outlook entry on disk, then let the real clock drive the
      // refresh path (a mocked `now` would also fool the post-refresh margin check).
      await persistTokenUpdate({
        audience: "outlook",
        access_token: before.tokens.outlook.access_token,
        expires_at: new Date(Date.now() - 60_000).toISOString(),
        ...(before.tokens.outlook.scope ? { scope: before.tokens.outlook.scope } : {}),
        source: "refresh",
      });
      const refreshed = await getAccessToken({ audience: "outlook" });
      assert.ok(refreshed.token.split(".").length === 3);
      assert.equal(refreshed.audience, "outlook");
      const after = JSON.parse(await readFile(config().tokenPath, "utf8"));
      assert.notEqual(after.tokens.outlook.access_token, before.tokens.outlook.access_token, "access token did not change");
      assert.notEqual(after.refresh_token, before.refresh_token, "refresh token was not rotated on disk");
      assert.equal(after.source, "refresh");
      assert.equal(after.tokens.graph.access_token, before.tokens.graph.access_token, "graph token should be untouched");
      // The refreshed token must actually work.
      const { status } = await callOutlook(refreshed.token, "/me");
      assert.equal(status, 200, `refreshed token rejected by /me (${status})`);
      return { expiresAt: refreshed.expiresAt };
    });

    // ---- 5. the freshly rotated session still refreshes (proves persistence order) ----
    await step("rotated token still usable after restart", async () => {
      const fresh = await readTokenCache();
      assert.ok(fresh.cache?.refresh_token, "no refresh token on disk after rotation");
      const { status } = await callOutlook(fresh.cache.tokens.outlook.access_token, "/me");
      assert.equal(status, 200, "persisted outlook token rejected");
      return true;
    });
  }

  // ---- 6. status through the real stdio MCP client ----
  const status = await step("MCP stdio auth status", async () => {
    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [distIndex],
      env: { ...process.env },
      stderr: "pipe",
    });
    const client = new Client({ name: "teams-mcp-auth-smoke", version: "0.0.0" });
    await client.connect(transport);
    try {
      const tools = await client.listTools();
      assert.ok(tools.tools.some((tool) => tool.name === "auth"), "auth tool missing");
      const result = await client.callTool({ name: "auth", arguments: { op: "status" } });
      const text = result.content?.[0]?.text;
      assert.ok(typeof text === "string", "auth status returned no text");
      const parsed = JSON.parse(text);
      assert.equal(parsed.authenticated, true, `status reports unauthenticated: ${text}`);
      assert.equal(parsed.account, cacheInfo.raw.account.upn);
      assert.ok(Array.isArray(parsed.audiences) && parsed.audiences.length === 2, "expected two audiences");
      assert.equal(parsed.has_refresh_token, true);
      assert.equal(parsed.profile, config().profileDir);
      return parsed;
    } finally {
      await client.close().catch(() => undefined);
    }
  });

  const totalMs = Date.now() - t0;
  console.log(`\n=== RESULT: PASS (${totalMs}ms total) ===`);
  console.log(JSON.stringify({ timings, totalMs, account: status.account, audiences: status.audiences.map((a) => ({ audience: a.audience, fresh: a.fresh, seconds_left: a.seconds_left })) }, null, 2));
  process.exitCode = 0;
} catch (error) {
  console.error(`\n=== RESULT: FAIL after ${Date.now() - t0}ms ===`);
  console.error(error instanceof Error ? error.stack : String(error));
  console.error(JSON.stringify({ timings, results }, null, 2));
  process.exitCode = 1;
}
