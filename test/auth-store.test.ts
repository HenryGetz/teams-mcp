import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { config, resetConfigCache } from "../src/config.js";
import {
  AuthRequiredError,
  getAccessToken,
  isTokenFresh,
  logout,
  persistTokenUpdate,
  readTokenCache,
  writeTokenCache,
  type CacheReadResult,
  type TokenCache,
} from "../src/auth/store.js";
import { writeTeamsWebCache, type TeamsWebCache } from "../src/auth/teams-web.js";
import { authTool } from "../src/tools/auth.js";

let dir: string;
const saved = { ...process.env };
const FUTURE = new Date(Date.now() + 3_600_000).toISOString();

function loaded(result: CacheReadResult): Extract<CacheReadResult, { cache: object }> {
  if (result.cache === null) throw new Error(`expected a token cache, got: ${result.reason}`);
  return result;
}

function failed(result: CacheReadResult): Extract<CacheReadResult, { cache: null }> {
  if (result.cache !== null) throw new Error("expected no token cache");
  return result;
}

function cache(overrides: Partial<TokenCache> = {}): TokenCache {
  return {
    version: 1,
    account: { upn: "user@example.com", id: "oid-1" },
    refresh_token: "rt-1",
    tokens: {
      outlook: { access_token: "outlook-token", expires_at: FUTURE, scope: "Mail.Read Calendars.ReadWrite" },
      graph: { access_token: "graph-token", expires_at: FUTURE, scope: "User.Read.All" },
    },
    source: "refresh",
    updated_at: new Date().toISOString(),
    ...overrides,
  };
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-auth-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  delete process.env.TEAMS_MCP_CACHE_DIR;
  delete process.env.TEAMS_MCP_PROFILE_DIR;
  resetConfigCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

describe("canonical token cache", () => {
  it("round-trips the dual-audience document with 0600 permissions", async () => {
    const path = join(dir, "token.json");
    await writeTokenCache(cache(), path);
    expect((await stat(path)).mode & 0o777).toBe(0o600);

    const result = loaded(await readTokenCache());
    expect(result.format).toBe("v1");
    expect(result.cache.tokens.outlook?.access_token).toBe("outlook-token");
    expect(result.cache.tokens.graph?.access_token).toBe("graph-token");
    expect(result.cache.refresh_token).toBe("rt-1");

    await expect(getAccessToken({ audience: "outlook" })).resolves.toMatchObject({
      token: "outlook-token",
      audience: "outlook",
      source: "refresh",
    });
    await expect(getAccessToken({ audience: "graph" })).resolves.toMatchObject({
      token: "graph-token",
      audience: "graph",
    });
  });

  it("defaults to the outlook audience", async () => {
    await writeTokenCache(cache());
    await expect(getAccessToken()).resolves.toMatchObject({ audience: "outlook" });
  });

  it("reports a missing cache without throwing, and asks for a login", async () => {
    const result = failed(await readTokenCache());
    expect(result.reason).toContain("no token cache");
    expect(result.corrupt).toBeUndefined();
    // Inject the login hook: the real one drives a browser, which tests must not do.
    let loginReason: string | undefined;
    await expect(
      getAccessToken({
        deps: {
          login: async (request) => {
            loginReason = request.reason;
            throw new AuthRequiredError("no session", "run the login flow");
          },
        },
      }),
    ).rejects.toMatchObject({ code: "auth_required" });
    // A missing cache is a *login* case, not corruption: the hook must have run.
    expect(loginReason).toContain("no token cache");
  });

  it("refuses to start a browser for a corrupt cache and says what to delete", async () => {
    const path = config().tokenPath;
    await writeFile(path, "{ this is not json");
    const result = failed(await readTokenCache());
    expect(result.corrupt).toBe(true);
    expect(result.reason).toContain("invalid:");

    let loginCalled = false;
    const error = await getAccessToken({
      deps: {
        login: async () => {
          loginCalled = true;
          throw new AuthRequiredError("should not run", "should not run");
        },
      },
    }).catch((caught: unknown) => caught);
    expect(loginCalled).toBe(false);
    expect(error).toMatchObject({ code: "auth_required" });
    expect(String((error as Error).message)).toContain("unreadable");
    expect(String((error as { hint?: string }).hint)).toContain("delete");
    // The file is left for the operator; nothing is invented or overwritten.
    expect(await readFile(path, "utf8")).toBe("{ this is not json");
  });

  it("asks for a login when the outlook token is expired", async () => {
    await writeTokenCache(
      cache({
        tokens: {
          outlook: { access_token: "old", expires_at: new Date(Date.now() - 60_000).toISOString() },
          graph: { access_token: "graph-token", expires_at: FUTURE },
        },
      }),
    );
    expect(isTokenFresh(new Date(Date.now() - 60_000).toISOString())).toBe(false);
    await expect(
      getAccessToken({
        audience: "outlook",
        deps: {
          refresh: async () => {
            throw new AuthRequiredError("refresh token revoked", "sign in again");
          },
          login: async () => {
            throw new AuthRequiredError("no session", "run the login flow");
          },
        },
      }),
    ).rejects.toMatchObject({ code: "auth_required" });
  });

  it("reports a missing audience distinctly from an expired token", async () => {
    await writeTokenCache(cache({ tokens: { outlook: { access_token: "o", expires_at: FUTURE } } }));
    await expect(
      getAccessToken({
        audience: "graph",
        deps: {
          refresh: async ({ cache: current }) => current,
          login: async () => {
            throw new AuthRequiredError("no session", "run the login flow");
          },
        },
      }),
    ).rejects.toMatchObject({ code: "auth_required" });
  });

  it("persists a rotated refresh token without dropping the other audience", async () => {
    await writeTokenCache(cache());
    const updated = await persistTokenUpdate({
      audience: "graph",
      access_token: "graph-token-2",
      expires_at: new Date(Date.now() + 7_200_000).toISOString(),
      refresh_token: "rt-2",
      source: "refresh",
    });
    expect(updated.tokens.graph?.access_token).toBe("graph-token-2");
    expect(updated.tokens.outlook?.access_token).toBe("outlook-token");
    expect(updated.refresh_token).toBe("rt-2");

    const onDisk = JSON.parse(await readFile(join(dir, "token.json"), "utf8")) as TokenCache;
    expect(onDisk.refresh_token).toBe("rt-2");
    expect(onDisk.tokens.outlook?.access_token).toBe("outlook-token");
    expect(onDisk.tokens.graph?.scope).toBe("User.Read.All");
  });

  it("deletes the cached token but not the browser profile path", async () => {
    await writeTokenCache(cache());
    await expect(logout()).resolves.toMatchObject({ removed: true });
    expect((await readTokenCache()).cache).toBeNull();
  });

  it("uses browser-profile as the default profile dir", () => {
    expect(config().profileDir.endsWith("browser-profile")).toBe(true);
  });

  it("still honours an explicit test path", async () => {
    const explicit = join(dir, "other.json");
    await writeTokenCache(cache({ tokens: { outlook: { access_token: "explicit", expires_at: FUTURE } } }), explicit);
    expect(loaded(await readTokenCache(explicit)).cache.tokens.outlook?.access_token).toBe("explicit");
  });
});

describe("legacy documents (read-only)", () => {
  it("normalises an OWA/MSAL capture and infers the graph audience from aud", async () => {
    await writeFile(
      join(dir, "tokens.json"),
      JSON.stringify({
        obtainedAt: new Date().toISOString(),
        obtainedFrom: "owa-msal-refresh-token",
        accessToken: "owa-token",
        refreshToken: "owa-rt",
        expiresAt: Date.now() + 3_600_000,
        scopes: ["Mail.Read"],
        claims: { aud: "https://graph.microsoft.com", upn: "user@example.com", oid: "oid-1" },
      }),
    );
    const result = loaded(await readTokenCache());
    expect(result.format).toBe("legacy");
    expect(result.cache.tokens.graph?.access_token).toBe("owa-token");
    expect(result.cache.tokens.outlook).toBeUndefined();
    expect(result.cache.account?.upn).toBe("user@example.com");
    await expect(getAccessToken({ audience: "graph" })).resolves.toMatchObject({ token: "owa-token" });
  });

  it("normalises the old single-audience v1 document as the outlook audience", async () => {
    await writeFile(
      join(dir, "token.json"),
      JSON.stringify({
        version: 1,
        source: "headless",
        access_token: "legacy-outlook",
        expires_at: FUTURE,
        refresh_token: "legacy-rt",
        account: { upn: "user@example.com" },
        updated_at: new Date().toISOString(),
      }),
    );
    const result = loaded(await readTokenCache());
    expect(result.format).toBe("legacy");
    expect(result.cache.tokens.outlook?.access_token).toBe("legacy-outlook");
    await expect(getAccessToken()).resolves.toMatchObject({ token: "legacy-outlook" });
  });

  it("prefers the canonical document over a legacy capture", async () => {
    await writeFile(join(dir, "tokens.json"), JSON.stringify({ accessToken: "legacy", expiresAt: Date.now() + 3_600_000 }));
    await writeTokenCache(cache());
    const result = loaded(await readTokenCache());
    expect(result.format).toBe("v1");
    expect(result.cache.tokens.outlook?.access_token).toBe("outlook-token");
  });
});

describe("auth.status semantics", () => {
  const PAST = new Date(Date.now() - 60_000).toISOString();

  async function status() {
    const result = await authTool.handle({ op: "status" });
    return JSON.parse((result.content[0] as { text: string }).text) as Record<string, unknown>;
  }

  function webCache(expiresAt: string): TeamsWebCache {
    return {
      version: 1,
      client_id: "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
      account: { upn: "user@example.com", id: "oid-1" },
      scope: "https://graph.microsoft.com/.default",
      access_token: "web-token",
      expires_at: expiresAt,
      source: "browser",
      updated_at: new Date().toISOString(),
    };
  }

  it("reports authenticated when only the Teams-web cache is fresh", async () => {
    // No token.json at all: channel history still works with this credential.
    await writeTeamsWebCache(webCache(FUTURE));
    const body = await status();
    expect(body.authenticated).toBe(true);
    expect(body.fresh_audiences).toEqual(["teams_web"]);
    expect(body.reason).toBeUndefined();
    expect(body.token_cache).toBeTruthy();
    expect(body.token_path).toBeUndefined();
    expect(body.teams_web).toMatchObject({
      authenticated: true,
      fresh: true,
      source: "browser",
      refreshable: false,
    });
  });

  it("reports authenticated when only the graph token is fresh", async () => {
    await writeTokenCache(
      cache({
        tokens: {
          outlook: { access_token: "outlook-token", expires_at: PAST, scope: "Mail.Read" },
          graph: { access_token: "graph-token", expires_at: FUTURE, scope: "User.Read" },
        },
      }),
    );
    const body = await status();
    expect(body.authenticated).toBe(true);
    expect(body.fresh_audiences).toEqual(["graph"]);
    expect(body.reason).toBeUndefined();
    const audiences = body.audiences as Array<Record<string, unknown>>;
    expect(audiences.map((entry) => [entry.audience, entry.fresh])).toEqual([
      ["outlook", false],
      ["graph", true],
    ]);
  });

  it("reports unauthenticated with a reason only when no audience is fresh", async () => {
    await writeTokenCache(
      cache({
        tokens: {
          outlook: { access_token: "outlook-token", expires_at: PAST, scope: "Mail.Read" },
          graph: { access_token: "graph-token", expires_at: PAST, scope: "User.Read" },
        },
      }),
    );
    const body = await status();
    expect(body.authenticated).toBe(false);
    expect(body.fresh_audiences).toEqual([]);
    expect(String(body.reason)).toContain("auth.login");
    expect(body.teams_web).toMatchObject({ authenticated: false, fresh: false, reason: "missing" });
  });

  it("ignores a stale Teams-web cache when the graph token is fresh", async () => {
    await writeTokenCache(cache());
    await writeTeamsWebCache(webCache(PAST));
    const body = await status();
    expect(body.authenticated).toBe(true);
    expect(body.fresh_audiences).toEqual(["outlook", "graph"]);
    expect(body.teams_web).toMatchObject({ authenticated: true, fresh: false });
  });

  it("never marks a corrupt Teams-web cache as fresh or authenticated", async () => {
    await writeFile(join(dir, "teams-web-token.json"), "{ not json");
    const body = await status();
    expect(body.authenticated).toBe(false);
    expect(body.fresh_audiences).toEqual([]);
    expect(body.teams_web).toMatchObject({ authenticated: false, fresh: false, corrupt: true });
    // The parse error must not echo cache bytes.
    expect(JSON.stringify(body)).not.toContain("not json");
  });
});

const realTokenPath = join(process.env.HOME ?? "", ".config", "teams-mcp", "token.json");

describe.skipIf(!existsSync(realTokenPath))("auth worker's token.json", () => {
  it("is accepted by the reader", async () => {
    const result = loaded(await readTokenCache(realTokenPath));
    expect(result.format).toBe("v1");
    expect(Object.keys(result.cache.tokens).length).toBeGreaterThan(0);
    expect(result.cache.account?.upn).toBeTruthy();
  });
});
