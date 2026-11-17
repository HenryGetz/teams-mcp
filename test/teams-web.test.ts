/**
 * Teams-web credential: MSAL v5 decryption, cache validation, refresh ordering
 * and the offline status shape.
 *
 * The decryption fixture is generated with the same scheme the browser uses
 * (HKDF over the cookie key, AES-GCM with a zero IV), so a regression in the
 * derivation is caught without a live Teams session.
 */
import { mkdtemp, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomBytes, randomUUID } from "node:crypto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { authTool } from "../src/tools/auth.js";
import { resetConfigCache } from "../src/config.js";
import {
  decryptMsalCacheEntry,
  extractTeamsWebCredential,
  getTeamsWebToken,
  isTeamsWebTokenFresh,
  parseEncryptionCookie,
  parseTeamsMsalStorage,
  readTeamsWebCache,
  TEAMS_WEB_CLIENT_ID,
  teamsWebStatus,
  writeTeamsWebCache,
  type TeamsWebCache,
} from "../src/auth/teams-web.js";
import { AuthRequiredError } from "../src/auth/store.js";

const saved = { ...process.env };
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), "teams-mcp-web-"));
  process.env.TEAMS_MCP_CONFIG_DIR = dir;
  process.env.TEAMS_MCP_CACHE_DIR = join(dir, "cache");
  resetConfigCache();
});

afterEach(async () => {
  process.env = { ...saved };
  resetConfigCache();
  await rm(dir, { recursive: true, force: true });
});

function b64url(value: Uint8Array | string): string {
  return Buffer.from(value).toString("base64url");
}

/** Reproduce MSAL v5's browser encryption for the fixture. */
async function seal(plain: Record<string, unknown>, clientId = TEAMS_WEB_CLIENT_ID) {
  const baseKey = randomBytes(32);
  const nonce = randomBytes(16);
  const id = randomUUID();
  const subtle = globalThis.crypto.subtle;
  const hkdf = await subtle.importKey("raw", baseKey, "HKDF", false, ["deriveKey"]);
  const key = await subtle.deriveKey(
    { name: "HKDF", salt: nonce, hash: "SHA-256", info: new TextEncoder().encode(clientId) },
    hkdf,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt"],
  );
  const cipher = await subtle.encrypt(
    { name: "AES-GCM", iv: new Uint8Array(12) },
    key,
    new TextEncoder().encode(JSON.stringify(plain)),
  );
  return {
    cookie: { id, key: b64url(baseKey) },
    blob: { id, nonce: b64url(nonce), data: b64url(new Uint8Array(cipher)) },
  };
}

/** Minimal unsigned JWT; only the payload is ever read. */
function fakeJwt(claims: Record<string, unknown>): string {
  return [b64url("{}"), b64url(JSON.stringify(claims)), "sig"].join(".");
}

function validCache(overrides: Partial<TeamsWebCache> = {}): TeamsWebCache {
  return {
    version: 1,
    client_id: TEAMS_WEB_CLIENT_ID,
    tenant_id: "11111111-1111-4111-8111-111111111111",
    account: { upn: "user@example.com", id: "oid-1" },
    scope: "https://graph.microsoft.com/.default",
    access_token: "access-token",
    expires_at: new Date(Date.now() + 3_600_000).toISOString(),
    source: "refresh",
    updated_at: new Date().toISOString(),
    ...overrides,
  };
}

describe("MSAL v5 decryption", () => {
  it("round-trips a sealed credential", async () => {
    const credential = { credentialType: "RefreshToken", secret: "1.AUYA-secret", homeAccountId: "oid.tid" };
    const { cookie, blob } = await seal(credential);
    const decrypted = await decryptMsalCacheEntry(blob, cookie, TEAMS_WEB_CLIENT_ID);
    expect(decrypted).toEqual(credential);
  });

  it("refuses a cache entry encrypted under a different key id", async () => {
    const { cookie, blob } = await seal({ secret: "x" });
    await expect(
      decryptMsalCacheEntry({ ...blob, id: "other" }, cookie, TEAMS_WEB_CLIENT_ID),
    ).rejects.toThrow(/key id does not match/);
  });

  it("parses the encryption cookie from a cookie header", async () => {
    const cookie = { id: "abc", key: "k" };
    expect(parseEncryptionCookie(`a=1; msal.cache.encryption=${encodeURIComponent(JSON.stringify(cookie))}; b=2`)).toEqual(cookie);
    expect(parseEncryptionCookie("a=1")).toBeUndefined();
  });

  it("recognises plaintext and encrypted entries by key shape", () => {
    const entries = parseTeamsMsalStorage({
      [`msal.2|oid.tid|login.windows.net|accesstoken|${TEAMS_WEB_CLIENT_ID}|tid|https://graph.microsoft.com/.default|`]:
        JSON.stringify({ credentialType: "AccessToken", secret: "tok", target: "graph" }),
      [`msal.2|oid.tid|login.windows.net|refreshtoken|${TEAMS_WEB_CLIENT_ID}|||`]: JSON.stringify({
        id: "k",
        nonce: "n",
        data: "d",
      }),
      "msal.2.account.keys": '["x"]',
    });
    expect(entries).toHaveLength(2);
    expect(entries.some((entry) => entry.credentialType === "AccessToken" && entry.secret === "tok")).toBe(true);
    expect(entries.some((entry) => entry.credentialType === "refreshtoken" && entry.secret === undefined)).toBe(true);
  });
});

describe("credential extraction", () => {
  it("decrypts the refresh token and reads tenant/account from claims and homeAccountId", async () => {
    const refresh = { credentialType: "RefreshToken", secret: "1.AUYA-refresh", homeAccountId: "oid-9.tid-9" };
    const { cookie, blob } = await seal(refresh);
    const idToken = fakeJwt({ preferred_username: "user@example.com", oid: "oid-9", tid: "tid-9" });
    const graphToken = fakeJwt({
      aud: "https://graph.microsoft.com",
      scp: "ChannelMessage.Read.All ChatMessage.Send",
      exp: Math.round(Date.now() / 1000) + 3600,
      tid: "tid-9",
    });
    const storage = {
      [`msal.2|oid-9.tid-9|login.windows.net|refreshtoken|${TEAMS_WEB_CLIENT_ID}|||`]: JSON.stringify(blob),
      [`msal.2|oid-9.tid-9|login.windows.net|idtoken|${TEAMS_WEB_CLIENT_ID}|tid-9||`]: JSON.stringify({
        credentialType: "IdToken",
        secret: idToken,
      }),
      [`msal.2|oid-9.tid-9|login.windows.net|accesstoken|${TEAMS_WEB_CLIENT_ID}|tid-9|https://graph.microsoft.com/.default|`]:
        JSON.stringify({ credentialType: "AccessToken", secret: graphToken }),
    };
    const found = await extractTeamsWebCredential({
      storage,
      cookieHeader: `msal.cache.encryption=${encodeURIComponent(JSON.stringify(cookie))}`,
    });
    expect(found.refreshToken).toBe("1.AUYA-refresh");
    expect(found.tenantId).toBe("tid-9");
    expect(found.upn).toBe("user@example.com");
    expect(found.oid).toBe("oid-9");
    expect(found.accessToken).toBe(graphToken);
  });

  it("falls back to an access-only credential when the cookie is gone", async () => {
    const refresh = { credentialType: "RefreshToken", secret: "1.AUYA-refresh" };
    const { blob } = await seal(refresh);
    const graphToken = fakeJwt({
      aud: "https://graph.microsoft.com",
      scp: "ChannelMessage.Read.All",
      exp: Math.round(Date.now() / 1000) + 600,
    });
    const found = await extractTeamsWebCredential({
      storage: {
        [`msal.2|oid.tid|login.windows.net|refreshtoken|${TEAMS_WEB_CLIENT_ID}|||`]: JSON.stringify(blob),
        [`msal.2|oid.tid|login.windows.net|accesstoken|${TEAMS_WEB_CLIENT_ID}|tid|x|`]: JSON.stringify({
          credentialType: "AccessToken",
          secret: graphToken,
        }),
      },
      cookieHeader: "",
    });
    expect(found.refreshToken).toBeUndefined();
    expect(found.accessToken).toBe(graphToken);
    expect(found.expiresAt).toBeDefined();
  });

  it("ignores a Graph token without ChannelMessage.Read.All", async () => {
    const found = await extractTeamsWebCredential({
      storage: {
        [`msal.2|oid.tid|login.windows.net|accesstoken|${TEAMS_WEB_CLIENT_ID}|tid|x|`]: JSON.stringify({
          credentialType: "AccessToken",
          secret: fakeJwt({ aud: "https://graph.microsoft.com", scp: "User.Read" }),
        }),
      },
      cookieHeader: "",
    });
    expect(found.accessToken).toBeUndefined();
  });
});

describe("cache", () => {
  it("writes and reads a 0600 document, rejecting corrupt files", async () => {
    await writeTeamsWebCache(validCache());
    const read = await readTeamsWebCache();
    expect(read.cache?.access_token).toBe("access-token");
    // The bearer material must be owner-only, not merely present.
    expect((await stat(join(dir, "teams-web-token.json"))).mode & 0o777).toBe(0o600);

    await writeFile(join(dir, "teams-web-token.json"), "{ not json");
    const corrupt = await readTeamsWebCache();
    expect(corrupt.cache).toBeNull();
    expect("corrupt" in corrupt && corrupt.corrupt).toBe(true);
  });

  it("applies the five-minute expiry margin", () => {
    expect(isTeamsWebTokenFresh(new Date(Date.now() + 6 * 60_000).toISOString())).toBe(true);
    expect(isTeamsWebTokenFresh(new Date(Date.now() + 4 * 60_000).toISOString())).toBe(false);
  });

  it("reports offline status without touching the network", async () => {
    await writeTeamsWebCache(validCache());
    const status = await teamsWebStatus();
    expect(status.authenticated).toBe(true);
    expect(status.fresh).toBe(true);
    expect(status.refreshable).toBe(false);
    expect(status.scope_count).toBe(1);
  });
});

describe("corrupt cache", () => {
  it("does not echo cache bytes in the status reason", async () => {
    await writeFile(join(dir, "teams-web-token.json"), JSON.stringify({ version: "SECRET-CACHE-VALUE" }));
    const status = await teamsWebStatus();
    expect(status.authenticated).toBe(false);
    expect(status.corrupt).toBe(true);
    expect(JSON.stringify(status)).not.toContain("SECRET-CACHE-VALUE");

    await writeFile(join(dir, "teams-web-token.json"), "SECRET-CACHE-VALUE is not json");
    const unparsable = await teamsWebStatus();
    expect(unparsable.corrupt).toBe(true);
    expect(JSON.stringify(unparsable)).not.toContain("SECRET-CACHE-VALUE");

    // The same material must not leak through the error path either.
    const error = await getTeamsWebToken({ allowBrowser: false }).then(
      () => undefined,
      (failure: unknown) => failure as Error,
    );
    expect(error?.message).not.toContain("SECRET-CACHE-VALUE");
  });
});

describe("refresh transport", () => {
  const freshToken = () =>
    fakeJwt({
      aud: "https://graph.microsoft.com",
      scp: "ChannelMessage.Read.All ChatMessage.Send",
      exp: Math.round(Date.now() / 1000) + 3600,
    });

  it("uses the Teams client id, Origin and .default scope, and persists the rotation", async () => {
    await writeTeamsWebCache(
      validCache({
        expires_at: new Date(Date.now() - 60_000).toISOString(),
        refresh_token: "rt-1",
      }),
    );
    const seen: Array<{ url: string; init: RequestInit | undefined }> = [];
    const accessToken = freshToken();
    vi.stubGlobal("fetch", async (url: string, init?: RequestInit) => {
      seen.push({ url, init });
      return new Response(
        JSON.stringify({ access_token: accessToken, refresh_token: "rt-2", expires_in: 3600 }),
        { status: 200 },
      );
    });
    try {
      const token = await getTeamsWebToken();
      expect(token.token).toBe(accessToken);
      expect(token.refreshable).toBe(true);
      expect(seen).toHaveLength(1);
      expect(seen[0]?.url).toBe(
        "https://login.microsoftonline.com/11111111-1111-4111-8111-111111111111/oauth2/v2.0/token",
      );
      const form = new URLSearchParams(String(seen[0]?.init?.body));
      expect(form.get("grant_type")).toBe("refresh_token");
      expect(form.get("client_id")).toBe(TEAMS_WEB_CLIENT_ID);
      expect(form.get("scope")).toBe("https://graph.microsoft.com/.default");
      expect(form.get("refresh_token")).toBe("rt-1");
      expect((seen[0]?.init?.headers as Record<string, string>).Origin).toBe("https://teams.microsoft.com");
    } finally {
      vi.unstubAllGlobals();
    }
    // Rotation is durable: the next process must not redeem rt-1 again.
    const persisted = (await readTeamsWebCache()).cache;
    expect(persisted?.refresh_token).toBe("rt-2");
    expect(persisted?.access_token).toBe(accessToken);
    expect(persisted?.source).toBe("refresh");
  });

  it("refreshes once for two concurrent callers", async () => {
    await writeTeamsWebCache(
      validCache({
        expires_at: new Date(Date.now() - 60_000).toISOString(),
        refresh_token: "rt-1",
      }),
    );
    let count = 0;
    const accessToken = freshToken();
    vi.stubGlobal("fetch", async () => {
      count += 1;
      await new Promise((resolve) => setTimeout(resolve, 25));
      return new Response(
        JSON.stringify({ access_token: accessToken, refresh_token: "rt-2", expires_in: 3600 }),
        { status: 200 },
      );
    });
    try {
      const [first, second] = await Promise.all([getTeamsWebToken(), getTeamsWebToken()]);
      expect(count).toBe(1);
      expect(first.token).toBe(accessToken);
      expect(second.token).toBe(accessToken);
    } finally {
      vi.unstubAllGlobals();
    }
    expect((await readTeamsWebCache()).cache?.refresh_token).toBe("rt-2");
  });
});

describe("logout", () => {
  it("clears the Teams-web cache and reports it as missing afterwards", async () => {
    await writeTeamsWebCache(validCache());
    expect((await teamsWebStatus()).authenticated).toBe(true);

    const result = await authTool.handle({ op: "logout" });
    expect(result.isError).toBeUndefined();
    const body = JSON.parse((result.content[0] as { text: string }).text) as Record<string, unknown>;
    expect(body.teams_web_logged_out).toBe(true);
    expect((await readTeamsWebCache()).cache).toBeNull();

    const status = await teamsWebStatus();
    expect(status.authenticated).toBe(false);
    expect(status.reason).toBe("missing");
  });
});

describe("getTeamsWebToken", () => {
  it("returns a fresh cache without refresh or browser", async () => {
    await writeTeamsWebCache(validCache({ source: "browser" }));
    let touched = false;
    const token = await getTeamsWebToken({
      deps: {
        refresh: async () => {
          touched = true;
          return validCache();
        },
        harvest: async () => {
          touched = true;
          return validCache();
        },
      },
    });
    expect(token.token).toBe("access-token");
    expect(token.source).toBe("browser");
    expect(touched).toBe(false);
  });

  it("force refreshes even while the cached token is still fresh", async () => {
    await writeTeamsWebCache(validCache({ refresh_token: "old-rt" }));
    let called = false;
    const refreshed = validCache({ access_token: "forced", refresh_token: "new-rt" });
    const token = await getTeamsWebToken({
      force: true,
      deps: {
        refresh: async ({ force }) => {
          expect(force).toBe(true);
          called = true;
          await writeTeamsWebCache(refreshed);
          return refreshed;
        },
      },
    });
    expect(called).toBe(true);
    expect(token.token).toBe("forced");
  });

  it("refreshes an expired token, persisting the rotated refresh token", async () => {
    await writeTeamsWebCache(
      validCache({ expires_at: new Date(Date.now() - 60_000).toISOString(), refresh_token: "old-rt" }),
    );
    const refreshed = validCache({
      access_token: "new-at",
      refresh_token: "new-rt",
      expires_at: new Date(Date.now() + 3_600_000).toISOString(),
    });
    const token = await getTeamsWebToken({
      deps: {
        refresh: async ({ cache }) => {
          expect(cache.refresh_token).toBe("old-rt");
          await writeTeamsWebCache(refreshed);
          return refreshed;
        },
        harvest: async () => {
          throw new Error("browser must not run");
        },
      },
    });
    expect(token.token).toBe("new-at");
    expect((await readTeamsWebCache()).cache?.refresh_token).toBe("new-rt");
  });

  it("falls back to a harvest when there is no refresh token", async () => {
    await writeTeamsWebCache(
      validCache({ expires_at: new Date(Date.now() - 60_000).toISOString(), refresh_token: undefined }),
    );
    const harvested = validCache({ access_token: "harvested", source: "browser" });
    const token = await getTeamsWebToken({
      deps: {
        harvest: async () => harvested,
      },
    });
    expect(token.token).toBe("harvested");
  });

  it("refuses to launch a browser when allowBrowser is false", async () => {
    await expect(getTeamsWebToken({ allowBrowser: false })).rejects.toBeInstanceOf(AuthRequiredError);
  });
});
