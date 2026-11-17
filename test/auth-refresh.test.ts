import { describe, expect, it, vi } from "vitest";
import {
  AUDIENCE_SCOPE,
  OWA_CLIENT_ID,
  OWA_ORIGIN,
  RefreshError,
  buildRefreshRequest,
  decodeJwtClaims,
  parseTokenError,
  parseTokenSuccess,
  refreshAccessToken,
} from "../src/auth/refresh.js";

const TENANT = "11111111-1111-4111-8111-111111111111";
const RT = "1.AUYA-refresh-token";

function jwt(claims: Record<string, unknown>): string {
  const payload = Buffer.from(JSON.stringify(claims), "utf8").toString("base64url");
  return `header.${payload}.signature`;
}

function formOf(body: string): URLSearchParams {
  return new URLSearchParams(body);
}

describe("buildRefreshRequest", () => {
  it("posts the refresh grant with the OWA client and no secret", () => {
    const request = buildRefreshRequest({ tenantId: TENANT, refreshToken: RT, audience: "outlook" });
    const form = formOf(request.body);
    expect(request.url).toBe(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`);
    expect(request.method).toBe("POST");
    expect(form.get("grant_type")).toBe("refresh_token");
    expect(form.get("client_id")).toBe(OWA_CLIENT_ID);
    expect(form.get("refresh_token")).toBe(RT);
    expect(form.get("scope")).toBe("https://outlook.office.com/.default");
    // A public client: no secret may ever be sent.
    expect(form.get("client_secret")).toBeNull();
  });

  it("sends the Origin header the SPA client requires", () => {
    const request = buildRefreshRequest({ tenantId: TENANT, refreshToken: RT, audience: "graph" });
    expect(request.headers.Origin).toBe(OWA_ORIGIN);
    expect(request.headers["Content-Type"]).toBe("application/x-www-form-urlencoded");
    expect(formOf(request.body).get("scope")).toBe(AUDIENCE_SCOPE.graph);
  });

  it("uses exactly one resource .default scope per audience", () => {
    const outlook = formOf(buildRefreshRequest({ tenantId: TENANT, refreshToken: RT, audience: "outlook" }).body).get("scope");
    const graph = formOf(buildRefreshRequest({ tenantId: TENANT, refreshToken: RT, audience: "graph" }).body).get("scope");
    expect(outlook?.split(" ")).toHaveLength(1);
    expect(graph?.split(" ")).toHaveLength(1);
    expect(outlook).toMatch(/^https:\/\/outlook\.office\.com\/\.default$/);
    expect(graph).toMatch(/^https:\/\/graph\.microsoft\.com\/\.default$/);
  });

  it("can omit scope for diagnostics", () => {
    const request = buildRefreshRequest({ tenantId: TENANT, refreshToken: RT, audience: "outlook", scope: null });
    expect(formOf(request.body).get("scope")).toBeNull();
  });
});

describe("parseTokenError", () => {
  it("classifies a dead refresh token as requiring login", () => {
    const error = parseTokenError(400, {
      error: "invalid_grant",
      error_description: "AADSTS70008: The provided authorization grant has expired.",
      trace_id: "trace-1",
    });
    expect(error).toBeInstanceOf(RefreshError);
    expect(error.code).toBe("invalid_grant");
    expect(error.aadsts).toBe("AADSTS70008");
    expect(error.requiresLogin).toBe(true);
    expect(error.message).not.toContain("Trace ID");
    expect(error.hint).toMatch(/run auth\.login/);
  });

  it("recognises the SPA cross-origin rejection as configuration, not a dead session", () => {
    const error = parseTokenError(400, {
      error: "invalid_request",
      error_description:
        "AADSTS9002327: Tokens issued for the 'Single-Page Application' client-type may only be redeemed via cross-origin requests.",
    });
    expect(error.aadsts).toBe("AADSTS9002327");
    expect(error.requiresLogin).toBe(false);
    expect(error.hint).toMatch(/Origin header/);
  });

  it("recognises first-party preauthorization failures", () => {
    const error = parseTokenError(400, {
      error: "invalid_request",
      error_description: "AADSTS65002: Consent between first party application and first party resource must be configured via preauthorization",
    });
    expect(error.aadsts).toBe("AADSTS65002");
    expect(error.requiresLogin).toBe(false);
    expect(error.hint).toMatch(/\.default scope/);
  });

  it("treats an unknown AADSTS code as needing a login but keeps the code", () => {
    const error = parseTokenError(401, { error_description: "AADSTS99999: something new" });
    expect(error.aadsts).toBe("AADSTS99999");
    expect(error.requiresLogin).toBe(false);
    expect(error.status).toBe(401);
  });
});

describe("parseTokenSuccess", () => {
  it("derives expiry, scopes and claims from the access token", () => {
    const accessToken = jwt({ aud: "https://outlook.office.com", tid: TENANT, scp: "Mail.ReadWrite Calendars.ReadWrite", exp: 1 });
    const result = parseTokenSuccess({ access_token: accessToken, refresh_token: "rt-2", expires_in: 3600 }, 1_000);
    expect(result.accessToken).toBe(accessToken);
    expect(result.refreshToken).toBe("rt-2");
    expect(result.expiresIn).toBe(3600);
    expect(result.expiresAt).toBe(new Date(1_000 + 3_600_000).toISOString());
    expect(result.scope).toBe("Mail.ReadWrite Calendars.ReadWrite");
    expect(result.claims?.tid).toBe(TENANT);
  });

  it("falls back to a one hour lifetime when expires_in is unusable", () => {
    const result = parseTokenSuccess({ access_token: jwt({ aud: "x" }), expires_in: "nonsense" }, 0);
    expect(result.expiresIn).toBe(3600);
    expect(result.refreshToken).toBeUndefined();
  });

  it("rejects a response without an access token", () => {
    expect(() => parseTokenSuccess({}, 0)).toThrow(/no access_token/);
  });
});

describe("refreshAccessToken", () => {
  it("returns the rotated refresh token and never invents one", async () => {
    const accessToken = jwt({ aud: "https://outlook.office.com", scp: "Mail.ReadWrite" });
    const fetchImpl = vi.fn(async (_input: string, _init?: RequestInit) =>
      new Response(JSON.stringify({ access_token: accessToken, refresh_token: "rt-rotated", expires_in: 4546 }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      }),
    );
    const result = await refreshAccessToken({ tenantId: TENANT, refreshToken: RT, audience: "outlook", fetchImpl, now: 0 });
    expect(result.refreshToken).toBe("rt-rotated");
    expect(result.expiresIn).toBe(4546);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [url, init] = fetchImpl.mock.calls[0] ?? ["", undefined];
    expect(url).toContain(TENANT);
    expect((init?.headers as Record<string, string>).Origin).toBe(OWA_ORIGIN);
  });

  it("omits refreshToken when the provider did not rotate it", async () => {
    const fetchImpl = vi.fn(async () => new Response(JSON.stringify({ access_token: jwt({ aud: "g" }), expires_in: 60 }), { status: 200 }));
    const result = await refreshAccessToken({ tenantId: TENANT, refreshToken: RT, audience: "graph", fetchImpl, now: 0 });
    expect(result.refreshToken).toBeUndefined();
  });

  it("raises a typed error for a rejected grant", async () => {
    const fetchImpl = vi.fn(async () =>
      new Response(JSON.stringify({ error: "invalid_grant", error_description: "AADSTS70008: expired" }), { status: 400 }),
    );
    await expect(refreshAccessToken({ tenantId: TENANT, refreshToken: RT, audience: "outlook", fetchImpl })).rejects.toMatchObject({
      name: "RefreshError",
      code: "invalid_grant",
      aadsts: "AADSTS70008",
    });
  });

  it("maps a transport failure to a non-fatal network error", async () => {
    const fetchImpl = vi.fn(async () => {
      throw new Error("ECONNRESET");
    });
    const error = await refreshAccessToken({ tenantId: TENANT, refreshToken: RT, audience: "outlook", fetchImpl }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(RefreshError);
    expect((error as RefreshError).code).toBe("network_error");
    expect((error as RefreshError).requiresLogin).toBe(false);
    expect((error as RefreshError).hint).toMatch(/does not require a new login/);
  });

  it("handles a non-JSON error page", async () => {
    const fetchImpl = vi.fn(async () => new Response("<html>gateway timeout</html>", { status: 504 }));
    await expect(refreshAccessToken({ tenantId: TENANT, refreshToken: RT, audience: "outlook", fetchImpl })).rejects.toMatchObject({
      status: 504,
    });
  });
});

describe("decodeJwtClaims", () => {
  it("returns undefined for malformed input", () => {
    expect(decodeJwtClaims("not-a-jwt")).toBeUndefined();
    expect(decodeJwtClaims("a.@@@.c")).toBeUndefined();
  });
});
