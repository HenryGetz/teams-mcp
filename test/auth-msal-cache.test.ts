import { describe, expect, it } from "vitest";
import {
  cachedAccessTokenAudiences,
  idTokenClaims,
  parseMsalKey,
  parseMsalStorage,
  pickRefreshToken,
} from "../src/auth/msal-cache.js";

const OWA_CLIENT = "9199bf20-a13f-4107-85dc-02114787ef48";
const HOME_ACCOUNT = "22222222-2222-4222-8222-222222222222.11111111-1111-4111-8111-111111111111";
const TENANT = "11111111-1111-4111-8111-111111111111";

function jwt(claims: Record<string, unknown>): string {
  return `header.${Buffer.from(JSON.stringify(claims), "utf8").toString("base64url")}.signature`;
}

/**
 * A synthetic fixture in the exact observed MSAL localStorage shape:
 * refresh token, id token, two access tokens, plus MSAL bookkeeping keys and one
 * unparsable value.
 */
function fixture(): Record<string, string> {
  return {
    [`msal.3|${HOME_ACCOUNT}|login.windows.net|refreshtoken|${OWA_CLIENT}|||`]: JSON.stringify({
      homeAccountId: HOME_ACCOUNT,
      environment: "login.windows.net",
      clientId: OWA_CLIENT,
      credentialType: "RefreshToken",
      secret: "1.AUYA…REDACTED…m97h7g",
      target: "",
    }),
    [`msal.3|${HOME_ACCOUNT}|login.windows.net|idtoken|${OWA_CLIENT}|${TENANT}||`]: JSON.stringify({
      credentialType: "IdToken",
      secret: jwt({ preferred_username: "student@example.edu", oid: "22222222-2222-4222-8222-222222222222", tid: TENANT }),
    }),
    [`msal.3|${HOME_ACCOUNT}|login.windows.net|accesstoken|${OWA_CLIENT}|${TENANT}|https://outlook.office.com/.default`]: JSON.stringify({
      credentialType: "AccessToken",
      secret: jwt({ aud: "https://outlook.office.com", scp: "Mail.ReadWrite" }),
      expiresOn: 1794867039,
    }),
    [`msal.3|${HOME_ACCOUNT}|login.windows.net|accesstoken|${OWA_CLIENT}|${TENANT}|https://graph.microsoft.com/.default`]: JSON.stringify({
      credentialType: "AccessToken",
      secret: jwt({ aud: "https://graph.microsoft.com", scp: "Directory.Read.All" }),
    }),
    [`msal.3.token.keys.${OWA_CLIENT}`]: JSON.stringify(["a", "b"]),
    "msal.3.account.keys": JSON.stringify([HOME_ACCOUNT]),
    "olk-isauthed": "1",
    "olk-sdmp": JSON.stringify({ TimeZoneStr: "Central Standard Time" }),
    [`msal.3|${HOME_ACCOUNT}|login.windows.net|badentry|${OWA_CLIENT}`]: "{not json",
  };
}

describe("parseMsalKey", () => {
  it("splits a credential key into its fields", () => {
    const parsed = parseMsalKey(`msal.3|${HOME_ACCOUNT}|login.windows.net|refreshtoken|${OWA_CLIENT}|||`);
    expect(parsed.homeAccountId).toBe(HOME_ACCOUNT);
    expect(parsed.environment).toBe("login.windows.net");
    expect(parsed.credentialType).toBe("refreshtoken");
    expect(parsed.clientId).toBe(OWA_CLIENT);
  });

  it("returns an empty object for non-MSAL keys", () => {
    expect(parseMsalKey("olk-isauthed")).toEqual({});
    expect(parseMsalKey("msal.3|short")).toEqual({});
  });
});

describe("parseMsalStorage", () => {
  it("extracts only credential entries and skips bookkeeping", () => {
    const entries = parseMsalStorage(fixture());
    expect(entries.map((entry) => entry.credentialType).sort()).toEqual(["AccessToken", "AccessToken", "IdToken", "RefreshToken"]);
    expect(entries.every((entry) => entry.key.startsWith("msal."))).toBe(true);
  });

  it("skips unparsable values instead of throwing", () => {
    expect(() => parseMsalStorage(fixture())).not.toThrow();
  });

  it("prefers an entry whose clientId matches when asked", () => {
    const entries = parseMsalStorage(fixture());
    const refresh = pickRefreshToken(entries, OWA_CLIENT);
    expect(refresh?.clientId).toBe(OWA_CLIENT);
    expect(refresh?.secret).toBe("1.AUYA…REDACTED…m97h7g");
  });

  it("refuses a refresh token bound to a different client", () => {
    // A refresh token is only redeemable by the client that received it, so a
    // mismatch must not be handed to the refresh call.
    expect(pickRefreshToken(parseMsalStorage(fixture()), "some-other-client")).toBeUndefined();
  });

  it("accepts an entry whose client id is not recorded", () => {
    const storage = {
      "msal.3|acct|login.windows.net|refreshtoken": JSON.stringify({ credentialType: "RefreshToken", secret: "opaque-rt" }),
    };
    expect(pickRefreshToken(parseMsalStorage(storage), OWA_CLIENT)?.secret).toBe("opaque-rt");
  });

  it("returns undefined when no refresh token is present", () => {
    const storage = fixture();
    for (const key of Object.keys(storage)) {
      if (key.includes("refreshtoken")) delete storage[key];
    }
    expect(pickRefreshToken(parseMsalStorage(storage), OWA_CLIENT)).toBeUndefined();
  });

  it("ignores a refresh token entry with an empty secret", () => {
    const storage = fixture();
    const key = Object.keys(storage).find((candidate) => candidate.includes("refreshtoken")) as string;
    storage[key] = JSON.stringify({ credentialType: "RefreshToken", secret: "" });
    expect(pickRefreshToken(parseMsalStorage(storage), OWA_CLIENT)).toBeUndefined();
  });
});

describe("identity claims", () => {
  it("reads upn and oid from the id token", () => {
    const claims = idTokenClaims(parseMsalStorage(fixture()));
    expect(claims?.preferred_username).toBe("student@example.edu");
    expect(claims?.oid).toBe("22222222-2222-4222-8222-222222222222");
    expect(claims?.tid).toBe(TENANT);
  });

  it("lists which audiences OWA already cached", () => {
    expect(cachedAccessTokenAudiences(parseMsalStorage(fixture())).sort()).toEqual([
      "https://graph.microsoft.com",
      "https://outlook.office.com",
    ]);
  });

  it("tolerates a missing id token", () => {
    const storage = fixture();
    for (const key of Object.keys(storage)) {
      if (key.includes("idtoken")) delete storage[key];
    }
    expect(idTokenClaims(parseMsalStorage(storage))).toBeUndefined();
  });
});
