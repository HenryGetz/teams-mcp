import { describe, expect, it } from "vitest";
import { classifySignInPage, extractAadsts, isOwaUrl, type PageSnapshot } from "../src/auth/state.js";

function page(overrides: Partial<PageSnapshot> = {}): PageSnapshot {
  return {
    url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    body: "",
    buttonValue: "",
    hasConsentButton: false,
    hasOtpInput: false,
    hasMfaSwitcher: false,
    shellReady: false,
    ...overrides,
  };
}

describe("isOwaUrl", () => {
  it("accepts both OWA hosts", () => {
    expect(isOwaUrl("https://outlook.office.com/mail/")).toBe(true);
    expect(isOwaUrl("https://outlook.cloud.microsoft/mail/")).toBe(true);
  });

  it("rejects the login host even when it carries an outlook redirect", () => {
    expect(isOwaUrl("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?redirect_uri=https://outlook.cloud.microsoft/mail/")).toBe(false);
  });
});

describe("extractAadsts", () => {
  it("pulls the code out of a page body", () => {
    expect(extractAadsts("Sorry, but we're having trouble signing you in. AADSTS50105: blocked")).toBe("AADSTS50105");
    expect(extractAadsts("nothing here")).toBeUndefined();
  });
});

describe("classifySignInPage", () => {
  it("reports the AOW shell only when it has actually rendered", () => {
    expect(classifySignInPage(page({ url: "https://outlook.office.com/mail/", shellReady: true })).step).toBe("owa");
    // URL flips before the app renders: must not be mistaken for success.
    expect(classifySignInPage(page({ url: "https://outlook.office.com/mail/", shellReady: false })).step).toBe("unknown");
  });

  it("aborts on an AADSTS error before anything else", () => {
    const result = classifySignInPage(page({ body: "AADSTS50105: your administrator", hasConsentButton: true }));
    expect(result).toEqual({ step: "aadsts", aadsts: "AADSTS50105" });
  });

  it("distinguishes the identifier step from the password step", () => {
    expect(classifySignInPage(page({ buttonValue: "Next" })).step).toBe("identifier");
    expect(classifySignInPage(page({ buttonValue: "Sign in", body: "Enter password" })).step).toBe("password");
  });

  it("regression: the identifier field being present does not make it a password step", () => {
    // Both #i0116 and #i0118 live in the DOM on every step; only the button label
    // changes. A password must never be typed on the identifier step.
    const identifier = classifySignInPage(page({ buttonValue: "Next", body: "Sign in No account? Create one!" }));
    expect(identifier.step).toBe("identifier");
  });

  it("detects the account picker used by a warm profile", () => {
    const result = classifySignInPage(page({ body: "Pick an account Student, Test S student@example.edu Signed in Use another account" }));
    expect(result.step).toBe("account-picker");
  });

  it("detects the TOTP input and the MFA switcher", () => {
    expect(classifySignInPage(page({ hasOtpInput: true, body: "Enter code" })).step).toBe("totp");
    expect(classifySignInPage(page({ hasMfaSwitcher: true })).step).toBe("mfa-switch");
  });

  it("detects consent", () => {
    expect(classifySignInPage(page({ hasConsentButton: true })).step).toBe("consent");
  });

  it("detects 'Stay signed in?' by body or button label", () => {
    expect(classifySignInPage(page({ body: "Stay signed in? Yes No" })).step).toBe("kmsi");
    expect(classifySignInPage(page({ buttonValue: "Yes" })).step).toBe("kmsi");
  });

  it("flags an MFA setup wall as unrecoverable", () => {
    expect(classifySignInPage(page({ body: "More information required" })).step).toBe("mfa-setup");
    expect(classifySignInPage(page({ body: "Let's keep your account secure" })).step).toBe("mfa-setup");
  });

  it("returns unknown for an unrecognised page", () => {
    expect(classifySignInPage(page({ body: "Loading…" })).step).toBe("unknown");
  });
});
