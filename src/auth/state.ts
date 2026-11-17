/**
 * Sign-in page classifier.
 *
 * The Microsoft login front end is a single SPA shell: the identifier field
 * `#i0116` and the password field `#i0118` are BOTH in the DOM on every step, so
 * field presence proves nothing. The reliable discriminators, observed on the
 * tested tenant, are the submit button's label plus the heading text:
 *
 *   identifier step -> button value "Next"
 *   password step   -> button value "Sign in" (heading "Enter password")
 *
 * Kept pure so every branch is unit-testable without a browser.
 */

export type SignInStep =
  | "owa"
  | "aadsts"
  | "consent"
  | "totp"
  | "mfa-switch"
  | "account-picker"
  | "kmsi"
  | "password"
  | "identifier"
  | "mfa-setup"
  | "unknown";

export interface PageSnapshot {
  url: string;
  /** `document.body.innerText`, whitespace-collapsed by the caller. */
  body: string;
  /** Value of `#idSIButton9`, or "" when absent. */
  buttonValue: string;
  hasConsentButton: boolean;
  hasOtpInput: boolean;
  hasMfaSwitcher: boolean;
  /** An OWA mailbox shell selector is present. */
  shellReady: boolean;
}

export interface Classification {
  step: SignInStep;
  /** The AADSTS code, when the page is a hard failure. */
  aadsts?: string;
}

const AADSTS = /AADSTS\d{5,7}/;
// Observed OWA hosts: outlook.office.com (redirects) and outlook.cloud.microsoft.
const OWA_HOST = /outlook\.(office\.com|cloud\.microsoft)/i;
const SIGN_IN_HOST = /login\.(microsoftonline|live)\.com|login\.microsoft\.com/i;

/** Is this URL an OWA app URL (as opposed to a login host)? */
export function isOwaUrl(url: string): boolean {
  return OWA_HOST.test(url) && !SIGN_IN_HOST.test(url);
}

/** First AADSTS code in a page's text, if any. */
export function extractAadsts(body: string): string | undefined {
  return AADSTS.exec(body)?.[0];
}

export function classifySignInPage(page: PageSnapshot): Classification {
  const aadsts = extractAadsts(page.body);
  if (aadsts) return { step: "aadsts", aadsts };

  if (page.shellReady && isOwaUrl(page.url)) return { step: "owa" };
  if (page.hasConsentButton) return { step: "consent" };
  if (page.hasOtpInput) return { step: "totp" };
  if (page.hasMfaSwitcher) return { step: "mfa-switch" };
  if (/Pick an account/i.test(page.body)) return { step: "account-picker" };
  if (/Stay signed in\?/i.test(page.body) || /^Yes$/i.test(page.buttonValue)) return { step: "kmsi" };

  // Password before identifier: on the password step the identifier field is
  // still present but empty in the DOM, so an "is the field there?" test would
  // misfire and type a password into the wrong step.
  const passwordStep = /Enter password/i.test(page.body) || /^Sign in$/i.test(page.buttonValue);
  const identifierStep = /^Next$/i.test(page.buttonValue);
  if (passwordStep && !identifierStep) return { step: "password" };
  if (identifierStep) return { step: "identifier" };

  if (/More information required|Let.s keep your account secure/i.test(page.body)) return { step: "mfa-setup" };
  return { step: "unknown" };
}
