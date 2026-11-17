/**
 * Sign-in flow: persistent browser profile -> MSAL harvest -> both tokens.
 *
 * The profile at `config().profileDir` (~/.config/teams-mcp/browser-profile) is
 * the anti-stale-login path: once it holds an SSO session, re-authentication is
 * a single "Pick an account" click — no password, no MFA. The password and TOTP
 * secret are only used when the profile is cold.
 *
 * Everything the flow learns comes from the live tenant; the observed selectors
 * and timings are documented in docs/auth.md.
 */
import { readFile, stat } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { chromium, type BrowserContext, type Page } from "playwright";
import { config } from "../config.js";
import {
  cachedAccessTokenAudiences,
  idTokenClaims,
  parseMsalStorage,
  pickRefreshToken,
  type MsalEntry,
} from "./msal-cache.js";
import { refreshAccessToken, decodeJwtClaims, OWA_CLIENT_ID, type RefreshClaims, type RefreshResult } from "./refresh.js";
import { classifySignInPage, type PageSnapshot, type SignInStep } from "./state.js";
import {
  persistTokenUpdate,
  readTokenCache,
  type Audience,
  type TokenCache,
} from "./store.js";
import { totpForSubmission } from "./totp.js";
import { ensureDisplay, XvfbUnavailableError, type XvfbHandle } from "./xvfb.js";

export const OWA_URL = "https://outlook.office.com/mail/";
export const CREDENTIALS_FILE = "credentials.json";
/** Selectors that only exist once the OWA mailbox shell has really rendered. */
export const OWA_SHELL_SELECTORS = [
  '[aria-label*="Message list" i]',
  "#MailList",
  '[data-automationid*="mail" i]',
  '[aria-label*="Folder pane" i]',
];

export interface Credentials {
  username: string;
  password: string;
  totpSecret?: string;
  /** Where they came from, for status output. Never includes the values. */
  origin: "env" | "file";
}

export class LoginError extends Error {
  readonly code: string;
  readonly hint?: string;
  constructor(code: string, message: string, hint?: string) {
    super(message);
    this.name = "LoginError";
    this.code = code;
    this.hint = hint;
  }
}

export type LogFn = (message: string) => void;

/** All diagnostics go to stderr: stdout is reserved for MCP frames. */
const stderrLog: LogFn = (message) => process.stderr.write(`${message}\n`);

/** Redact anything token-shaped before it can reach a log line. */
export function redactSecret(value: string | undefined): string {
  if (!value) return "(none)";
  if (value.length <= 20) return "REDACTED";
  return `${value.slice(0, 6)}…${value.slice(-6)} (len ${value.length})`;
}

/**
 * Credentials come from the environment first (useful for CI/one-off runs) and
 * otherwise from `credentials.json` in the config dir, which must be 0600.
 */
export async function readCredentials(): Promise<Credentials> {
  const envUser = process.env.TEAMS_MCP_USERNAME?.trim();
  const envPassword = process.env.TEAMS_MCP_PASSWORD;
  const envTotp = process.env.TEAMS_MCP_TOTP_SECRET?.trim();
  if (envUser && envPassword) {
    return {
      username: envUser,
      password: envPassword,
      ...(envTotp ? { totpSecret: envTotp } : {}),
      origin: "env",
    };
  }

  const path = join(config().configDir, CREDENTIALS_FILE);
  let raw: string;
  try {
    raw = await readFile(path, "utf8");
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    throw new LoginError(
      "credentials_missing",
      code === "ENOENT" ? `no credentials file at ${path} and TEAMS_MCP_USERNAME/TEAMS_MCP_PASSWORD are unset` : `credentials file unreadable at ${path} (${code ?? "error"})`,
      `create ${path} with {"username","password","totpSecret"} and chmod 600, or export TEAMS_MCP_USERNAME/TEAMS_MCP_PASSWORD/TEAMS_MCP_TOTP_SECRET`,
    );
  }
  const mode = (await stat(path)).mode & 0o777;
  if (mode !== 0o600) {
    stderrLog(`[auth] warning: ${path} is mode ${mode.toString(8)}, expected 600`);
  }
  let parsed: Record<string, unknown>;
  try {
    parsed = JSON.parse(raw) as Record<string, unknown>;
  } catch {
    throw new LoginError("credentials_invalid", `credentials file at ${path} is not valid JSON`);
  }
  const username = typeof parsed.username === "string" ? parsed.username.trim() : "";
  const password = typeof parsed.password === "string" ? parsed.password : "";
  const totpSecret = typeof parsed.totpSecret === "string" ? parsed.totpSecret.trim() : undefined;
  if (!username || !password) {
    throw new LoginError(
      "credentials_incomplete",
      `credentials file at ${path} must contain non-empty "username" and "password"`,
      'add "totpSecret" (base32) as well, otherwise MFA cannot be answered automatically',
    );
  }
  return { username, password, ...(totpSecret ? { totpSecret } : {}), origin: "file" };
}

export interface LaunchOptions {
  profileDir: string;
}

export interface SignInSession {
  context: BrowserContext;
  page: Page;
  xvfb: XvfbHandle | null;
  close(): Promise<void>;
}

/**
 * Launch headed Chromium on the persistent profile. Headless is avoided on
 * purpose: the Microsoft login front end fingerprints it.
 */
export async function openBrowserSession(options: LaunchOptions, log: LogFn): Promise<SignInSession> {
  let xvfb: XvfbHandle | null = null;
  try {
    const display = await ensureDisplay();
    xvfb = display.xvfb;
    if (xvfb) log(`[auth] started Xvfb on ${display.display}`);
  } catch (error) {
    if (error instanceof XvfbUnavailableError) throw new LoginError(error.code, error.message, error.hint);
    throw error;
  }

  const context = await chromium.launchPersistentContext(options.profileDir, {
    headless: false,
    viewport: { width: 1280, height: 900 },
    locale: "en-US",
    timezoneId: config().timezone,
    args: ["--disable-blink-features=AutomationControlled", "--no-first-run", "--no-default-browser-check"],
  });
  // The login page checks this property; a headed Chromium behaves normally otherwise.
  await context.addInitScript(() => {
    Object.defineProperty(navigator, "webdriver", { get: () => undefined });
  });
  const page = context.pages()[0] ?? (await context.newPage());
  return {
    context,
    page,
    xvfb,
    async close() {
      await context.close().catch(() => undefined);
      xvfb?.stop();
      if (xvfb) log(`[auth] stopped Xvfb on :${xvfb.display}`);
    },
  };
}

async function snapshot(page: Page): Promise<PageSnapshot> {
  const url = page.url();
  const body = await page.evaluate(() => document.body?.innerText ?? "").catch(() => "");
  const buttonValue = await page.locator("#idSIButton9").inputValue({ timeout: 400 }).catch(() => "");
  const [hasConsentButton, hasOtpInput, hasMfaSwitcher, shellReady] = await Promise.all([
    page.locator("#idBtn_Accept").count().then((n) => n > 0).catch(() => false),
    page.locator('#idTxtBcOtc, input[name="otc"], input[autocomplete="one-time-code"]').count().then((n) => n > 0).catch(() => false),
    page
      .getByText(/Use a verification code|I can.t use my Microsoft Authenticator app right now|Use a different verification option|Sign in another way/i)
      .count()
      .then((n) => n > 0)
      .catch(() => false),
    page
      .evaluate(
        (selectors) => selectors.some((selector) => Boolean(document.querySelector(selector))),
        OWA_SHELL_SELECTORS,
      )
      .catch(() => false),
  ]);
  return { url, body: body.replace(/\s+/g, " "), buttonValue, hasConsentButton, hasOtpInput, hasMfaSwitcher, shellReady };
}

async function screenshotOnFailure(page: Page, label: string, log: LogFn): Promise<string | undefined> {
  const dir = config().cacheDir;
  const file = join(dir, `auth-${label}-${Date.now()}.png`);
  try {
    await page.screenshot({ path: file, fullPage: false });
    log(`[auth] screenshot: ${file}`);
    return file;
  } catch {
    return undefined;
  }
}

const STEP_ACTIONS: Partial<Record<SignInStep, (page: Page, credentials: Credentials, log: LogFn) => Promise<void>>> = {
  async consent(page, _credentials, log) {
    log("[auth] accepting consent");
    await page.locator("#idBtn_Accept").click();
  },
  async totp(page, credentials, log) {
    const secret = credentials.totpSecret;
    if (!secret) {
      throw new LoginError(
        "totp_secret_missing",
        "the tenant asked for a verification code but no totpSecret is configured",
        'add "totpSecret" (base32) to credentials.json or export TEAMS_MCP_TOTP_SECRET',
      );
    }
    // Do not submit a code that is about to roll over; it would be rejected.
    let submitted;
    try {
      submitted = totpForSubmission(secret, { minimumSecondsLeft: 5 });
    } catch {
      log("[auth] TOTP window closing, waiting for the next one");
      await delay(6000);
      submitted = totpForSubmission(secret, { minimumSecondsLeft: 5 });
    }
    log(`[auth] submitting TOTP (${submitted.secondsLeft}s left in window)`);
    await page.locator('#idTxtBcOtc, input[name="otc"], input[autocomplete="one-time-code"]').first().fill(submitted.code);
    await page
      .locator("#idSubmit_SAOTCC_Continue, #idSIButton9")
      .first()
      .click()
      .catch(() => page.keyboard.press("Enter"));
  },
  async "mfa-switch"(page, _credentials, log) {
    log("[auth] switching MFA to verification code");
    await page
      .getByText(/Use a verification code|I can.t use my Microsoft Authenticator app right now|Use a different verification option|Sign in another way/i)
      .first()
      .click()
      .catch(() => undefined);
  },
  async "account-picker"(page, credentials, log) {
    // Warm profile: clicking the tile is enough. Never type a password here.
    log(`[auth] picking cached account ${credentials.username}`);
    await page
      .locator(`div[role="button"]:has-text("${credentials.username}")`)
      .first()
      .click({ timeout: 8000 })
      .catch(async () => {
        await page.getByText(credentials.username, { exact: false }).first().click().catch(() => undefined);
      });
  },
  async kmsi(page, _credentials, log) {
    log("[auth] answering 'Stay signed in?' with Yes");
    await page
      .locator("#idSIButton9")
      .click()
      .catch(() => page.getByRole("button", { name: /^Yes$/ }).first().click());
  },
  async password(page, credentials, log) {
    log("[auth] submitting password");
    await page.locator("#i0118").fill(credentials.password, { timeout: 15_000 });
    await page.locator("#idSIButton9").click().catch(() => page.keyboard.press("Enter"));
  },
  async identifier(page, credentials, log) {
    const current = await page.locator("#i0116").inputValue({ timeout: 2000 }).catch(() => "");
    if (current) return;
    log(`[auth] submitting identifier ${credentials.username}`);
    await page.locator("#i0116").fill(credentials.username, { timeout: 15_000 });
    await page.locator("#idSIButton9").click().catch(() => page.keyboard.press("Enter"));
  },
};

export interface PerformLoginOptions {
  log?: LogFn;
  /** Overall budget for reaching the OWA shell. */
  timeoutMs?: number;
  /** Poll interval while waiting for the page to settle. */
  stepIntervalMs?: number;
  maxSteps?: number;
  /** Settle time after the shell renders, so MSAL finishes writing the cache. */
  harvestDelayMs?: number;
  now?: () => number;
}

export interface LoginTimings {
  totalMs: number;
  shellMs: number;
  harvestMs: number;
  outlookRefreshMs: number;
  graphRefreshMs: number;
  steps: number;
}

export interface LoginOutcome {
  cache: TokenCache;
  timings: LoginTimings;
  warm: boolean;
  credentialsOrigin: Credentials["origin"];
  cachedAudiences: string[];
  profileDir: string;
  path: string;
}

export interface DriveSignInOptions {
  /**
   * Return true once the target application is usable. Receives the page
   * snapshot so callers do not pay for a second DOM read.
   */
  ready: (state: PageSnapshot) => boolean | Promise<boolean>;
  /** Overall budget for reaching the ready state. */
  timeoutMs?: number;
  stepIntervalMs?: number;
  maxSteps?: number;
  log?: LogFn;
}

/**
 * Drive whatever Microsoft sign-in page is on screen until `ready()` says the
 * application has loaded. Shared by the OWA flow (mail shell selector) and the
 * Teams web harvest (app shell + MSAL cache).
 */
export async function driveSignIn(
  page: Page,
  credentials: Credentials,
  options: DriveSignInOptions,
): Promise<{ steps: number; warm: boolean }> {
  const log = options.log ?? stderrLog;
  const stepIntervalMs = options.stepIntervalMs ?? 1500;
  const maxSteps = options.maxSteps ?? 20;
  const deadline = Date.now() + (options.timeoutMs ?? 180_000);
  let steps = 0;
  let sawPassword = false;

  while (steps < maxSteps) {
    if (Date.now() > deadline) {
      throw new LoginError(
        "login_timeout",
        `sign-in did not reach the application within the time budget (${steps} steps)`,
        "check connectivity and the screenshot path in the log; the profile may need an interactive login",
      );
    }
    await delay(stepIntervalMs);
    const state = await snapshot(page);
    steps += 1;
    if (await options.ready(state)) return { steps, warm: !sawPassword };

    const { step, aadsts } = classifySignInPage(state);
    if (step === "aadsts") {
      const file = await screenshotOnFailure(page, "aadsts", log);
      throw new LoginError(
        "signin_rejected",
        `${aadsts ?? "AADSTS error"}: ${state.body.slice(0, 200)}`,
        [
          "this tenant blocks non-assigned apps (AADSTS50105) and requires the Outlook Web client",
          file ? `screenshot: ${file}` : "no screenshot available",
        ].join("; "),
      );
    }
    if (step === "mfa-setup") {
      await screenshotOnFailure(page, "mfa-setup", log);
      throw new LoginError(
        "mfa_setup_required",
        "the account has an unfinished MFA registration; complete it in a browser first",
        "sign in once interactively with the same profile dir, then retry",
      );
    }

    const action = STEP_ACTIONS[step];
    if (!action) {
      log(`[auth] waiting: unrecognised page at step ${steps} (${state.url.slice(0, 80)})`);
      continue;
    }
    if (step === "password") sawPassword = true;
    try {
      await action(page, credentials, log);
    } catch (error) {
      if (error instanceof LoginError) throw error;
      const message = error instanceof Error ? error.message : String(error);
      await screenshotOnFailure(page, `step-${step}`, log);
      throw new LoginError("login_step_failed", `${step} step failed: ${message.slice(0, 200)}`);
    }
  }
  await screenshotOnFailure(page, "max-steps", log);
  throw new LoginError(
    "login_max_steps",
    `sign-in did not converge after ${maxSteps} steps`,
    "the login flow changed; capture a screenshot and update src/auth/state.ts",
  );
}

/** Drive the sign-in pages until the OWA shell is up. */
async function signIn(
  page: Page,
  credentials: Credentials,
  options: { deadline: number; maxSteps: number; stepIntervalMs: number },
  log: LogFn,
): Promise<{ steps: number; warm: boolean }> {
  const remaining = Math.max(1, options.deadline - Date.now());
  return driveSignIn(page, credentials, {
    ready: (state) => classifySignInPage(state).step === "owa",
    timeoutMs: remaining,
    stepIntervalMs: options.stepIntervalMs,
    maxSteps: options.maxSteps,
    log,
  });
}

/** Read the MSAL cache out of the loaded OWA page. */
interface Harvest {
  refreshToken: string;
  entries: MsalEntry[];
  claims: RefreshClaims | undefined;
}

async function harvestMsal(page: Page): Promise<Harvest> {
  const storage = await page.evaluate(() => {
    const out: Record<string, string> = {};
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      if (key === null) continue;
      const value = localStorage.getItem(key);
      if (value !== null) out[key] = value;
    }
    return out;
  });
  const entries = parseMsalStorage(storage);
  const refresh = pickRefreshToken(entries, OWA_CLIENT_ID) ?? pickRefreshToken(entries);
  if (!refresh?.secret) {
    throw new LoginError(
      "refresh_token_missing",
      `no refresh token found in the MSAL cache (${entries.length} credential entries)`,
      "the mail shell may not have finished starting; retry, or clear the profile and log in interactively",
    );
  }
  return { refreshToken: refresh.secret, entries, claims: idTokenClaims(entries) };
}

/**
 * Redeem one audience, persisting the rotated refresh token atomically together
 * with the access token before either is returned to the caller.
 */
async function mint(
  audience: Audience,
  tenantId: string,
  refreshToken: string,
  account: TokenCache["account"],
  log: LogFn,
): Promise<{ result: RefreshResult; rotated: string; ms: number }> {
  const started = Date.now();
  const result = await refreshAccessToken({ tenantId, refreshToken, audience });
  const rotated = result.refreshToken ?? refreshToken;
  await persistTokenUpdate({
    audience,
    access_token: result.accessToken,
    expires_at: result.expiresAt,
    ...(result.scope ? { scope: result.scope } : {}),
    refresh_token: rotated,
    source: "browser",
    ...(account ? { account } : {}),
  });
  const ms = Date.now() - started;
  log(
    `[auth] ${audience}: aud=${result.claims?.aud ?? "?"} ttl=${result.expiresIn}s rt=${redactSecret(rotated)} rotated=${result.refreshToken !== undefined} stored`,
  );
  return { result, rotated, ms };
}

/** Tenant id from a previously cached token's claims, when the id token lacks it. */
async function tenantFromCache(): Promise<string | undefined> {
  const cached = await readTokenCache();
  for (const entry of Object.values(cached.cache?.tokens ?? {})) {
    const tid = entry ? decodeJwtClaims(entry.access_token)?.tid : undefined;
    if (tid) return tid;
  }
  return undefined;
}

/**
 * Full flow: sign in (interactively only if the profile is cold), harvest the
 * refresh token, mint both audiences, persist the canonical cache.
 */
export async function performLogin(options: PerformLoginOptions = {}): Promise<LoginOutcome> {
  const log = options.log ?? stderrLog;
  const now = options.now ?? Date.now;
  const timeoutMs = options.timeoutMs ?? 180_000;
  const stepIntervalMs = options.stepIntervalMs ?? 1500;
  const maxSteps = options.maxSteps ?? 20;
  const harvestDelayMs = options.harvestDelayMs ?? 8000;
  const started = now();
  const profileDir = config().profileDir;

  const credentials = await readCredentials();
  log(`[auth] credentials from ${credentials.origin}; profile ${profileDir}`);

  const session = await openBrowserSession({ profileDir }, log);
  let shellMs = 0;
  let harvestMs = 0;
  try {
    log(`[auth] opening ${OWA_URL}`);
    await session.page
      .goto(OWA_URL, { waitUntil: "domcontentloaded", timeout: 90_000 })
      .catch((error: unknown) => log(`[auth] navigation: ${error instanceof Error ? error.message : String(error)}`));

    const signedIn = await signIn(session.page, credentials, { maxSteps, stepIntervalMs, deadline: now() + timeoutMs }, log);
    const shellAt = now();
    shellMs = shellAt - started;
    log(`[auth] OWA shell ready after ${signedIn.steps} step(s), ${shellMs}ms (${signedIn.warm ? "warm profile" : "cold profile: password used"})`);

    // The shell renders before its API calls; MSAL writes the cache a moment later.
    await delay(harvestDelayMs);
    const harvestAt = now();
    const harvested = await harvestMsal(session.page);
    harvestMs = now() - harvestAt;
    const cachedAudiences = cachedAccessTokenAudiences(harvested.entries);
    log(
      `[auth] harvested ${harvested.entries.length} MSAL entries; cached audiences [${cachedAudiences.join(", ")}]; refresh token ${redactSecret(harvested.refreshToken)}`,
    );

    const tenantId = harvested.claims?.tid ?? (await tenantFromCache());
    if (!tenantId) {
      throw new LoginError(
        "tenant_unknown",
        "could not determine the tenant id from the id token",
        "the MSAL idtoken entry was missing; clear the profile and log in interactively",
      );
    }
    const account = {
      ...(harvested.claims?.preferred_username ? { upn: harvested.claims.preferred_username } : {}),
      ...(harvested.claims?.oid ? { id: harvested.claims.oid } : {}),
      ...(harvested.claims?.upn && !harvested.claims?.preferred_username ? { upn: harvested.claims.upn } : {}),
    };

    // Outlook first: mail/calendar is the primary audience, so a crash after this
    // point still leaves a usable cache.
    const outlook = await mint("outlook", tenantId, harvested.refreshToken, account, log);
    const graph = await mint("graph", tenantId, outlook.rotated, account, log);

    const persisted = await readTokenCache();
    if (!persisted.cache) {
      throw new LoginError("cache_write_failed", "token cache was not readable immediately after writing");
    }
    return {
      cache: persisted.cache,
      timings: {
        totalMs: now() - started,
        shellMs,
        harvestMs,
        outlookRefreshMs: outlook.ms,
        graphRefreshMs: graph.ms,
        steps: signedIn.steps,
      },
      warm: signedIn.warm,
      credentialsOrigin: credentials.origin,
      cachedAudiences,
      profileDir,
      path: persisted.path,
    };
  } finally {
    await session.close();
  }
}
