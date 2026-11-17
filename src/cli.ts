#!/usr/bin/env node
/**
 * CLI entrypoint:
 *   node dist/cli.js settings get|set|validate|reload|path
 *   node dist/cli.js auth status|login|logout
 *
 * Same validation and file handling as the MCP tools, so scripts and operators
 * do not need an MCP client. Writes JSON to stdout; diagnostics to stderr.
 */
import { config } from "./config.js";
import { LoginError, readCredentials } from "./auth/login.js";
import {
  AuthRequiredError,
  forceLogin,
  isTokenFresh,
  logout,
  readTokenCache,
  type Audience,
} from "./auth/store.js";
import { getSetting, reloadSettings, SettingsError, settingsPath, setSetting, validateSettings } from "./settings.js";

const USAGE = `teams-mcp

Usage:
  cli.js settings get [dot.path]
  cli.js settings set <dot.path> <value>
  cli.js settings validate
  cli.js settings reload
  cli.js settings path

  cli.js auth status
  cli.js auth login [--audience outlook|graph] [--force]
  cli.js auth logout

Examples:
  node dist/cli.js settings get timezone
  node dist/cli.js settings set work_hours.start 08:30
  node dist/cli.js auth status
  node dist/cli.js auth login
  node dist/cli.js auth login --audience graph

Precedence: environment variables > config.yaml > defaults.
auth login signs in with the persistent browser profile (a warm profile needs one
click); auth logout removes token.json only and keeps the profile.
`;

function print(value: unknown): void {
  process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
}

function secondsLeft(expiresAt: string | undefined): number | null {
  if (!expiresAt) return null;
  const expiry = Date.parse(expiresAt);
  return Number.isNaN(expiry) ? null : Math.round((expiry - Date.now()) / 1000);
}

async function authStatus(): Promise<number> {
  const result = await readTokenCache();
  const cfg = config();
  if (result.cache === null) {
    print({
      authenticated: false,
      reason: result.reason,
      token_path: result.path,
      profile: cfg.profileDir,
      hint: "run cli.js auth login",
    });
    return 1;
  }
  const audiences = (["outlook", "graph"] as const).flatMap((audience) => {
    const entry = result.cache?.tokens[audience];
    if (!entry) return [];
    return [
      {
        audience,
        expires_at: entry.expires_at,
        fresh: isTokenFresh(entry.expires_at),
        seconds_left: secondsLeft(entry.expires_at),
        scope_count: (entry.scope ?? "").split(" ").filter(Boolean).length,
      },
    ];
  });
  const primary = result.cache.tokens.outlook;
  print({
    authenticated: primary ? isTokenFresh(primary.expires_at) : false,
    account: result.cache.account?.upn,
    account_id: result.cache.account?.id,
    source: result.cache.source,
    updated_at: result.cache.updated_at,
    audiences,
    has_refresh_token: Boolean(result.cache.refresh_token),
    cache_format: result.format,
    token_path: result.path,
    profile: cfg.profileDir,
  });
  return primary && isTokenFresh(primary.expires_at) ? 0 : 1;
}

async function authLogin(argv: string[]): Promise<number> {
  const audienceArg = argv.includes("--audience") ? argv[argv.indexOf("--audience") + 1] : undefined;
  const audience: Audience = audienceArg === "graph" ? "graph" : "outlook";
  const force = argv.includes("--force");
  try {
    await readCredentials();
  } catch (error) {
    if (error instanceof LoginError) {
      process.stderr.write(`${error.code}: ${error.message}${error.hint ? `\n  ${error.hint}` : ""}\n`);
      return 3;
    }
    throw error;
  }
  const started = Date.now();
  const token = await forceLogin({ audience, force });
  const cache = await readTokenCache();
  print({
    authenticated: true,
    audience: token.audience,
    expires_at: token.expiresAt,
    source: token.source,
    account: cache.cache?.account?.upn,
    account_id: cache.cache?.account?.id,
    audiences: (["outlook", "graph"] as const).flatMap((name) => {
      const entry = cache.cache?.tokens[name];
      return entry ? [{ audience: name, expires_at: entry.expires_at, seconds_left: secondsLeft(entry.expires_at) }] : [];
    }),
    has_refresh_token: Boolean(cache.cache?.refresh_token),
    token_path: cache.path,
    elapsed_ms: Date.now() - started,
  });
  return 0;
}

async function authLogout(): Promise<number> {
  const result = await logout();
  print({ logged_out: result.removed, token_cache: result.path, profile_kept: config().profileDir });
  return result.removed ? 0 : 1;
}

async function runSettings(command: string | undefined, rest: string[]): Promise<number> {
  switch (command) {
    case "get":
      print(await getSetting(rest[0]));
      return 0;
    case "set": {
      const [key, value] = rest;
      if (!key || value === undefined) {
        process.stderr.write("set requires <dot.path> <value>\n\n");
        process.stderr.write(USAGE);
        return 2;
      }
      print(await setSetting(key, value));
      return 0;
    }
    case "validate": {
      const result = await validateSettings();
      print(result);
      return result.valid ? 0 : 1;
    }
    case "reload":
      print(await reloadSettings());
      return 0;
    case "path":
      print(settingsPath());
      return 0;
    case "--help":
    case "-h":
    case undefined:
      process.stderr.write(USAGE);
      return 0;
    default:
      process.stderr.write(`unknown settings command "${command}"\n\n${USAGE}`);
      return 2;
  }
}

async function runAuth(command: string | undefined, rest: string[]): Promise<number> {
  switch (command) {
    case "status":
      return await authStatus();
    case "login":
      return await authLogin(rest);
    case "logout":
      return await authLogout();
    case "--help":
    case "-h":
    case undefined:
      process.stderr.write(USAGE);
      return 0;
    default:
      process.stderr.write(`unknown auth command "${command}"\n\n${USAGE}`);
      return 2;
  }
}

async function run(argv: string[]): Promise<number> {
  const [group, command, ...rest] = argv;
  if (group === "settings") return await runSettings(command, rest);
  if (group === "auth") return await runAuth(command, rest);
  process.stderr.write(USAGE);
  return group === "--help" || group === "-h" || group === undefined ? 0 : 2;
}

run(process.argv.slice(2))
  .then((code) => {
    process.exitCode = code;
  })
  .catch((error: unknown) => {
    if (error instanceof SettingsError) {
      process.stderr.write(`${error.code}: ${error.message}\n`);
      process.exitCode = 3;
      return;
    }
    if (error instanceof LoginError || error instanceof AuthRequiredError) {
      process.stderr.write(`${error.code}: ${error.message}${error.hint ? `\n  ${error.hint}` : ""}\n`);
      process.exitCode = 3;
      return;
    }
    process.stderr.write(`error: ${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
