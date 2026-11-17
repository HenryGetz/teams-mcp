/**
 * Configuration.
 *
 * Precedence, highest first: environment variables > config.yaml > built-in defaults.
 *
 * config.yaml lives at `${TEAMS_MCP_CONFIG_DIR:-~/.config/teams-mcp}/config.yaml`,
 * is re-read whenever its mtime/size changes (hot reload per call), and is
 * validated with zod; invalid values raise `ConfigError` naming every issue.
 * A missing file is not an error — defaults apply.
 */
import { statSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { parse as parseYaml } from "yaml";
import { z } from "zod";

/** Default timezone for every date/time interpretation and rendering. */
export const DEFAULT_TIMEZONE = "America/Chicago";

/** A token is considered stale this long before it actually expires. */
export const TOKEN_EXPIRY_SKEW_MS = 5 * 60 * 1000;

export const WEEKDAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"] as const;
export type Weekday = (typeof WEEKDAYS)[number];

const TIME_OF_DAY = z
  .string()
  .regex(/^([01]\d|2[0-3]):[0-5]\d$/, "expected HH:MM (24h, zero padded)");

/**
 * True when `Intl` accepts `zone` as a time zone. Every date/time in this server
 * is rendered through `Intl.DateTimeFormat`, which throws `RangeError` for an
 * unknown zone — validating here turns a raw crash into an actionable error.
 */
export function isValidTimeZone(zone: string): boolean {
  if (!zone || zone.trim() !== zone) return false;
  try {
    new Intl.DateTimeFormat("en-US", { timeZone: zone });
    return true;
  } catch {
    return false;
  }
}

const timezoneSchema = z
  .string()
  .min(1)
  .refine(isValidTimeZone, "expected an IANA time zone, e.g. America/Chicago or UTC")
  .default(DEFAULT_TIMEZONE);

const workHoursSchema = z
  .object({
    days: z.array(z.enum(WEEKDAYS)).min(1).default(["mon", "tue", "wed", "thu", "fri"]),
    start: TIME_OF_DAY.default("09:00"),
    end: TIME_OF_DAY.default("17:00"),
  })
  .default({ days: ["mon", "tue", "wed", "thu", "fri"], start: "09:00", end: "17:00" });

const blockedTimeSchema = z.object({
  start: TIME_OF_DAY,
  end: TIME_OF_DAY,
  days: z.array(z.enum(WEEKDAYS)).optional(),
  reason: z.string().optional(),
});

const cacheSchema = z
  .object({
    own_calendar_ttl_minutes: z.number().int().min(1).max(1440).default(5),
    busy_ttl_minutes: z.number().int().min(1).max(1440).default(30),
    directory_ttl_minutes: z.number().int().min(1).max(1440).default(60),
  })
  .default({ own_calendar_ttl_minutes: 5, busy_ttl_minutes: 30, directory_ttl_minutes: 60 });

const maxResultsSchema = z
  .object({
    mail: z.number().int().min(1).max(50).default(25),
    calendar: z.number().int().min(1).max(50).default(50),
    people: z.number().int().min(1).max(25).default(25),
    slots: z.number().int().min(1).max(20).default(5),
    preview_chars: z.number().int().min(40).max(1000).default(300),
  })
  .default({ mail: 25, calendar: 50, people: 25, slots: 5, preview_chars: 300 });

/**
 * Capability scoping for autonomous use.
 *
 * `read_only` is the single most valuable switch for an eval: it rejects every
 * mutating op before a request is made. `unsubscribe.enabled` is deliberately
 * off by default because a one-click unsubscribe is an irreversible external
 * action; `daily_cap` bounds external actions even when enabled. The
 * `write_rate_limit_per_minute` bucket covers every mutating op.
 */
const safetySchema = z
  .object({
    read_only: z.boolean().default(false),
    unsubscribe: z
      .object({
        enabled: z.boolean().default(false),
        daily_cap: z.number().int().min(0).max(1000).default(5),
      })
      .default({ enabled: false, daily_cap: 5 }),
    write_rate_limit_per_minute: z.number().int().min(1).max(600).default(30),
  })
  .default({
    read_only: false,
    unsubscribe: { enabled: false, daily_cap: 5 },
    write_rate_limit_per_minute: 30,
  });

/**
 * Operator allowlist for sending into chats/channels this server did not create:
 * `chat:<chat_id>` or `channel:<team_id>/<channel_id>`. The model never gets a
 * flag for this; a target is writable only when the operator listed it.
 */
export const teamsTargetSchema = z
  .string()
  .regex(
    /^(chat:.+|channel:[^/]+\/.+)$/i,
    "expected chat:<chat_id> or channel:<team_id>/<channel_id>",
  );

const teamsSchema = z
  .object({ allowed_targets: z.array(teamsTargetSchema).default([]) })
  .default({ allowed_targets: [] });

const fileSchema = z.object({
  timezone: timezoneSchema,
  work_hours: workHoursSchema,
  default_duration_min: z.number().int().min(5).max(480).default(30),
  /** Friendly name -> email address, resolved before any Graph lookup. */
  aliases: z.record(z.string(), z.string().min(3)).default({}),
  /** Emails or domain suffixes ("@example.com"); empty = no restriction. */
  invite_allowlist: z.array(z.string().min(3)).default([]),
  blocked_times: z.array(blockedTimeSchema).default([]),
  snapshot_dir: z.string().min(1).optional(),
  cache: cacheSchema,
  test_event_prefix: z.string().default("[test]"),
  max_results: maxResultsSchema,
  safety: safetySchema,
  teams: teamsSchema,
});

type FileConfig = z.output<typeof fileSchema>;

export interface BlockedTime {
  start: string;
  end: string;
  days?: Weekday[];
  reason?: string;
}

export interface WorkHours {
  days: Weekday[];
  start: string;
  end: string;
}

export interface ResolvedConfig {
  configDir: string;
  configFile: string;
  tokenPath: string;
  /** Teams-web (Graph, client 5e3ce6c0…) token cache, separate from token.json. */
  teamsWebTokenPath: string;
  /** Ids of chats/channels this server created, so send can default to them. */
  teamsStatePath: string;
  /** Read-only fallback documents written by earlier tooling (OWA/MSAL shape). */
  legacyTokenPaths: string[];
  /** Persistent Chromium profile used for silent re-auth. */
  profileDir: string;
  cacheDir: string;
  snapshotDir: string;
  snapshotPath: string;
  /** Outlook REST v2.0 root — primary API for mail, calendar and scheduling. */
  outlookBaseUrl: string;
  /** Graph root — directory (`/users`) only in this tenant. */
  graphBaseUrl: string;
  timezone: string;
  requestTimeoutMs: number;
  maxAttempts: number;
  workHours: WorkHours;
  defaultDurationMin: number;
  aliases: Record<string, string>;
  inviteAllowlist: string[];
  blockedTimes: BlockedTime[];
  testEventPrefix: string;
  cache: { ownCalendarTtlSeconds: number; busyTtlSeconds: number; directoryTtlSeconds: number };
  maxResults: { mail: number; calendar: number; people: number; slots: number; previewChars: number };
  /** Reject every mutating op before any network call. */
  readOnly: boolean;
  unsubscribe: { enabled: boolean; dailyCap: number };
  /** One global token bucket for every mutating op (mail/calendar/settings/auth). */
  writeRateLimitPerMinute: number;
  /** Operator-owned pre-existing send targets (`chat:<id>` / `channel:<team>/<channel>`). */
  teams: { allowedTargets: string[] };
}

export class ConfigError extends Error {
  readonly code = "config_invalid";
  constructor(readonly path: string, readonly issues: string[]) {
    super(`invalid config ${path}: ${issues.join("; ")}`);
    this.name = "ConfigError";
  }
}

function envValue(name: string): string | undefined {
  const raw = process.env[name];
  if (raw === undefined) return undefined;
  const trimmed = raw.trim();
  return trimmed === "" ? undefined : trimmed;
}

function envNumber(name: string): number | undefined {
  const raw = envValue(name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

const TRUE_WORDS = ["1", "true", "yes", "on", "enabled"];
const FALSE_WORDS = ["0", "false", "no", "off", "disabled"];

/**
 * Boolean env override. An unrecognised value is a config error rather than a
 * silent fallback: a typo in `TEAMS_MCP_READ_ONLY` must not quietly leave writes
 * enabled.
 */
function envFlag(name: string): boolean | undefined {
  const raw = envValue(name);
  if (raw === undefined) return undefined;
  const value = raw.toLowerCase();
  if (TRUE_WORDS.includes(value)) return true;
  if (FALSE_WORDS.includes(value)) return false;
  throw new ConfigError("<env>", [`${name} must be one of ${[...TRUE_WORDS, ...FALSE_WORDS].join(", ")}`]);
}

/** Integer env override, range-checked like the YAML schema. */
function envInteger(name: string, min: number, max: number): number | undefined {
  const raw = envValue(name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new ConfigError("<env>", [`${name} must be an integer between ${min} and ${max}`]);
  }
  return value;
}

function expandHome(path: string): string {
  if (path === "~") return homedir();
  if (path.startsWith("~/")) return join(homedir(), path.slice(2));
  return path;
}

export function configDir(): string {
  return expandHome(envValue("TEAMS_MCP_CONFIG_DIR") ?? join(homedir(), ".config", "teams-mcp"));
}

export function configFile(): string {
  return expandHome(envValue("TEAMS_MCP_CONFIG_FILE") ?? join(configDir(), "config.yaml"));
}

/** Parse and validate YAML text; missing sections fall back to defaults. */
export function parseConfigYaml(text: string, path = "<string>"): FileConfig {
  let raw: unknown;
  try {
    raw = parseYaml(text) ?? {};
  } catch (error) {
    throw new ConfigError(path, [`YAML parse error: ${(error as Error).message}`]);
  }
  const result = fileSchema.safeParse(raw);
  if (!result.success) {
    throw new ConfigError(
      path,
      result.error.issues.map((issue) => `${issue.path.join(".") || "<root>"}: ${issue.message}`),
    );
  }
  return result.data;
}

/**
 * Canonical form of an allowlist entry: a lower-case `chat:`/`channel:` prefix
 * and a trimmed id/pair. Ids themselves keep their case.
 */
export function normalizeTeamsTarget(value: string): string {
  const trimmed = value.trim();
  const separator = trimmed.indexOf(":");
  if (separator < 0) return trimmed;
  return `${trimmed.slice(0, separator).toLowerCase()}:${trimmed.slice(separator + 1).trim()}`;
}

let cache: { path: string; mtimeMs: number; size: number; value: ResolvedConfig } | undefined;

/** Forget the memoised config (tests, or after writing config.yaml). */
export function resetConfigCache(): void {
  cache = undefined;
}

/**
 * Fully resolved configuration. Revalidates the YAML file when it changes on
 * disk, so edits take effect without restarting the server.
 */
export function config(): ResolvedConfig {
  const path = configFile();
  const dir = configDir();
  let stat: { mtimeMs: number; size: number } | undefined;
  try {
    const s = statSync(path);
    stat = { mtimeMs: s.mtimeMs, size: s.size };
  } catch {
    stat = undefined;
  }

  // The path is part of the key: switching TEAMS_MCP_CONFIG_DIR with no
  // config.yaml in either place otherwise hands back the previous dir's
  // tokenPath/profileDir.
  if (
    cache &&
    cache.path === path &&
    cache.mtimeMs === (stat?.mtimeMs ?? -1) &&
    cache.size === (stat?.size ?? -1)
  ) {
    return cache.value;
  }

  let file: FileConfig;
  if (stat) {
    file = parseConfigYaml(readFileSync(path, "utf8"), path);
  } else {
    file = fileSchema.parse({});
  }

  const cacheDir = expandHome(envValue("TEAMS_MCP_CACHE_DIR") ?? join(envValue("XDG_CACHE_HOME") ?? join(homedir(), ".cache"), "teams-mcp"));
  const snapshotDir = expandHome(envValue("TEAMS_MCP_SNAPSHOT_DIR") ?? file.snapshot_dir ?? join(cacheDir, "snapshots"));

  // The env override bypasses `fileSchema`; validate it too, so a bad
  // TEAMS_MCP_TIMEZONE fails with a named issue instead of a RangeError from
  // Intl inside the first calendar/availability call.
  const timezone = envValue("TEAMS_MCP_TIMEZONE") ?? file.timezone;
  if (!isValidTimeZone(timezone)) {
    throw new ConfigError(path, [
      `timezone: "${timezone}" is not an IANA time zone (check TEAMS_MCP_TIMEZONE)`,
    ]);
  }

  // Env override for the send allowlist; entries are validated like the YAML,
  // because a typo here silently makes a target unwritable.
  let allowedTargets = file.teams.allowed_targets;
  const envTargets = envValue("TEAMS_MCP_TEAMS_ALLOWED_TARGETS");
  if (envTargets !== undefined) {
    try {
      allowedTargets = z
        .array(teamsTargetSchema)
        .parse(envTargets.split(",").map((entry) => entry.trim()).filter(Boolean));
    } catch {
      throw new ConfigError(path, [
        "TEAMS_MCP_TEAMS_ALLOWED_TARGETS entries must be chat:<chat_id> or channel:<team_id>/<channel_id>",
      ]);
    }
  }
  allowedTargets = allowedTargets.map(normalizeTeamsTarget);

  const value: ResolvedConfig = {
    configDir: dir,
    configFile: path,
    tokenPath: join(dir, "token.json"),
    teamsWebTokenPath: join(dir, "teams-web-token.json"),
    teamsStatePath: join(dir, "teams-state.json"),
    legacyTokenPaths: [
      join(dir, "tokens.json"),
      join(dir, "owa-tokens.json"),
    ],
    profileDir: expandHome(envValue("TEAMS_MCP_PROFILE_DIR") ?? join(dir, "browser-profile")),
    cacheDir,
    snapshotDir,
    snapshotPath: join(snapshotDir, "calendar.yaml"),
    outlookBaseUrl: (envValue("TEAMS_MCP_OUTLOOK_BASE_URL") ?? "https://outlook.office.com/api/v2.0").replace(/\/+$/, ""),
    graphBaseUrl: (envValue("TEAMS_MCP_GRAPH_BASE_URL") ?? "https://graph.microsoft.com/v1.0").replace(/\/+$/, ""),
    timezone,
    requestTimeoutMs: envNumber("TEAMS_MCP_REQUEST_TIMEOUT_MS") ?? 30_000,
    maxAttempts: envNumber("TEAMS_MCP_MAX_ATTEMPTS") ?? 4,
    workHours: file.work_hours,
    defaultDurationMin: envNumber("TEAMS_MCP_DEFAULT_DURATION_MIN") ?? file.default_duration_min,
    aliases: file.aliases,
    inviteAllowlist: file.invite_allowlist,
    blockedTimes: file.blocked_times,
    testEventPrefix: file.test_event_prefix,
    cache: {
      ownCalendarTtlSeconds:
        (envNumber("TEAMS_MCP_OWN_CALENDAR_TTL_MINUTES") ?? file.cache.own_calendar_ttl_minutes) * 60,
      busyTtlSeconds: (envNumber("TEAMS_MCP_BUSY_TTL_MINUTES") ?? file.cache.busy_ttl_minutes) * 60,
      directoryTtlSeconds:
        (envNumber("TEAMS_MCP_DIRECTORY_TTL_MINUTES") ?? file.cache.directory_ttl_minutes) * 60,
    },
    maxResults: {
      mail: file.max_results.mail,
      calendar: file.max_results.calendar,
      people: file.max_results.people,
      slots: file.max_results.slots,
      previewChars: file.max_results.preview_chars,
    },
    readOnly: envFlag("TEAMS_MCP_READ_ONLY") ?? file.safety.read_only,
    unsubscribe: {
      enabled: envFlag("TEAMS_MCP_UNSUBSCRIBE_ENABLED") ?? file.safety.unsubscribe.enabled,
      dailyCap:
        envInteger("TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP", 0, 1000) ?? file.safety.unsubscribe.daily_cap,
    },
    writeRateLimitPerMinute:
      envInteger("TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE", 1, 600) ??
      file.safety.write_rate_limit_per_minute,
    teams: { allowedTargets },
  };

  cache = { path, mtimeMs: stat?.mtimeMs ?? -1, size: stat?.size ?? -1, value };
  return value;
}

/** Resolve a friendly alias (or passthrough address) to an email. */
export function resolveAlias(input: string): string {
  const key = input.trim().toLowerCase();
  const aliases = config().aliases;
  return aliases[key] ?? input.trim();
}
