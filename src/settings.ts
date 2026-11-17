/**
 * User-editable settings: the YAML file at
 * `${TEAMS_MCP_CONFIG_DIR:-~/.config/teams-mcp}/config.yaml`.
 *
 * The same code backs the `settings` MCP tool and `node dist/cli.js settings …`,
 * so both paths validate identically. Writes are atomic and owner-only (0600),
 * and the resulting document is validated with the zod schema before it is
 * committed — an invalid value never reaches disk.
 *
 * Keys are YAML dot paths ("work_hours.start", "cache.busy_ttl_minutes",
 * "aliases.joel"). `get` reports *effective* values, i.e. after environment
 * overrides, and flags keys that an env var is shadowing.
 */
import { chmod, mkdir, readFile, rename, stat, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { dirname } from "node:path";
import { parse as parseYaml, stringify as stringifyYaml } from "yaml";
import {
  config,
  configFile,
  ConfigError,
  parseConfigYaml,
  resetConfigCache,
  type BlockedTime,
  type WorkHours,
} from "./config.js";
import { isoTime } from "./format.js";

export class SettingsError extends Error {
  readonly code = "settings_invalid";
  constructor(message: string, readonly issues: string[] = []) {
    super(issues.length > 0 ? `${message}: ${issues.join("; ")}` : message);
    this.name = "SettingsError";
  }
}

/** Leaf paths a user may set, mirroring the YAML schema. */
const LEAF_PATHS = [
  "timezone",
  "work_hours.days",
  "work_hours.start",
  "work_hours.end",
  "default_duration_min",
  "invite_allowlist",
  "blocked_times",
  "snapshot_dir",
  "cache.own_calendar_ttl_minutes",
  "cache.busy_ttl_minutes",
  "cache.directory_ttl_minutes",
  "test_event_prefix",
  "max_results.mail",
  "max_results.calendar",
  "max_results.people",
  "max_results.slots",
  "max_results.preview_chars",
  "safety.read_only",
  "safety.unsubscribe.enabled",
  "safety.unsubscribe.daily_cap",
  "safety.write_rate_limit_per_minute",
  "teams.allowed_targets",
] as const;

const FORBIDDEN_KEYS = ["__proto__", "constructor", "prototype"];

/** Environment variable that shadows a YAML key, if any. */
const ENV_BY_KEY: Record<string, string> = {
  timezone: "TEAMS_MCP_TIMEZONE",
  default_duration_min: "TEAMS_MCP_DEFAULT_DURATION_MIN",
  snapshot_dir: "TEAMS_MCP_SNAPSHOT_DIR",
  "cache.own_calendar_ttl_minutes": "TEAMS_MCP_OWN_CALENDAR_TTL_MINUTES",
  "cache.busy_ttl_minutes": "TEAMS_MCP_BUSY_TTL_MINUTES",
  "cache.directory_ttl_minutes": "TEAMS_MCP_DIRECTORY_TTL_MINUTES",
  "safety.read_only": "TEAMS_MCP_READ_ONLY",
  "safety.unsubscribe.enabled": "TEAMS_MCP_UNSUBSCRIBE_ENABLED",
  "safety.unsubscribe.daily_cap": "TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP",
  "safety.write_rate_limit_per_minute": "TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE",
  "teams.allowed_targets": "TEAMS_MCP_TEAMS_ALLOWED_TARGETS",
};

/** Resolved settings in YAML shape, so callers can read fields without casts. */
export interface EffectiveSettings {
  timezone: string;
  work_hours: WorkHours;
  default_duration_min: number;
  aliases: Record<string, string>;
  invite_allowlist: string[];
  blocked_times: BlockedTime[];
  snapshot_dir: string;
  cache: { own_calendar_ttl_minutes: number; busy_ttl_minutes: number; directory_ttl_minutes: number };
  test_event_prefix: string;
  max_results: { mail: number; calendar: number; people: number; slots: number; preview_chars: number };
  safety: {
    read_only: boolean;
    unsubscribe: { enabled: boolean; daily_cap: number };
    write_rate_limit_per_minute: number;
  };
  teams: { allowed_targets: string[] };
}

/** Effective configuration, shaped like the YAML document. */
export function effectiveSettings(): EffectiveSettings {
  const cfg = config();
  return {
    timezone: cfg.timezone,
    work_hours: cfg.workHours,
    default_duration_min: cfg.defaultDurationMin,
    aliases: cfg.aliases,
    invite_allowlist: cfg.inviteAllowlist,
    blocked_times: cfg.blockedTimes,
    snapshot_dir: cfg.snapshotDir,
    cache: {
      own_calendar_ttl_minutes: Math.round(cfg.cache.ownCalendarTtlSeconds / 60),
      busy_ttl_minutes: Math.round(cfg.cache.busyTtlSeconds / 60),
      directory_ttl_minutes: Math.round(cfg.cache.directoryTtlSeconds / 60),
    },
    test_event_prefix: cfg.testEventPrefix,
    max_results: {
      mail: cfg.maxResults.mail,
      calendar: cfg.maxResults.calendar,
      people: cfg.maxResults.people,
      slots: cfg.maxResults.slots,
      preview_chars: cfg.maxResults.previewChars,
    },
    safety: {
      read_only: cfg.readOnly,
      unsubscribe: { enabled: cfg.unsubscribe.enabled, daily_cap: cfg.unsubscribe.dailyCap },
      write_rate_limit_per_minute: cfg.writeRateLimitPerMinute,
    },
    teams: { allowed_targets: cfg.teams.allowedTargets },
  };
}

function assertWritableKey(key: string): void {
  if (!key.trim()) throw new SettingsError("key is required");
  const segments = key.split(".");
  if (segments.some((segment) => segment === "")) {
    throw new SettingsError(`invalid key "${key}": empty path segment`);
  }
  if (segments.some((segment) => FORBIDDEN_KEYS.includes(segment))) {
    throw new SettingsError(`invalid key "${key}": reserved segment`);
  }
  const known =
    (LEAF_PATHS as readonly string[]).includes(key) ||
    (segments.length === 2 && segments[0] === "aliases");
  if (!known) {
    throw new SettingsError(
      `unknown key "${key}"`,
      [`valid keys: ${LEAF_PATHS.join(", ")}, aliases.<name>`],
    );
  }
}

function getIn(document: unknown, key: string): unknown {
  let cursor: unknown = document;
  for (const segment of key.split(".")) {
    if (cursor === null || typeof cursor !== "object" || Array.isArray(cursor)) return undefined;
    if (!Object.hasOwn(cursor as Record<string, unknown>, segment)) return undefined;
    cursor = (cursor as Record<string, unknown>)[segment];
  }
  return cursor;
}

/** Set a dot path, creating intermediate objects. */
export function setIn(document: Record<string, unknown>, key: string, value: unknown): void {
  const segments = key.split(".");
  let cursor: Record<string, unknown> = document;
  for (const segment of segments.slice(0, -1)) {
    const next = cursor[segment];
    if (next === null || typeof next !== "object" || Array.isArray(next)) {
      cursor[segment] = {};
    }
    cursor = cursor[segment] as Record<string, unknown>;
  }
  cursor[segments[segments.length - 1] as string] = value;
}

/** Coerce CLI-style strings into YAML-native values. */
export function coerceValue(raw: unknown): unknown {
  if (typeof raw !== "string") return raw;
  const text = raw.trim();
  if (text === "true") return true;
  if (text === "false") return false;
  if (text === "null") return null;
  if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
  if ((text.startsWith("[") && text.endsWith("]")) || (text.startsWith("{") && text.endsWith("}"))) {
    try {
      return JSON.parse(text);
    } catch {
      return raw;
    }
  }
  if (text.includes(",") ) return text.split(",").map((part) => part.trim()).filter(Boolean);
  return raw;
}

async function readRawDocument(path: string): Promise<Record<string, unknown>> {
  let text: string;
  try {
    text = await readFile(path, "utf8");
  } catch {
    return {};
  }
  const parsed = parseYaml(text);
  if (parsed === null || parsed === undefined) return {};
  if (typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new SettingsError(`config file ${path} does not contain a mapping`);
  }
  return parsed as Record<string, unknown>;
}

/** Validate a raw document against the schema without touching disk. */
export function validateDocument(document: unknown): { ok: boolean; issues: string[] } {
  try {
    parseConfigYaml(stringifyYaml(document));
    return { ok: true, issues: [] };
  } catch (error) {
    if (error instanceof ConfigError) return { ok: false, issues: error.issues };
    return { ok: false, issues: [(error as Error).message] };
  }
}

export interface SettingsResult {
  config_file: string;
  key?: string;
  value?: unknown;
  settings?: EffectiveSettings;
  env_override?: string;
  valid?: boolean;
  issues?: string[];
  reloaded?: boolean;
  mtime?: string;
}

export async function getSetting(key?: string): Promise<SettingsResult> {
  const settings = effectiveSettings();
  if (!key) return { config_file: configFile(), settings };
  const value = getIn(settings, key);
  if (value === undefined) {
    throw new SettingsError(`unknown key "${key}"`, [`valid keys: ${LEAF_PATHS.join(", ")}, aliases.<name>`]);
  }
  const env = ENV_BY_KEY[key];
  return {
    config_file: configFile(),
    key,
    value,
    ...(env && process.env[env] !== undefined ? { env_override: env } : {}),
  };
}

/**
 * Serialise read-modify-write of the settings file per path.
 *
 * Two parallel `settings.set` calls used to read the same document, write the
 * same `${path}.${pid}.tmp`, and race on rename (one failed with ENOENT); even
 * with unique temp names the later write would drop the earlier key. The queue
 * makes each call see the previous result.
 */
const writeQueues = new Map<string, Promise<unknown>>();

async function serialise<T>(path: string, task: () => Promise<T>): Promise<T> {
  const prior = writeQueues.get(path) ?? Promise.resolve();
  const run = prior.then(task, task);
  const tail = run.then(
    () => undefined,
    () => undefined,
  );
  writeQueues.set(path, tail);
  try {
    return await run;
  } finally {
    if (writeQueues.get(path) === tail) writeQueues.delete(path);
  }
}

/** Set one key, validate the whole document, and persist atomically. */
export async function setSetting(key: string, rawValue: unknown): Promise<SettingsResult> {
  assertWritableKey(key);
  const path = configFile();
  return serialise(path, async () => {
    const document = await readRawDocument(path);
    const value = coerceValue(rawValue);
    setIn(document, key, value);

    const check = validateDocument(document);
    if (!check.ok) throw new SettingsError(`refusing to write invalid config (${path})`, check.issues);

    await mkdir(dirname(path), { recursive: true, mode: 0o700 });
    const tmp = `${path}.${process.pid}.${randomUUID()}.tmp`;
    await writeFile(tmp, stringifyYaml(document, { lineWidth: 200 }), { mode: 0o600 });
    await rename(tmp, path);
    await chmod(path, 0o600).catch(() => undefined);

    resetConfigCache();
    const effective = getIn(effectiveSettings(), key);
    const env = ENV_BY_KEY[key];
    return {
      config_file: path,
      key,
      value: effective,
      ...(env && process.env[env] !== undefined ? { env_override: env } : {}),
    };
  });
}

/** Drop the memoised config so the next call re-reads disk (mtime check). */
export async function reloadSettings(): Promise<SettingsResult> {
  resetConfigCache();
  const info = await stat(configFile()).catch(() => undefined);
  return {
    config_file: configFile(),
    reloaded: true,
    ...(info ? { mtime: isoTime(info.mtimeMs) } : {}),
    settings: effectiveSettings(),
  };
}

/** Validate the file on disk, reporting schema issues rather than throwing. */
export async function validateSettings(): Promise<SettingsResult> {
  const path = configFile();
  let text: string;
  try {
    text = await readFile(path, "utf8");
  } catch {
    return { config_file: path, valid: true, issues: [], settings: effectiveSettings() };
  }
  try {
    parseConfigYaml(text, path);
    return { config_file: path, valid: true, issues: [] };
  } catch (error) {
    if (error instanceof ConfigError) return { config_file: path, valid: false, issues: error.issues };
    return { config_file: path, valid: false, issues: [(error as Error).message] };
  }
}

export function settingsPath(): SettingsResult {
  return { config_file: configFile() };
}
