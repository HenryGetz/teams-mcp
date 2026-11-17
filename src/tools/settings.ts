/**
 * `settings` tool: read/write the user-editable config.yaml.
 *
 * Ops: get (whole effective config or one dot-path) | set (dot-path + value,
 * validated before write) | reload (re-read from disk) | validate | path.
 */
import { z } from "zod";
import { ConfigError } from "../config.js";
import {
  getSetting,
  reloadSettings,
  settingsPath,
  setSetting,
  SettingsError,
  validateSettings,
} from "../settings.js";
import { auditWrite, auditedWrite, readOnlyGuard, refundWriteToken, writeRateGuard } from "../safety.js";
import { errorResult, opOf, stringArg, textResult, type ToolModule } from "./shared.js";

const OPS = ["get", "set", "reload", "validate", "path"] as const;

export const settingsTool: ToolModule = {
  name: "settings",
  title: "Settings",
  description:
    "Read/write ~/.config/teams-mcp/config.yaml. op=get (effective values or key=dot.path) | set (validated before write; blocked in read-only mode; safety.*, teams.allowed_targets and invite_allowlist are operator-only and cannot be changed from this tool) | reload | validate | path. Keys: timezone, work_hours.{days,start,end}, default_duration_min, aliases.<name>, invite_allowlist, blocked_times, snapshot_dir, cache.*_ttl_minutes, test_event_prefix, max_results.*, safety.read_only, safety.unsubscribe.{enabled,daily_cap}, safety.write_rate_limit_per_minute. Env vars override the file and are reported as env_override.",
  shape: {
    op: z.enum(OPS),
    key: z.string().optional().describe("dot path for get/set, e.g. work_hours.start or aliases.joel"),
    value: z
      .union([z.string(), z.number(), z.boolean(), z.array(z.unknown()), z.record(z.string(), z.unknown())])
      .optional()
      .describe("set: new value; strings are coerced to numbers/booleans/lists"),
  },
  async handle(args) {
    const op = opOf(args, OPS);
    try {
      const blocked = readOnlyGuard("settings", op, args);
      if (blocked) return blocked;
      if (op === "path") return textResult(settingsPath());
      if (op === "get") return textResult(await getSetting(stringArg(args, "key")));
      if (op === "reload") return textResult(await reloadSettings());
      if (op === "validate") return textResult(await validateSettings());
      if (op === "set") {
        const key = stringArg(args, "key");
        if (!key) return errorResult("settings_invalid", "key is required for op=set");
        if (args.value === undefined) return errorResult("settings_invalid", "value is required for op=set");
        // The model must not be able to relax its own safety policy: enabling
        // unsubscribe, raising the daily cap or the write rate, flipping
        // read_only, or widening the send allowlist is an operator action
        // through config.yaml/env/the CLI.
        if (
          key === "safety" ||
          key.startsWith("safety.") ||
          key === "teams" ||
          key.startsWith("teams.") ||
          key === "invite_allowlist"
        ) {
          auditWrite({
            tool: "settings",
            op,
            target: { key },
            outcome: "safety_key_protected",
          });
          return errorResult(
            "settings_invalid",
            `${key} is operator-owned; the model-facing settings tool cannot change safety.* or teams.allowed_targets`,
            "edit config.yaml, set TEAMS_MCP_* env vars, or use `node dist/cli.js settings set`; write attempts were audited",
          );
        }
        const limited = writeRateGuard("settings", op, { key });
        if (limited) return limited;
        try {
          const written = await auditedWrite("settings", op, { key }, () => setSetting(key, args.value));
          return textResult(written);
        } catch (error) {
          // A rejected key/value wrote nothing: give the token back.
          if (error instanceof SettingsError) refundWriteToken();
          throw error;
        }
      }
      return errorResult("settings_invalid", `unknown op "${op}"`, `available ops: ${OPS.join(", ")}`);
    } catch (error) {
      // A bad key or an invalid file is a tool error, not a crash the SDK has to
      // wrap into an untyped message.
      if (error instanceof SettingsError) return errorResult(error.code, error.message);
      if (error instanceof ConfigError) return errorResult(error.code, error.message);
      throw error;
    }
  },
};
