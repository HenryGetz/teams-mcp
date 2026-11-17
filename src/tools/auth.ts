/**
 * `auth` tool: status | login | logout.
 *
 * `status` is offline and cheap (one file read) — an agent can call it freely.
 * `login` runs the real sign-in flow (20–60s, warm profile faster) and writes
 * the canonical token cache. `logout` removes only that cache file; the browser
 * profile is preserved so the next login stays a single click.
 */
import { z } from "zod";
import { config } from "../config.js";
import { LoginError, readCredentials } from "../auth/login.js";
import {
  AuthRequiredError,
  forceLogin,
  isTokenFresh,
  logout,
  readTokenCache,
  type Audience,
} from "../auth/store.js";
import { booleanArg, errorResult, opOf, textResult, unknownOp, type ToolModule } from "./shared.js";
import { clearTeamsWebCache, teamsWebStatus } from "../auth/teams-web.js";
import { auditedWrite, readOnlyGuard, writeRateGuard } from "../safety.js";

const OPS = ["status", "login", "logout"] as const;
const AUDIENCES: readonly Audience[] = ["outlook", "graph"];

/** Seconds of usable life left after the 5-minute skew; null when not applicable. */
function secondsLeft(expiresAt: string | undefined): number | null {
  if (!expiresAt) return null;
  const expiry = Date.parse(expiresAt);
  return Number.isNaN(expiry) ? null : Math.round((expiry - Date.now()) / 1000);
}

export const authTool: ToolModule = {
  name: "auth",
  title: "Authentication",
  description:
    "Account credential state. op=status reads the local token caches offline and reports whether the outlook (mail/calendar), graph (directory) or Teams-web (channel history) credential is fresh (authenticated is true if any is), with expiry, source, account and profile path; op=login runs the sign-in flow and writes the token cache (may take 20-60s; allowed in read-only mode); op=logout deletes the cached tokens but keeps the browser profile (blocked in read-only mode).",
  shape: {
    op: z.enum(OPS),
    audience: z.enum(["outlook", "graph"]).optional().describe("login: audience to sign in for, default outlook"),
    force: z.boolean().optional().describe("login: ignore a still-valid token"),
  },
  async handle(args) {
    const op = opOf(args, OPS);
    const blocked = readOnlyGuard("auth", op, args);
    if (blocked) return blocked;

    if (op === "status") {
      const [result, web] = await Promise.all([readTokenCache(), teamsWebStatus()]);
      const cfg = config();
      const cache = result.cache;
      const audiences = cache
        ? AUDIENCES.flatMap((audience) => {
            const entry = cache.tokens[audience];
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
          })
        : [];
      // Any fresh audience can serve its ops, including the Teams-web cache,
      // which lives in its own file and backs channel history.
      const freshAudiences: string[] = audiences
        .filter((entry) => entry.fresh)
        .map((entry) => entry.audience);
      if (web.fresh) freshAudiences.push("teams_web");
      const authenticated = freshAudiences.length > 0;
      const reason = authenticated
        ? undefined
        : cache === null
          ? result.reason
          : cache.tokens.outlook || cache.tokens.graph
            ? "no cached access token is fresh; run auth.login (a refresh token may still renew it)"
            : "the token cache has no access tokens";
      return textResult({
        authenticated,
        fresh_audiences: freshAudiences,
        ...(reason ? { reason } : {}),
        ...(cache === null && result.corrupt ? { corrupt: true } : {}),
        ...(cache
          ? {
              account: cache.account?.upn,
              account_id: cache.account?.id,
              source: cache.source,
              updated_at: cache.updated_at,
              audiences,
              has_refresh_token: Boolean(cache.refresh_token),
              cache_format: result.format,
              token_path: result.path,
            }
          : { token_cache: result.path }),
        teams_web: web,
        profile: cfg.profileDir,
        ...(authenticated
          ? {}
          : { hint: "run auth.login to sign in (a warm browser profile only needs one click)" }),
      });
    }

    if (op === "logout") {
      const limited = writeRateGuard("auth", op);
      if (limited) return limited;
      const result = await auditedWrite("auth", op, undefined, async () => ({
        outlook: await logout(),
        teamsWeb: await clearTeamsWebCache(),
      }));
      return textResult({
        logged_out: result.outlook.removed,
        token_cache: result.outlook.path,
        teams_web_cache: result.teamsWeb.path,
        teams_web_logged_out: result.teamsWeb.removed,
        profile_kept: config().profileDir,
      });
    }

    if (op === "login") {
      const audience: Audience = args.audience === "graph" ? "graph" : "outlook";
      // Surface a missing credentials file before paying for a browser launch.
      try {
        await readCredentials();
      } catch (error) {
        if (error instanceof LoginError) return errorResult(error.code, error.message, error.hint);
        throw error;
      }
      // Sign-in is a credential/network operation: one shared write budget keeps
      // a loop from launching browsers or hammering the identity provider.
      const limited = writeRateGuard("auth", op);
      if (limited) return limited;
      try {
        const token = await auditedWrite("auth", op, { audience }, () =>
          forceLogin({ audience, force: booleanArg(args, "force") ?? false }),
        );
        const cache = await readTokenCache();
        return textResult({
          authenticated: true,
          audience: token.audience,
          expires_at: token.expiresAt,
          source: token.source,
          account: cache.cache?.account?.upn,
          has_refresh_token: Boolean(cache.cache?.refresh_token),
          token_path: cache.path,
        });
      } catch (error) {
        if (error instanceof LoginError) return errorResult(error.code, error.message, error.hint);
        if (error instanceof AuthRequiredError) return errorResult(error.code, error.message, error.hint);
        const message = error instanceof Error ? error.message : String(error);
        return errorResult(
          "login_failed",
          `sign-in failed: ${message}`,
          "check the screenshot path in the server log; run auth.login again or complete one interactive sign-in with the same profile",
        );
      }
    }

    return unknownOp("auth", op, OPS);
  },
};
