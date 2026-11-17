# Authentication

How `teams-mcp` obtains and keeps a working Microsoft credential. Every claim here was
measured against the live account (`user@example.com`, tenant
`<tenant-id>`); the raw recon evidence is recorded in the
development checkout (recon notes and probe artifacts) and is not shipped in this
package.

---

## 1. Why there are two audiences

This tenant blocks every general-purpose client (`AADSTS50105`), so the only usable client
is Outlook Web itself: `9199bf20-a13f-4107-85dc-02114787ef48`. One rotating refresh token
from that client mints two different access tokens:

| Audience | Scope requested | Used for | Scopes |
|---|---|---|---|
| `outlook` | `https://outlook.office.com/.default` | **mail, calendar, free/busy** — Outlook REST v2.0 | 74, incl. `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite` |
| `graph` | `https://graph.microsoft.com/.default` | **directory only** — `GET /users?$search` | 26, incl. `Directory.Read.All`, `People.Read` |

The Graph token cannot read mail or calendar (403 `ErrorAccessDenied`) and asking for
explicit Graph scopes is refused (`AADSTS65002`). Do not try to widen it; use the right
audience.

```
                 ┌──────────────────────────────┐
                 │ credentials.json (0600)      │  cold profile only
                 │ or TEAMS_MCP_USERNAME/…      │
                 └──────────────┬───────────────┘
                                │  password + TOTP
┌───────────────┐   warm: one   ▼
│ browser-profile│──click──► OWA login ──► localStorage MSAL cache
└───────────────┘                              │ refresh token (rotates)
                                               ▼
                            POST login.microsoftonline.com/{tid}/oauth2/v2.0/token
                              grant_type=refresh_token  client_id=9199bf20-…
                              Origin: https://outlook.office.com   ← required
                                   │                       │
                                   ▼                       ▼
                          tokens.outlook            tokens.graph
                                   └────────┬──────────────┘
                                            ▼
                          ~/.config/teams-mcp/token.json  (0600, atomic)
```

---

## 2. Module map

| File | Responsibility |
|---|---|
| `src/auth/store.ts` | Canonical token cache (read/validate/write), orchestration `getAccessToken` / `forceLogin` / `logout`, cross-process lock, legacy document normalisation |
| `src/auth/refresh.ts` | Refresh-token grant: request builder, typed errors, response parsing, JWT claim decode |
| `src/auth/login.ts` | Playwright persistent-profile sign-in, MSAL harvest, minting both audiences (also exports the shared browser/sign-in driver) |
| `src/auth/teams-web.ts` | Teams-web Graph credential: MSAL v5 decrypt, refresh, silent harvest, offline status |
| `src/auth/lock.ts` | Shared cross-process lock used by both credential caches |
| `src/auth/state.ts` | Pure sign-in page classifier (`classifySignInPage`) |
| `src/auth/msal-cache.ts` | Pure localStorage → MSAL credential parser |
| `src/auth/totp.ts` | RFC 6238 TOTP (SHA-1 / 6 digits / 30 s) |
| `src/auth/xvfb.ts` | Display detection and Xvfb lifecycle |
| `src/tools/auth.ts` | MCP `auth` tool: `status` / `login` / `logout` |
| `src/cli.ts` | `node dist/cli.js auth status|login|logout` |

---

## 3. Teams-web credential (separate file)

Channel messages need a second credential: the Graph token minted by the Teams
web client (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`), because the OWA Graph
token has no `ChannelMessage.Read.All`. It lives in its own file and lock:

```
~/.config/teams-mcp/teams-web-token.json  (0600, atomic)
~/.config/teams-mcp/teams-web-token.json.lock
{version, client_id, tenant_id?, account?, scope, access_token, expires_at,
 refresh_token?, source: browser|refresh|access-only, updated_at}
```

`src/auth/teams-web.ts` refreshes it headlessly with the standard refresh grant
(`Origin: https://teams.microsoft.com`, scope `https://graph.microsoft.com/.default`,
5-minute skew, rotation persisted before use). If the refresh token is missing
or dead it harvests from the Teams web app in the same persistent profile: the
MSAL v5 refresh-token entry is AES-GCM encrypted with a key in the
`msal.cache.encryption` cookie (HKDF over the cookie key, zero IV); the decrypted
token is redeemed and persisted. When decryption is impossible it stores the
cached Graph access token as `source: access-only`. `auth.logout` removes both
credential files and keeps the browser profile.

Details and the observed live transcript are kept in the operator's private notes.

---

## 4. Token cache

`~/.config/teams-mcp/token.json`, mode `0600`, written atomically
(`tmp` + `rename`, `0600` on the temp file):

```json
{
  "version": 1,
  "account": { "upn": "user@example.com", "id": "…" },
  "refresh_token": "<rotating, single-use>",
  "tokens": {
    "outlook": { "access_token": "…", "expires_at": "2026-01-15T22:44:13.501Z", "scope": "…74 scopes…" },
    "graph":   { "access_token": "…", "expires_at": "2026-01-15T23:02:34.023Z", "scope": "…26 scopes…" }
  },
  "source": "browser | refresh | headless",
  "updated_at": "2026-01-15T21:39:34.663Z"
}
```

Legacy documents are still **read** (never rewritten): `tokens.json` in the OWA/MSAL capture
shape, and the earlier single-audience v1 with a top-level `access_token` (treated as
`tokens.outlook`).

### Expiry margin

A token is treated as stale **5 minutes before** `expires_at` (`TOKEN_EXPIRY_SKEW_MS`).
`getAccessToken` will never return something that dies inside that window; if a refresh
comes back with a shorter life than the margin it raises an error instead of returning it.

### Refresh-token rotation — the ordering that matters

Every redemption returns a **new** refresh token and invalidates the old one. The code
therefore:

1. takes the lock (§4),
2. re-reads `token.json` (another process may have rotated while we waited),
3. redeems the newest refresh token,
4. persists `refresh_token` **together with** the new access token in one atomic write,
5. only then returns the access token.

If step 3 succeeds but step 4 fails, the session is lost — which is why the write happens
before the token is handed to any caller.

---

## 5. Cross-process locking

Refresh-token rotation means two `teams-mcp` processes (or the CLI plus a server) refreshing
concurrently would invalidate each other. So:

* Lock file: `~/.config/teams-mcp/token.lock`, content `{"pid":…,"at":"…"}`, mode `0600`,
  created with `open(path, "wx")` (fails if present).
* Held **only around refresh** — a 60-second browser login never holds it.
* After acquiring, the cache is re-read; if a fresh token is already there (another process
  won the race) the lock is released and that token is used, with no second redemption.
* Stale after **120 s** (a killed process cannot wedge the tool): the lock is taken over.
* If a live lock blocks us for **120 s**, the call fails with an actionable error naming the
  lock path, rather than spinning:

  ```
  auth_required: token lock …/token.lock is still held after 120s (lock is 118s old)
    another teams-mcp process is refreshing the session; retry shortly, or delete the lock
    file if nothing else is running
  ```

* Even without the lock (e.g. taken over), correctness holds: after a failed refresh the code
  re-reads the cache and adopts a concurrently-written fresh token before reporting failure.

---

## 6. Failure handling

`getAccessToken({ audience })` in order:

1. **Fresh cache** → return.
2. **Refresh** (cache has a refresh token). On failure:
   * re-read the cache — if someone else refreshed, adopt that token;
   * `invalid_grant` / `AADSTS70008` / `AADSTS700082` / `AADSTS50173` / `AADSTS9002313` /
     `AADSTS70000` → the grant is dead → **one** sign-in attempt;
   * transient (`network_error`, `timeout`, HTTP 5xx) → **one** retry, then a *retryable*
     error. **The cache is never wiped** and no browser is launched; a network blip must not
     cost an interactive login.
3. **Sign in** (no refresh token, or the grant is dead) → the browser flow, once.

Every failure surfaces as `AuthRequiredError` (or `LoginError` from the flow) carrying a
`hint`; nothing leaks a stack trace to an MCP client.

### Error table

| `code` | Meaning | What to do |
|---|---|---|
| `auth_required` | no usable credential after refresh/login | `auth login` |
| `credentials_missing` | no `credentials.json` and no env vars | create it, `chmod 600` |
| `credentials_incomplete` | file lacks `username`/`password` | fix the file |
| `totp_secret_missing` | MFA asked for a code, no secret configured | add `totpSecret` |
| `signin_rejected` | page showed an `AADSTS…` code (e.g. `AADSTS50105`) | see the code + screenshot path in the log |
| `mfa_setup_required` | unfinished MFA registration | sign in once interactively with the same profile |
| `login_timeout` / `login_max_steps` | flow did not converge | capture the screenshot, update `src/auth/state.ts` |
| `refresh_token_missing` | MSAL cache had no refresh token | retry; clear profile if it persists |
| `xvfb_unavailable` | headless host without Xvfb | `apt-get install xvfb`, or set `DISPLAY` |
| `tenant_unknown` | id token lacked `tid` and no cached token had one | clear profile, log in again |

---

## 7. The sign-in flow

Runs headed Chromium (headless is fingerprinted by the login front end) on the persistent
profile, under Xvfb when `DISPLAY` is unset.

1. `GET https://outlook.office.com/mail/`.
2. Classify the page each step (`src/auth/state.ts`) and act:

   | Step | Signal | Action |
   |---|---|---|
   | `owa` | OWA shell selector **and** OWA host | done |
   | `aadsts` | `AADSTS\d{5,7}` in body | abort + screenshot |
   | `consent` | `#idBtn_Accept` | click |
   | `totp` | `#idTxtBcOtc` / `input[name=otc]` / `autocomplete=one-time-code` | fill code, submit |
   | `mfa-switch` | "Use a verification code" / "I can't use my Microsoft Authenticator app right now" | click |
   | `account-picker` | body "Pick an account" (warm profile) | click the tile — **never** a password |
   | `kmsi` | "Stay signed in?" | Yes |
   | `password` | body "Enter password" **or** button value `Sign in` | fill password |
   | `identifier` | button value `Next` | fill e-mail (only if empty) |
   | `mfa-setup` | "More information required" | abort, actionable error |

   The identifier and password fields are **both** in the DOM on every step, so the branch is
   on the submit button's *label*, never on field presence.
3. Wait 8 s for MSAL to finish writing, then read `localStorage`, parse the MSAL entries, take
   the `RefreshToken` for client `9199bf20…` plus the `IdToken` claims (`preferred_username`,
   `oid`, `tid`).
4. Mint `outlook` then `graph`, persisting rotation after each. Outlook is first so a crash
   still leaves the primary audience usable.
5. Close the browser, stop Xvfb, return timings.

TOTP codes are refused with under 5 s of window left (the submission would land after
rollover); the flow waits for the next window instead.

### Measured timings (this machine, warm profile)

| Operation | Time |
|---|---|
| `forceLogin` end-to-end (launch → OWA shell → harvest → both mints) | **15.1–16.7 s** |
| `getAccessToken` cache hit | ~0.3 s |
| refresh via token endpoint (network, rotation persisted) | **281 ms** |
| `GET /me` with the outlook token | 189–225 ms |
| `GET /users?$search` with the graph token | 186–329 ms |
| MCP stdio server boot + `auth status` | 2.0 s |

Access-token lifetimes observed: outlook **4 546–5 351 s** (~75–90 min), graph **4 119–4 184 s**
(~70 min) — the endpoint varies per mint, so never assume a fixed TTL; always read
`expires_in`/`expires_at`.

---

## 8. Security

* **Credentials** live only in `~/.config/teams-mcp/credentials.json` (mode `0600`), or in
  `TEAMS_MCP_USERNAME` / `TEAMS_MCP_PASSWORD` / `TEAMS_MCP_TOTP_SECRET`. A looser mode logs a
  warning; values are never logged.
* **Browser profile** `~/.config/teams-mcp/browser-profile` (mode `0700`) holds the SSO
  cookies. It is the reason a re-login costs one click, so `auth logout` deliberately keeps
  it and removes only `token.json`.
* **Tokens are never printed.** Logs show `first6…last6 (len N)`; `auth status` reports
  expiries and scope *counts*, not scope strings.
* **No secrets in the repo.** `src/` and `docs/` contain no credential material; the recon
  artifacts are redacted.
* **Writes are atomic** (`tmp` + `rename`) and `0600`, so a crash cannot leave a half-written
  cache or a world-readable token.
* The refresh call has **no client secret** (public SPA client) and always sends the `Origin`
  header; without it the endpoint answers `AADSTS9002327`.

---

## 9. Usage

### MCP tool

```jsonc
{ "op": "status" }                              // offline, one file read — safe to poll
{ "op": "login" }                               // 15–60 s; writes the token cache
{ "op": "login", "audience": "graph", "force": true }
{ "op": "logout" }                              // removes token.json + teams-web-token.json, keeps the profile
```

`status` output (abridged):

```json
{
  "authenticated": true,
  "account": "user@example.com",
  "source": "refresh",
  "audiences": [
    { "audience": "outlook", "expires_at": "2026-01-15T22:44:13.501Z", "fresh": true, "seconds_left": 3860, "scope_count": 74 },
    { "audience": "graph",   "expires_at": "2026-01-15T23:02:34.023Z", "fresh": true, "seconds_left": 4961, "scope_count": 26 }
  ],
  "has_refresh_token": true,
  "cache_format": "v1",
  "token_path": "/home/<user>/.config/teams-mcp/token.json",
  "profile": "/home/<user>/.config/teams-mcp/browser-profile"
}
```

### CLI

```bash
node dist/cli.js auth status     # cache only, exit 0 when the outlook token is fresh
node dist/cli.js auth login      # warm profile: one click; prints account + both expiries
node dist/cli.js auth login --audience graph --force
node dist/cli.js auth logout     # removes token.json only
```

`auth login` prints `elapsed_ms` so a slow interactive login is visible.

### Re-login / rotating credentials

* **Session expired but profile warm** — just call `auth login` (or any tool that needs a
  token): one click, ~15 s.
* **Password changed / MFA reset** — update `credentials.json`, then `auth login`. The profile
  may need clearing if the old session is still cached:
  `rm -rf ~/.config/teams-mcp/browser-profile`.
* **Fully cold machine** — write `credentials.json` (`chmod 600`) with `username`, `password`,
  `totpSecret`; `auth login` handles password + TOTP. On a headless host install `xvfb`.
* **Concurrent processes** — nothing to do; the lock and cache re-reads handle it.

Credentials file shape:

```json
{ "username": "…", "password": "…", "totpSecret": "<base32>" }
```

---

## 10. Verification

```bash
npm run typecheck
npm test                                    # 556 tests (3 skipped), incl. RFC 6238 vectors, refresh wire
                                            # format, MSAL parsing, page classifier, Xvfb
npm run build
node scripts/dev/auth-login.mjs             # live: login → both audiences → refresh → MCP status
node scripts/dev/auth-login.mjs --offline   # plumbing only, no network
```

`auth-login.mjs` asserts, against the compiled code: the cache shape and `0600` mode, both
audiences present with the right `aud` and scopes, `GET /me` with the outlook token, the
directory lookup with the graph token, a browser-free refresh that rotates the refresh token
on disk while leaving the other audience untouched, and `auth status` over the real stdio MCP
transport. It performs exactly **one** refresh so a concurrently running agent keeps a valid
token.

---

## 11. Known limits and risks

* **SPA refresh fragility.** The session depends on the SPA-client refresh token plus the
  mandatory `Origin` header. A change to that rule breaks refresh; the browser fallback then
  carries the load.
* **Tenant policy.** Conditional Access changes (compliant device, blocked automation IPs,
  shorter session lifetimes) could make the browser flow demand interaction. None interfered
  in testing.
* **Outlook REST v2.0 is formally deprecated** (`deprecation`/`sunset` headers) but is the
  only surface the tested tenant allows for mail/calendar. It sits behind the clients in
  `src/outlook/` so it can be swapped if that ever changes.
* **Automation detection.** Headed Chromium with `navigator.webdriver` masked is what works;
  headless is not supported on purpose.
* **Xvfb** is required on headless hosts; the error names the install command.
