# Security

> **Unofficial.** Not affiliated with, endorsed by, or supported by Microsoft,
> University of Northwestern – St. Paul, or any employer. See
> [`DISCLAIMER.md`](DISCLAIMER.md).

teams-mcp runs as **you**, against your real mailbox and calendar, with whatever
permissions your sign-in has. It is not a sandbox and it is not read-only by
default at the library level — the shipped examples are, deliberately.

> **Live-account warning.** Every write op reaches the real service: mail is sent
> to real people, events land on real calendars and notify attendees, and
> `unsubscribe_apply` is an **irreversible external** action that a third party
> records against your address. There is no undo, no staged copy and no
> server-side dry run. Start read-only and grant writes per task.

## Threat model

| Asset | Who can reach it | Failure mode |
| --- | --- | --- |
| Mailbox contents, calendar, directory | any process holding `token.json`, and any agent wired to this server | Bulk read of personal and third-party data |
| Mailbox state | agent-issued mutating ops | Bulk move/mark, mail sent as the user, deletions, rule changes |
| Team chats, channels, channel messages | agent-issued reads and `teams.send`/`teams.create`/`teams.delete` | Posting into someone else's channel, creating or deleting channels, reading channel history |
| Third-party subscriptions | `mail.unsubscribe_apply` | Irreversible opt-out from senders the user still wants |
| Credentials | local filesystem | Token/profile theft, silent re-auth by another process |

The interesting threat is not a remote attacker; it is **agent autonomy**. A
capable model handed an open-ended goal ("get my inbox under control", "clean up
my subscriptions") will optimize the goal as stated and can act on thousands of
messages in minutes, faster than a human can interrupt, and it will report
success. Prompt-level instructions ("never delete", "only send to me") are
advisory: they live in a context window that the model writes into and that a
long run can drop. Only the tool layer is enforceable, so the safety controls
below live there and are checked **before** any network request.

Second-order risk: untrusted content. Message bodies, subjects and sender names
are attacker-controlled text that reaches the model. Treat instructions found in
mail as data, never as operator intent — this server will not widen its own
permissions because a message asked it to, and neither should a wrapper.

## Controls

### Read-only mode

`safety.read_only: true` in `config.yaml`, or `TEAMS_MCP_READ_ONLY=1` in the MCP
server env (the env var wins). When set, every mutating op returns a `read_only`
error with no request:

- `mail` `send`, `draft`, `send_draft`, `reply`, `forward`, `mark`, `move`, `unsubscribe_apply`
- `calendar` `create`, `update`, `delete`, `respond`
- `teams` `send`, `create`, `delete`
- `settings.set`, `auth.logout`

Reads, availability, the snapshot cache, `auth.status`, `auth.login` and every
Teams read (`status`, `list`, `read`, `replies`) remain available, so an agent can
still inspect and report. That includes `calendar.agenda`,
`availability.check` (one exact slot) and `availability.free` (open windows):
they are computed from reads only, take no write or rate token, and are the
supported way to answer "is 3 pm free?" and "what's my day?" while the server
is read-only. It also includes `mail.attachments` (metadata only) and
`mail.get_attachment`, which take no write token and never touch the provider's
write paths: an inline text/calendar decode is returned, and any other type is
saved under `~/.cache/teams-mcp/attachments/` (dir 0700, file 0600) and reported
by path — the server never opens or executes a downloaded file. The refusal is
audited. `mail.draft` counts as a write
even though it sends nothing (it creates an object in Drafts); `mail.send_draft`
needs `confirm: true` and re-reads the live draft so the `invite_allowlist` is
applied to the recipients it has *now*, not the ones it had when it was created.
The same rule covers the calendar: a `conflict_policy` check is a read that runs
only after the read-only/rate/allowlist guards, and `refuse` aborts before the
write.

All three shipped examples enable it (`README.md` snippets, `mcp.example.json`,
`config.example.yaml`). Enable writes only for a task you would perform yourself:
drop `TEAMS_MCP_READ_ONLY` and set `safety.read_only: false`, then put it back.
An agent cannot lift the switch: `settings.set` is blocked entirely while
read-only, and even in write mode the model-facing tool refuses `safety.*`
(see [Operator-only settings](#operator-only-settings)) — only an operator
editing the config file or using the shell CLI can change it.

### Write rate limit

`safety.write_rate_limit_per_minute` (default 30, `TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE`)
is **one global per-process token bucket covering every mutating op** — mail
`send`/`draft`/`send_draft`/`reply`/`forward`/`mark`/`move`/`unsubscribe_apply`, calendar
`create`/`update`/`delete`/`respond`, `teams.send`/`teams.create`/`teams.delete`,
`settings.set`, and `auth.login`/`auth.logout`. A token is taken before the
first request, so a `rate_limited` refusal (`retry in Ns`) performs no network
call at all; if a call is refused afterwards for another reason, the token is
refunded. It exists to bound blast radius and make bulk damage slow and visible,
not as a performance knob; lower it for autonomous runs.

### Operator-only settings

The MCP-facing `settings.set` refuses `safety.*` (`read_only`,
`unsubscribe.enabled`, `unsubscribe.daily_cap`,
`write_rate_limit_per_minute`), `invite_allowlist` and `teams.allowed_targets`
with `settings_invalid` (`operator-owned`), and journals the attempt as
`safety_key_protected`. Without this an agent in write mode could disable
read-only, turn unsubscribe on, raise the daily cap, or widen the mail/Teams send
allowlists — privilege escalation through the tool it is holding. These keys live
in the operator's hands: `config.yaml`, the `TEAMS_MCP_*` environment variables
(`TEAMS_MCP_TEAMS_ALLOWED_TARGETS` for Teams targets), or
`node dist/cli.js settings set …` from a shell.

### Unsubscribe: opt-in and capped

`safety.unsubscribe.enabled` is **off** by default (`TEAMS_MCP_UNSUBSCRIBE_ENABLED`).
While off, `unsubscribe_apply` is refused with no request. When enabled it still
requires, per call:

- `confirm: true`, and
- `expected_sender` equal to that message's From address, so a stale or
  model-invented message id cannot unsubscribe a different sender.

`safety.unsubscribe.daily_cap` (default 5, `TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP`) counts
**external actions**, not apply calls: a message carrying both an https one-click
URL and a `mailto` target consumes two slots, so 5 means at most five real
opt-outs per local day. Slots are reserved atomically immediately before acting;
if the remainder cannot cover the planned actions the call returns `cap_reached`
with the reset time and sends nothing. One call acts on at most that message's own
targets — there is no bulk path and no mail deletion.
`mail.unsubscribe_preview` is read-only and is the right way to show a human what
a target would do.

### Write audit log

`~/.cache/teams-mcp/write-audit.jsonl` (0600, append-only JSONL, rotated to
`write-audit.jsonl.1` at 5 MB) records every write attempt that reached a tool's
write path — allowed, refused by policy (confirm/sender/cap/allowlist), blocked
(read-only, rate limit) or protected (`safety_key_protected`) — with timestamp,
tool, op, an id or recipient count, and an outcome. It now also journals
`subject_mismatch` (guard refusals on calendar `update`/`delete`/`respond`),
allowlist refusals (`recipient_not_allowed`) and every `auth.login` attempt,
successful or failed, so a credential operation cannot be invisible. Every
`teams.send`/`teams.create` that reaches the write path is journalled the same
way, with the surface and target id, as is `teams.delete`. Teams policy refusals
are recorded too — `confirmation_required`, `target_not_allowed`,
`target_not_created`, `name_mismatch`, `invalid_argument`, `private_required`,
`not_owner` and `already_absent` — so an attempted deletion or an attempt to post
into a non-allowlisted target is visible even when nothing happened. Mail drafts
are journaled as writes (`draft`, `send_draft`), including `confirmation_required`
when `send_draft` is called without `confirm: true` and `no_recipients` /
`recipient_not_allowed` when the live draft fails the send-time check. Calendar
conflict refusals are journaled as `conflict_detected` (nothing written) and
`conflict_check_failed` (the calendar read itself failed). An
`already_absent` delete (404/410 before or during the request) prunes its ledger
row and writes exactly one journal line for that attempt; the DELETE is not
audited twice. Teams refusals that occur after the rate token was taken refund
it, so a rejected write does not consume the budget.
It never contains message bodies, address lists or tokens. It is the way to
reconstruct what an autonomous run did; read it after any unfamiliar agent
session.

## Teams writes and the Teams-web credential

Teams is a second, broader surface and gets its own gates on top of read-only and
the rate limit:

- `teams.send` and `teams.create` require `confirm: true`.
- `teams.create` is **channel-only**. `kind=chat` is absent from the tool schema:
  Graph requires every chat member to be `owner` or `guest` (an empty/omitted
  role or `member` is a 400), so any API-valid group chat hands every invitee
  ownership, and the tested tenant has no `Chat.ManageDeletion.All` to remove it. The
  server therefore never creates a chat. Existing chats remain readable and
  postable only when the operator allowlists the exact target.
- `teams.create kind=channel` requires `team_id`, `name` and `private: true`
  (`private_required` otherwise); it never becomes a team-wide channel. Optional
  `members` are capped at 20 (`invalid_argument` before any lookup), and the
  signed-in user must own the team (`not_owner` after one owner check). It never
  creates a team.
- Posting into a chat or channel this server did **not** create requires the exact
  target in the operator-owned `teams.allowed_targets` (`chat:<chat_id>` or
  `channel:<team_id>/<channel_id>`); otherwise the call is `target_not_allowed`
  and makes **no request**. There is no model-facing override: the MCP
  `settings.set` refuses `teams.allowed_targets`, so only `config.yaml`,
  `TEAMS_MCP_TEAMS_ALLOWED_TARGETS`, or the shell CLI can widen it. Targets
  recorded in `teams-state.json` need no entry.
- Existing chats created by earlier builds still have both users as owners. Graph
  cannot fix that here: neither credential carries `ChatMember.ReadWrite*` and a
  member `PATCH` answers 405 (v1.0) / 404 (beta), so correct roles in the Teams
  UI or after granting that permission. The diagnostic chat used during testing
  was deliberately left labelled and empty.
- `teams.delete` is channel-only and doubly restricted: `confirm: true` **and**
  `expected_name` matching the live channel name (`name_mismatch` otherwise), plus
  the channel must be in the `teams-state.json` ledger — i.e. one this server
  created. Chats and pre-existing channels cannot be deleted through this server.
  It reads the channel first, treats 404/410 as already absent, and prunes the
  ledger entry either way. It uses the OWA-derived token, not the Teams-web one.
- It never creates a team. Channel creation and deletion are the only structural
  writes.
- No attachments, mentions, message edits, bulk sends, chat deletion or channel
  deletion outside the created-by-server ledger; no channel-send loop.

**Two tokens, by design.** Chats, chat sends, team/channel metadata, and channel
create/delete use the OWA-derived `graph` token. Channel messages, replies,
rosters and channel sends use a second credential minted by the Teams web client
(`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`), because the tested tenant's OWA token has no
`ChannelMessage.Read.All`. That token is broader than mail-only use suggests
(~37 scopes, including mail), so it is worth stealing on its own.

**Silent harvest.** When no refresh token is cached, the first Teams-web call
drives the persistent Chromium profile to `teams.cloud.microsoft` and extracts the
SSO session: no password, no prompt, ~11–25 s. Any process that can run the server
and read the profile can therefore mint a fresh credential without the user's
interaction — reads included, since a channel read needs that token. Treat
`browser-profile/` as equivalent to a password, and note that `auth.logout`
deletes `token.json` **and** `teams-web-token.json` but keeps the profile, so
revocation must happen server-side too (see below).

## Credentials and files

| Path | Contents | Mode |
| --- | --- | --- |
| `~/.config/teams-mcp/` | config/credential root | 0700 |
| `~/.config/teams-mcp/token.json` | access + refresh tokens | 0600 |
| `~/.config/teams-mcp/teams-web-token.json` | Teams-web Graph credential (~37 scopes) + own lock file | 0600 |
| `~/.config/teams-mcp/teams-state.json` | chats/channels this server created (send provenance) | 0600 |
| `~/.config/teams-mcp/config.yaml` | settings incl. `safety` | 0600 on write |
| `~/.config/teams-mcp/credentials.json` | optional CI username/password/TOTP | 0600 |
| `~/.config/teams-mcp/browser-profile/` | persistent Chromium profile for silent re-auth | 0700 |
| `~/.cache/teams-mcp/` | snapshots, audit journal, unsubscribe counter | 0600 files |
| `~/.cache/teams-mcp/attachments/` | files saved by `mail.get_attachment` (sanitised names, random suffix) | dir 0700, files 0600 |

- **No secrets in the repository.** `token.json`, `teams-web-token.json`,
  `credentials.json`, the browser profile and the env vars `TEAMS_MCP_PASSWORD` /
  `TEAMS_MCP_TOTP_SECRET` (`TEAMS_MCP_USERNAME` too) are for the local machine or
  a CI secret store only.
  Check with `git status --ignored` before pushing; never paste a token into an
  issue, a bug report, or a transcript shared with a model.
- The browser profile is a credential by another name: it powers the silent
  Teams-web harvest. `teams-state.json` is not secret, but it decides which Teams
  sends need no allowlist entry, so keep it intact rather than casually deleting
  it (removing it only makes the send gate stricter).
- Run the server as a dedicated OS user on shared machines; the files above are
  only as private as the account that owns them.
- `TEAMS_MCP_PASSWORD` / `TEAMS_MCP_TOTP_SECRET` are visible to any process that
  can read the MCP client's config **and** to anything that can read
  `/proc/<pid>/environ`. Prefer an interactive login and the token cache over
  putting a password in a JSON env block.

## Incident summary

An early live-account test with an autonomous agent performed bulk mailbox
actions; the capability controls exist to prevent a repeat. The controls above
(read-only default, unsubscribe opt-in with sender binding, daily cap, write rate
limit, audit journal) are the response. Regression evidence: a cold agent run under the shipped read-only
default was refused before any request, with zero writes. Design rationale:
[`docs/design.md`](docs/design.md) §10. A red-team review of the safety layer
found and fixed four gaps — write rate limit scope, unsubscribe cap
accounting per external action, operator-only `safety.*`/`invite_allowlist`, and
audit coverage for `subject_mismatch`/allowlist/login.

## If a token or profile leaks, or an agent misbehaves

1. **Stop the process** running the server and set `TEAMS_MCP_READ_ONLY=1` in its
   configuration, so any restart comes back read-only.
2. **Read the audit journal** (`~/.cache/teams-mcp/write-audit.jsonl`) for the
   session window; note `ok` outcomes for `mail.move`, `mail.send`, `calendar.*`,
   `teams.send`, `teams.create`, `teams.delete` and `unsubscribe_apply`.
3. **Revoke server-side.** Sign the account out of all sessions / revoke the app
   consent for this client in the Microsoft account portal. That invalidates the
   refresh tokens (OWA and Teams-web); a local delete alone does not, and the
   browser profile can otherwise mint a new Teams-web token silently.
4. **Rotate locally:** delete `token.json`, `teams-web-token.json`,
   `credentials.json` and `browser-profile/`, rotate any password/TOTP secret that
   was stored in env, then log in again interactively. New files are recreated
   0600/0700.
5. **Repair what was changed** from the audit ids: move messages back, re-subscribe
   with the sender (an unsubscribe cannot be undone from this tool), recreate any
   channel this server created and deleted, and tell attendees if invitations went
   out.
6. **Report** privately (for example a GitHub security advisory) rather than in a
   public issue if the leak involved a token or another person's mail. Include the
   audit window, not message contents.

Prevention: keep `TEAMS_MCP_READ_ONLY=1` for every eval, demo and open-ended
prompt; grant writes per task; leave unsubscribe off unless the task is explicitly
about subscriptions; read Teams rather than posting to existing channels; prefer a
dedicated low-privilege account where possible.
