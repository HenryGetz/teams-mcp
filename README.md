# teams-mcp

**Unofficial project.** Not affiliated with, endorsed by, or supported by
Microsoft, University of Northwestern – St. Paul, or any employer. Automating
sign-in and internal APIs may violate Microsoft or tenant terms and can lead to
account lockout or other account action. Use it only with your own account and
read [`DISCLAIMER.md`](DISCLAIMER.md) before you do. MIT licensed.

Token-efficient MCP server for Outlook mail and calendar, scheduling across
people, directory lookup, and Microsoft Teams chats/channels/messages via
Microsoft Graph. Seven `op`-dispatched tools with
compact output and a local calendar snapshot.

**API split (the tested tenant):** Outlook REST v2.0
(`https://outlook.office.com/api/v2.0`) is the primary API for mail, calendar and
scheduling; Microsoft Graph is used for directory search (`/users`) and for
**Microsoft Teams** (chats, channels, messages, channel creation), because Graph
app assignment for mail/calendar is blocked here (403).

Status: all seven tools and every advertised op are implemented. Mail and calendar
writes have explicit-argument and subject guards. Live probes covered
`send`, `get`, `mark`, `move`, `create`, `update`, `delete`, `list`, `search`,
`thread`, `freebusy`, `suggest` and `snapshot`. `reply`, `forward`, `respond`
and `unsubscribe_apply` were exercised end-to-end over MCP against a local mock
REST endpoint (never against the tenant). Teams was exercised live end to end:
`status`, `list` (teams/chats/channels), `read`, `replies`, one confirmed send,
and create → send → delete of a private channel (see [`docs/design.md`](docs/design.md)
§12). Login/refresh is implemented in `src/auth/*` and consumed through
`token.json`; the Teams-web token has its own cache (below).

## Safety modes

Start read-only. This README's registration snippets, `mcp.example.json` and
`config.example.yaml` all ship read-only, so a copy-paste agent cannot change a
real mailbox until you deliberately allow it. `unsubscribe_apply` is an
**irreversible external** action — a one-click opt-out is recorded by the third
party and cannot be undone — so it is off by default and capped when enabled.
Full threat model: [`SECURITY.md`](SECURITY.md).

| Control | Shipped example | Effect |
| --- | --- | --- |
| `safety.read_only` / `TEAMS_MCP_READ_ONLY=1` | **on** | Refuse **every** mutating op before any request: mail send/draft/send_draft/reply/forward/mark/move/`unsubscribe_apply`, calendar create/update/delete/respond, `teams.send`/`teams.create`/`teams.delete`, `settings.set`, `auth.logout`. Reads stay available, including `mail.attachments`/`mail.get_attachment`, `calendar.list`/`agenda`, `availability` (`freebusy`/`suggest`/`check`/`free`), the snapshot cache, `auth.status`, `auth.login` and Teams reads. |
| `safety.unsubscribe.enabled` / `TEAMS_MCP_UNSUBSCRIBE_ENABLED` | off | `unsubscribe_apply` is refused (no request) unless an operator opts in. Even then it needs `confirm: true` **and** `expected_sender` matching the message's From address, so a stale message id cannot unsubscribe a different sender. |
| `safety.unsubscribe.daily_cap` / `TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP` | 5 | Maximum **external unsubscribe actions** per local day, not apply calls: a message carrying both an https one-click URL and a `mailto` target consumes two slots. `cap_reached` reports the reset time. |
| `safety.write_rate_limit_per_minute` / `TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE` | 30 | **One global** per-process token bucket for every mutating op: mail send/draft/send_draft/reply/forward/mark/move/`unsubscribe_apply`, calendar create/update/delete/respond, `teams.send`/`teams.create`/`teams.delete`, `settings.set`, `auth.login`/`auth.logout`. A token is taken before the first request, so `rate_limited` (with `retry in Ns`) makes no network call at all. |

Teams adds refusals of its own, independent of read-only mode:
`teams.send`/`teams.create`/`teams.delete` need `confirm: true`. Posting into a
chat or channel this server did not create requires that exact target in the
operator allowlist `teams.allowed_targets` (`chat:<chat_id>` or
`channel:<team_id>/<channel_id>`); anything else is `target_not_allowed` before
any request, and there is no model-facing flag to widen it. `teams.delete`
additionally needs `expected_name` and only works on channels recorded in
`teams-state.json` (`name_mismatch` / `target_not_created` otherwise). Each
refusal happens before any write request.

Teams also validates the whole shape of a write before spending a rate token:
`teams.create` is **channel-only** and returns `invalid_argument` for a missing
or over-long argument list (`kind=channel` needs `team_id` and `name`, with at
most 20 optional `members`) and `private_required` for a channel without
`private: true`; a channel create in a team the signed-in user does not own
returns `not_owner` after one owner lookup. `kind=chat` is not accepted at all —
Graph requires every chat member to be an owner or guest, so this server never
creates a chat. Every Teams refusal is journaled, and
a refusal after the rate token was taken refunds it, so a rejected call cannot
exhaust the write budget. A delete whose channel is already gone (404/410 before
or during the request) answers `{deleted: false, already_absent: true}`, prunes
its `teams-state.json` row and writes exactly one `already_absent` journal line.

The model-facing `settings.set` refuses `safety.*`, `invite_allowlist` and
`teams.allowed_targets`: those keys are operator-owned, so an agent cannot enable
unsubscribe, raise a cap, switch off read-only or widen the Teams send allowlist.
Edit `config.yaml` or run `node dist/cli.js settings set …` from a shell instead.

Every write that reaches a tool's write path — allowed, refused by policy, or
blocked by read-only/rate-limit — is appended to
`~/.cache/teams-mcp/write-audit.jsonl` with a timestamp, tool/op, an id or
recipient count and an outcome; never message bodies or tokens. That includes
policy refusals and near misses: `subject_mismatch`, allowlist rejections
(`recipient_not_allowed`), protected-settings attempts (`safety_key_protected`),
`read_only`, `rate_limited`, `cap_reached`, mail `confirmation_required` /
`no_recipients` for `send_draft`, calendar `conflict_detected` /
`conflict_check_failed`, and every `auth.login` attempt.

**Remove the read-only setting only for a trusted write task** (drop
`TEAMS_MCP_READ_ONLY` from the env, or set `safety.read_only: false`), and put it
back afterwards. An agent cannot undo it on its own: `settings.set` is itself
blocked in read-only mode; an operator edits `config.yaml` or runs
`node dist/cli.js settings set safety.read_only false` from a shell.

Why these defaults exist: an early live-account test with an autonomous agent
performed bulk mailbox actions; the capability controls exist to prevent a
repeat. A read-only regression run with a cold agent was blocked before any
network call and wrote nothing, and a red-team review of this layer found and
fixed four gaps (rate limit scope,
unsubscribe cap accounting, operator-only `safety.*`, audit coverage). Design
rationale: [`docs/design.md`](docs/design.md) §10.

## Install

Requires Node >= 22 (Node 26 tested). The server uses `Promise.withResolvers`
and `AbortSignal.any`; both need Node 22+.

Run it straight from GitHub with `npx` — no clone, no committed `dist`:

```bash
npx -y github:HenryGetz/teams-mcp
```

That is how a stdio MCP client should launch it (`command: "npx"`, `args: ["-y",
"github:HenryGetz/teams-mcp"]`). The first run downloads the repo, runs
`npm install` and the `prepare` build (`tsc`), then starts the server on stdio;
expect **60–90 s** on a clean machine. Later launches still revalidate the git
ref and typically start in **10–40 s**; pin a commit (`github:HenryGetz/teams-mcp#<sha>`)
or install globally if you want a fixed, faster start. Nothing is published to
npm.

Some MCP clients abort a server that has not finished initializing within 60 s,
which that first install can exceed. Warm the cache once from a shell before you
register the server:

```bash
npx -y github:HenryGetz/teams-mcp </dev/null
```

It prints a `ready` line and exits on EOF, so client start-up only pays the
~10–40 s warm-path cost. A client with a stricter timeout should install globally
(`npm install -g github:HenryGetz/teams-mcp`) and use the `teams-mcp` bin.

The sign-in flow drives Chromium through Playwright, so install that browser
once on the machine that will run `auth login`:

```bash
npx playwright install chromium
```

Then sign in (`auth login` from your MCP client, or `node dist/cli.js auth
login`); an interactive login takes roughly **25–60 s**, and the token cache
makes later starts immediate.

Alternatives:

```bash
# global install, same git source
npm install -g github:HenryGetz/teams-mcp
teams-mcp

# from a checkout (development)
git clone https://github.com/HenryGetz/teams-mcp.git
cd teams-mcp
npm ci
npm run build
```

`npm pack` produces a tarball that installs the same way (`npm install -g
./teams-mcp-0.1.0.tgz`). `private: true` in `package.json` prevents an accidental
npm publish; install it from this repository, not the registry.

## Register with an MCP client

JSON (`mcp.json`, Claude Desktop, Cursor, VS Code):

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "github:HenryGetz/teams-mcp"],
      "env": {
        "TEAMS_MCP_TIMEZONE": "America/Chicago",
        "TEAMS_MCP_READ_ONLY": "1"
      }
    }
  }
}
```

`mcp.example.json` is the same file, ready to copy. `TEAMS_MCP_READ_ONLY=1` is
part of the shipped default: remove that line only for a trusted task that must
write (see [Safety modes](#safety-modes)). A source checkout can use
`"command": "node", "args": ["dist/index.js"], "cwd": "/absolute/path/to/teams-mcp"`
instead, after `npm run build`. To have a CLI write the client config for you:

```bash
npx -y add-mcp github:HenryGetz/teams-mcp -a cursor -n teams \
  --env TEAMS_MCP_TIMEZONE=America/Chicago --env TEAMS_MCP_READ_ONLY=1 -y
```

Swap `-a cursor` for your client's identifier; the env flags are the same
read-only defaults as above.

Installed from git or a tarball, the `teams-mcp` bin is the entrypoint — drop
`cwd` and use `"command": "teams-mcp"` with no args. With `npx`, use
`"command": "npx", "args": ["-y", "github:HenryGetz/teams-mcp"]`. Do not use
`npm install teams-mcp`: that registry name is not this project. The
repo-relative form above is for a source checkout after `npm run build`.

YAML (`opencode` style):

```yaml
mcp:
  teams:
    type: local
    command: ["node", "/absolute/path/to/teams-mcp/dist/index.js"]
    enabled: true
    environment:
      TEAMS_MCP_TIMEZONE: America/Chicago
      TEAMS_MCP_READ_ONLY: "1"
```

No secrets go in either file: credentials live in `~/.config/teams-mcp/token.json`
(0600), written by the login flow. Config-file equivalent: `safety.read_only: true`
in [`config.example.yaml`](config.example.yaml), which also ships read-only.
Environment variables win over the config file, so a stray
`TEAMS_MCP_READ_ONLY=1` keeps an agent read-only even after you edit the YAML.

## Tools

| Tool | Ops | Live |
| --- | --- | --- |
| `auth` | `status`, `login`, `logout` | all (login runs the real sign-in flow) |
| `mail` | `list`, `search`, `get`, `thread`, `folders`, `attachments`, `get_attachment`, `send`, `draft`, `send_draft`, `reply`, `forward`, `mark`, `move`, `unsubscribe_preview`, `unsubscribe_apply` | all implemented; send/list/search/get/thread/mark/move plus attachment metadata/content live-probed, reply/forward/draft/send_draft exercised against the mock REST endpoint |
| `calendar` | `list`, `agenda`, `get`, `create`, `update`, `delete`, `respond`, `instances` | all implemented; agenda/respond exercised against the mock REST endpoint |
| `availability` | `freebusy`, `suggest`, `snapshot`, `check`, `free` | all implemented; check/free covered by fixture tests over snapshot busy data |
| `people` | `search`, `resolve` | all |
| `settings` | `get`, `set`, `reload`, `validate`, `path` | all |
| `teams` | `status`, `list`, `read`, `replies`, `send`, `create` (channel-only), `delete` | all (list/read/replies/send/private-channel create/delete exercised live; created channel deleted, 404 on re-read; chat creation is refused by the schema) |

Write behaviour worth knowing:

- `mail.attachments` returns **metadata only** — up to 20 rows of
  `{id, name, content_type, size, is_inline}` plus `count` and `truncated` — for
  an attachment-only message. It never requests content.
- `mail.get_attachment` fetches one attachment by `attachment_id`. `text/*` and
  `application/calendar*` are decoded inline as UTF-8, cut at 20 000 chars with
  `text_truncated`; `text: true` forces that decode for other textual types
  (JSON, `.eml`). Every other type is saved to
  `~/.cache/teams-mcp/attachments/` (directory 0700, file 0600, sanitised
  basename plus a random suffix) and returned as `saved: true` with `path`.
  Decoded content over 10 MB returns `note: "too_large"` without saving; a
  reference attachment with no bytes returns `note: "content_unavailable"`. Both
  ops are reads: allowed in read-only mode, no write token, and saved files are
  never opened or executed.
- Mail write responses echo the resolved addresses as
  `recipients: {to?, cc?, bcc?}` (empty lists omitted) for `send`, `send_draft`,
  `reply`, `reply_all` and `forward`, so confirming where a message went needs
  no follow-up `mail.get`. For `send_draft` these are the draft's actual
  recipients, re-read immediately before sending.
- `mail.send` needs explicit `to[]`, `subject` and `body`; it answers
  `{ok, op:"send", recipient_count, recipients}` plus a best-effort `id` when exactly one
  copy of the message can be found in Sent Items (matching subject and a short
  receive window), otherwise a `note` instead — Outlook returns 202 with no id
  and an id is never guessed. Recipients outside `invite_allowlist` (when set)
  are refused.
- `mail.draft` takes the same `to[]`/`cc`/`bcc`/`subject`/`body` and saves the
  message to **Drafts**; it returns `{ok, id?, web_link?, note}` and sends
  nothing. It is a write on every safety gate (read-only refuses it, it takes a
  rate token, allowlist checks apply, the attempt is audited).
- `mail.send_draft` needs `draft_id` **and** `confirm: true`. It re-reads the
  live draft (a draft edited elsewhere may name new recipients), refuses when it
  has no recipients or any recipient is outside `invite_allowlist`, then sends
  it and reports the Sent Items `id`/`note` the same way `send` does. The rate
  token is refunded when the re-read or allowlist check refuses.
- `mail.reply` accepts `reply_all: true` (REST `/replyall`): the reply reaches
  the sender plus every To/Cc address, and every derived address is checked
  against `invite_allowlist` **before** the send, so a reply-all that would
  leave the allowlist is `recipient_not_allowed` with nothing sent.
- `mail.folders` lists top-level folders in one page (max 50) as
  `{id, name, total, unread, child_folders}` with `count` and `truncated`
  (true on a full page). Read-only.
- Recipient/attendee values without `@` are resolved through the tenant
  directory before any write: one match is used, several return
  `ambiguous_recipient` with every candidate and `needs_disambiguation`, none
  returns `recipient_not_found`. A name is never guessed. This applies to
  `mail.send` (to/cc/bcc), `mail.forward`, and `calendar.create`/`update`.
- Mail list/search rows carry `conversation_id` (usable as `mail.thread`'s
  `thread_id`). `list` with `from=<address>` cannot use `$orderby` with that
  filter (Exchange answers `InefficientFilter`) and a small `$top` was seen
  returning the *oldest* matches, so it scans whole pages (50 at a time, at most
  5), sorts the union newest-first client-side, returns the newest `top`, and
  reports `scanned_pages`, plus `truncated` and a resumable `next_cursor` when
  the 5-page cap stopped before the match set ran out. `web_link` is omitted by default (`include_web_link: true`
  restores it); `include_unsubscribe: true` adds `unsub{has, one_click}` so a
  newsletter can be found in one call, and `include_unsubscribe_targets: true`
  additionally inlines `url`/`mailto` (~1.2 KB over 25 rows; use
  `unsubscribe_preview` when you need the detail for one message).
  `mail.get` caps the body at 4 000 chars by default (`body_chars`, max 20 000;
  raise it when the body matters).
- Every list/search row also carries `received_local` — the received instant
  rendered in the configured timezone — so a caller can judge "today" without
  converting the UTC `received` value itself. `brief: true` keeps only
  `id, subject, from, received, received_local, unread` (plus `has_attachments`
  when true), dropping preview, flag, conversation id and links.
  `preview_chars` (40–1000, default `max_results.preview_chars`, shipped 300)
  caps the `preview` text; `include_body_preview: false` removes it entirely and
  wins over any `preview_chars`.
- `mail.list`/`search` declare `top` up to **50** in the schema, but the page is
  additionally capped by `max_results.mail` (default **25**, max 50). The config
  value is the operator's opt-in: an agent cannot exceed it per call, and
  `truncated`/`next_cursor` report when the cap cut the answer.
- `mail.thread` returns compact rows by default. `include_body: true` adds each
  message's body as text, cut at 500 chars (`body_chars`, max 2 000) with
  `body_truncated`; HTML bodies are entity-decoded and newline-converted before
  the cut. `top`/`cursor` page the conversation and `truncated` marks a full
  page.
- `mail.move` returns the message's **new** id.
- `calendar.update`/`delete`/`respond` require `event_id` **and**
  `expected_subject`; a mismatch returns `subject_mismatch` and performs no write.
  `calendar.update` also applies `invite_allowlist` to new attendees.
- `calendar.create` defaults: timezone from config, start at the next 15-minute
  boundary, duration `default_duration_min`; all-day events use local dates.
  `duration_min` supplies the end when `end` is absent (measured from `start`,
  or from the default start when both are omitted) and is rejected with
  `invalid_argument` when both `start` and `end` are given — pass
  `start`+`duration_min`, or `start`+`end`. A `start`/`end` without `Z`/offset is
  read in `timezone` (a time the zone skips at
  a DST change is refused by name). Write responses are converted back to UTC
  using the `TimeZone` echoed by the API, so `start`, `end` and `local` always
  describe the stored event. `create` and `update` return `attendees_detail`
  (`email, name, type, status`) exactly as `get` does.
- `calendar.create`/`update` accept `conflict_policy: warn|refuse|ignore`
  (default `warn`). The check reads the user's own `calendarview` for
  `[start, end)`, ignores events marked `free`, and (on update) the event being
  edited. `warn` books and reports `conflicts` (≤5 rows) plus `conflict_count`
  and `conflict_warning` when any overlap; `refuse` returns `conflict_detected`
  with `conflicts`/`conflict_count` and writes nothing; `ignore` skips the read.
  If the check itself failed under `warn` the write still happens and the result
  says `conflict_check: "unavailable"`. The read runs only after the read-only,
  rate-limit and allowlist guards, so a refused write issues no calendar read.
- `calendar.update` time fields: `start` alone moves the start and keeps the
  event's existing **duration**; `end` alone moves the end and keeps the start;
  `duration_min` derives the missing bound (with `start`, `end`, or neither);
  `shift_min` moves both bounds by a signed number of minutes keeping duration.
  `shift_min` cannot be combined with `start`/`end`/`duration_min`, and
  `duration_min` cannot be combined with both bounds — those are
  `invalid_argument` before any request. Every successful update returns
  `changed{start_before, start_after, duration_before_min, duration_after_min}`
  so the actual move is visible, and a time change triggers the conflict check
  above (a subject-only edit does not read the calendar).
- `calendar.list` omits `web_link` by default (`include_web_link: true`); `brief:
  true` keeps only `id, subject, start, end, local, show_as`.
- `calendar.agenda` is the "what's my day" view: one compact row per event over a
  day or range, defaulting to today in `timezone` (whole local day, DST-safe
  across midnight; pass `start`/`end` for another day or a multi-day range).
  Rows are `{date, start_local, end_local, time, subject}` — no attendees,
  location, organizer or web link — and `time` is `09:00–09:30` or `all day`.
  `include_ids: false` drops the event id; the default `true` keeps it so a row
  can be passed to `calendar.get`. `top` (default 25, cap 50) plus
  `truncated`/`next_cursor` page a long range. Read-only: no write token.
- `availability.check` answers one exact proposed slot: `start` and `end` are
  required (a missing bound is `invalid_argument` before any read), attendees
  default to the signed-in user, and the result is `available` (true only when
  every attendee is free), `range` with local labels, per-attendee rows
  `{email, status, data_available, unknown}`, overlapping `conflicts` (up to 5
  with the true `conflict_count`) and `unknown: [emails]`. An unknown mailbox
  makes the slot unverifiable — `available: false`, never a guess. It writes
  nothing and takes no write token, so it is the cheapest "can we meet at 3?"
  probe.
- `availability.free` lists explicit free windows rather than a ranked subset:
  `start`/`end` default to today..+7d, `duration_min` defaults to
  `default_duration_min`, `working_hours_only` is true by default, `buffer_min`
  pads around meetings, and `max_slots` (≤20) caps the answer. It is computed
  from snapshot busy gaps locally — no `findmeetingtimes` — and returns
  `slots[{start, end, local}]`, `count`, `truncated`, `data_available` and
  `unknown`. Use it when the question is "what windows are open", not "pick
  one". Both `check` and `free` are reads: available in read-only mode, no
  write token, no rate-limit token.
- `availability.suggest` returns up to `max_candidates` **ranked** slots inside
  the requested `start`/`end` only (never the cached snapshot's wider range),
  honours `working_hours_only` inside that window, and clamps `max_candidates`
  to 20. It is a ranked subset, not every free window: the output reports
  `returned`, `max_candidates`, `truncated`, and (local search) `total_available`.
  Those are counts over the attendees with schedule data. Each candidate carries
  `verified`/`data_available` — false (and `data_warning` set) when any requested
  attendee has no schedule data — and when `findmeetingtimes` returned a
  candidate with no per-attendee availability, `free` is the string `"unknown"`
  plus a `data_warning` rather than a fabricated `n/n`.
  Windows are capped at 366 days (local slot search is O(window/interval)).
  A request whose only attendee is the signed-in user falls back to the local
  snapshot search when the service returns candidates without per-attendee
  availability (`free: "unknown"`), so a self-only "when am I free" question
  still gets real slots instead of unverified ones.
  Use `freebusy` when you need the full busy view. Availability attendees accept
  names/aliases/addresses like mail and calendar, including for `freebusy` and
  `snapshot`.
- `busy` never contains `status:"free"` intervals (absence of a block means
  free; `availability_view` digits still encode free time). If the service
  cannot resolve an attendee's mailbox (a typo, a foreign address) it returns no
  working hours, which is **not** evidence of availability: `freebusy` reports
  `unknown: [emails]`, marks those rows `unknown: true` and omits their digits,
  while `suggest` excludes them with `excluded_unknown_attendees` + `warning`
  (or refuses when nobody resolves). Every `freebusy` attendee row also carries
  `data_available` (false means "the service returned nothing for this mailbox",
  not "free all day"), and the snapshot metadata records the mailboxes with data
  as an optional `data_available[]` list.
- `mail.reply`/`forward` look the sent copy up in Sent Items (same conversation,
  short window after the 202) and return its `id` only when exactly one copy
  matches; otherwise they return a `note` and the message id stays discoverable
  via `list folder=sentitems`.
- `unsubscribe_apply` does nothing without `confirm: true`; with it, it acts on
  at most the two targets of that one message (one https one-click POST and/or
  one `mailto` sendmail). No bulk unsubscribe, no mail deletion.
- `people.search` returns every match; multiple matches come back with
  `ambiguous: true` and `needs_disambiguation` instead of a silent pick.
- Input validation is strict at the boundary: each `to`/`cc`/`bcc`/`attendees`
  element must be one addr-spec (a comma list or CRLF is `invalid_recipient`),
  `since`/`until` are parsed and canonicalised before they reach the provider
  filter, an unknown IANA `timezone` or impossible date is `invalid_argument`,
  availability windows are capped at 366 days, and `people.resolve` accepts at
  most 50 inputs per call.

### Compact answers

Prefer the row-level views when the question is a summary, not a full record:
`calendar.agenda` for "what is my day", `calendar.list brief: true` for a
range, `availability.check` for one proposed slot and `availability.free` for
the open windows, and `mail brief: true`/`preview_chars` for triage. These
return fewer fields and shorter strings than the detail ops; the actual byte
saving depends on the mailbox and the window, so treat it as a habit rather than
a fixed promise. Reach for `get`/`freebusy`/`snapshot` when the details are
what the task needs.

Full contracts and rationale: [`docs/design.md`](docs/design.md).
Worked agent-facing examples (schedule with someone, triage/reply mail,
unsubscribe preview, person resolution): [`docs/usage.md`](docs/usage.md).
Login, token cache and refresh flow: [`docs/auth.md`](docs/auth.md).
Every environment variable: [`.env.example`](.env.example); every config key:
[`config.example.yaml`](config.example.yaml).

## Auth status

`auth.status` reads only the local token cache (offline, safe to poll).
`authenticated` is true when **any** cached audience has a fresh access token;
`fresh_audiences` names which (`["outlook"]`, `["graph"]`, `["teams_web"]`, or a
combination), and `reason` is present only when none is fresh. Each row in
`audiences` carries its own `fresh`, `seconds_left` and `scope_count`, and the
payload always includes a `teams_web` row (freshness, source, expiry, path) so
the Teams-web cache is visible even when it is stale.

So a stale `outlook` token does not make a working Graph/Teams-web session look
signed out: mail and calendar may need `auth.login` while directory search, chat
reads, and channel reads/sends keep working. Conversely
`fresh_audiences: ["outlook"]` means Graph-only operations (Teams via the OWA
`graph` token, directory search) will fail with `auth_required`.

## Microsoft Teams

Chats, team/channel metadata, channel messages and replies, plus three writes
(message send; private-channel create; deletion of a private channel this server
created). Chats themselves are **never created** by this server (see below).
Teams needs a **second** credential, selected per operation, never by user
choice:

- **OWA `graph` token** (the one `auth.login` maintains): chats and chat messages,
  team and channel metadata, channel create and delete, and chat sends.
- **Teams-web Graph token**: channel messages, replies, team/channel rosters, and
  channel sends. In the tested tenant the OWA token has no
  `ChannelMessage.Read.All`; the Teams web client
  (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`) mints a Graph token that does, so
  every channel-message read and every channel send uses it.

First use **silently harvests** the Teams web session from the persistent browser
profile — no password, no prompt, ~11–25 s (typically ~18–25 s warm) — then
persists it. After that, renewal is a headless refresh (typically 1–3 s) and no
browser opens. It is a separate cache:
`~/.config/teams-mcp/teams-web-token.json` (0600, own refresh lock).
`~/.config/teams-mcp/teams-state.json` (0600) is the ledger of chats and channels
this server created: it decides what may be written to and deleted.

Reads:

- `list kind=chats` accepts `chat_type=group|oneOnOne|meeting` and `query`
  (topic or member name), applied across pages, and returns `count`,
  `scanned_pages`, `truncated` and `next_cursor` — a filtered list is a complete
  answer, not one page. Duplicate ids from a mid-scan membership change are
  collapsed to the newest snapshot, and rows carry `last_updated` plus
  `last_message_preview`/`last_message_at` when Graph exposes a preview (omitted
  otherwise, so same-titled chats stay distinguishable).
- `read` hides roster/system events by default (`include_system: false`): they are
  filtered before the page is filled, so `top` counts real messages — `top` means
  "exactly up to N real messages", never "the latest N", and defaults to **10**
  (cap 25). Message text is cut at 500 chars by default (`text_chars`, max
  1 000); `replies` keeps its 1 000-char default. Set
  `include_system: true` to see them rendered as readable text. Paging reports
  `truncated` and `next_cursor`.

Writes:

- `teams.send` and `teams.create` require `confirm: true`.
- `teams.create` is **channel-only**. `kind=chat` is not in the tool schema and
  Graph would not accept it safely: every chat member must be `owner` or `guest`
  (`roles: []`, omitted roles and `"member"` all answer 400), so the only valid
  group chat grants every invitee ownership — and this app cannot delete it
  (no `Chat.ManageDeletion.All`). Create group chats in Microsoft Teams, ask the
  operator to add the chat to `teams.allowed_targets`, then `send` with
  `confirm: true`. Chats that already exist can be read and posted to when
  allowlisted.
- `teams.create kind=channel` needs `team_id`, `name` and `private: true`
  (`private_required` otherwise: team-wide channels are never created), plus at
  most **20** optional `members` (`invalid_argument` beyond that, before any
  lookup), and the signed-in user must be an owner of the team (`not_owner` after
  one owner lookup).
- Sending to a chat/channel this server did **not** create requires the exact
  target in `teams.allowed_targets` — `chat:<chat_id>` or
  `channel:<team_id>/<channel_id>` — set in `config.yaml`,
  `TEAMS_MCP_TEAMS_ALLOWED_TARGETS`, or via the shell CLI. Otherwise the call is
  `target_not_allowed` and issues no request. `settings.set` refuses the key, so
  an agent cannot widen its own send scope. Targets this server created are in
  `teams-state.json` and need no entry.
- Existing test chats still have both users as owners; Graph cannot change chat
  member roles here (no `ChatMember.ReadWrite*`; a member `PATCH` answers 405 on
  v1.0 / 404 on beta), so correct that in the Teams UI. The diagnostic chat used
  during testing was deliberately left labelled and empty.
- `teams.delete` is **channel-only** and narrower still: it needs `confirm: true`
  plus `expected_name` matching the live channel name (else `name_mismatch`), and
  the channel must be in `teams-state.json` — i.e. one this server created.
  Pre-existing channels and chats cannot be deleted here. It uses the OWA token,
  treats a 404/410 as already absent rather than an error
  (`already_absent: true`), and prunes the ledger entry either way.
- Read-only mode, the one global write rate limit and the audit journal apply
  exactly as they do to mail and calendar. Policy refusals are journaled too:
  `confirmation_required`, `target_not_allowed`, `target_not_created`,
  `name_mismatch`, `invalid_argument`, `private_required`, `not_owner` and
  `already_absent`. A refusal that already took a rate token refunds it (the
  `already_absent` prune included), and each delete attempt writes exactly one
  journal line.

Default posture is read-what-exists: use `status`, `list`, `read` and `replies`,
and do not post into an existing channel unless the task explicitly allows it.
No attachments, mentions, edits, bulk sends, chat deletion or deletion of
channels created outside this server.

## Configuration

User-editable YAML at `~/.config/teams-mcp/config.yaml`; see
[`config.example.yaml`](config.example.yaml). Precedence:
**environment variables > config.yaml > defaults**. The file is validated with
zod and re-read when it changes (no restart needed).

```bash
node dist/cli.js settings path
node dist/cli.js settings get
node dist/cli.js settings set work_hours.start 08:30
node dist/cli.js settings set default_duration_min 45
node dist/cli.js settings set aliases.joel joel@example.com
node dist/cli.js settings validate
```

Invalid values are refused before the file is written; writes are atomic (0600).
The same operations exist as the `settings` MCP tool.

Important keys: `timezone` (default `America/Chicago`), `work_hours`,
`default_duration_min`, `aliases`, `invite_allowlist`, `blocked_times`,
`cache.{own_calendar_ttl_minutes,busy_ttl_minutes,directory_ttl_minutes}`,
`max_results.*` (`mail` 1–50, default 25 — the per-call `top` cap an operator
can raise; `calendar` 1–50; `people` 1–25; `slots` 1–20; `preview_chars` 40–1000,
default 300), `snapshot_dir`, `test_event_prefix`, and the `safety` block
(`read_only`, `unsubscribe.{enabled,daily_cap}`, `write_rate_limit_per_minute`).
The example file ships `safety.read_only: true`; the built-in default is `false`,
so a hand-written config without a `safety` block is write-enabled.

Paths:

| Path | Purpose |
| --- | --- |
| `~/.config/teams-mcp/config.yaml` | user settings |
| `~/.config/teams-mcp/token.json` | dual-audience token cache (0600) |
| `~/.config/teams-mcp/teams-web-token.json` | Teams-web Graph credential, own refresh lock (0600) |
| `~/.config/teams-mcp/teams-state.json` | chats/channels this server created (0600) |
| `~/.config/teams-mcp/browser-profile/` | persistent Chromium profile for silent re-auth |
| `~/.cache/teams-mcp/snapshots/calendar.yaml` | calendar snapshot (0600, safe to delete) |
| `~/.cache/teams-mcp/attachments/` | files saved by `mail.get_attachment` (dir 0700, files 0600, safe to delete) |
| `~/.cache/teams-mcp/write-audit.jsonl` | append-only write journal (0600, rotates at 5 MB) |
| `~/.cache/teams-mcp/unsubscribe-state.json` | daily unsubscribe counter (0600) |

Environment overrides (see [`.env.example`](.env.example)):
`TEAMS_MCP_CONFIG_DIR`, `TEAMS_MCP_CONFIG_FILE`, `TEAMS_MCP_TIMEZONE`,
`TEAMS_MCP_PROFILE_DIR`, `TEAMS_MCP_CACHE_DIR`, `TEAMS_MCP_SNAPSHOT_DIR`,
`TEAMS_MCP_OUTLOOK_BASE_URL`, `TEAMS_MCP_GRAPH_BASE_URL`, request/TTL knobs,
`TEAMS_MCP_READ_ONLY`, `TEAMS_MCP_UNSUBSCRIBE_ENABLED`,
`TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP`, `TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE`.

## Calendar snapshots

`availability op=snapshot` (and the optional resource
`teams-mcp://calendar/snapshot`) exposes the full YAML view — metadata, own
events and per-attendee busy blocks; it can be ~10 KB and is intended for bulk
planning. For a single scheduling question, `availability.freebusy`/`suggest`
return compact summaries and are cheaper. Freshness: own calendar 5 minutes,
other people's busy 30 minutes (configurable). There is **no delta sync** in this
tenant (405), so each refresh is a bounded `calendarview` + `getSchedule` poll;
`freebusy`/`suggest` refresh only when the snapshot is stale or its range does not
cover the request. Concurrent refreshes coalesce per process, temp names are
unique per writer, and the file is replaced atomically — a reader never sees a
partial document and the previous snapshot survives a failed write.

## Development

The repo includes a GitHub Actions workflow (`.github/workflows/ci.yml`: Node 22,
`npm ci`, typecheck, test, build, no secrets). It is not running right now because
GitHub Actions is blocked at the account level by a billing issue, so no CI badge
is shown; the same steps pass locally on Node 22.

```bash
npm run dev          # tsx src/index.ts
npm run typecheck
npm test             # vitest, no network
npm run build
npm run smoke        # spawns dist/index.js, initializes, lists tools/resources
node scripts/dev/measure-surface.mjs  # bytes + bytes/4 estimates: tools/list, per-tool schema, 4 read probes
LIVE_WRITES=1 npm run live-writes   # one-time real probe: 1 event, 1 self-email
TEAMS_MCP_LIVE_TEST=1 npx vitest run test/live-read.test.ts  # read-only live controls
```

A few tests can read optional redacted fixture payloads from
`docs/recon-artifacts/`; that directory is not committed, so those tests skip.

`live-writes` drives the built server over stdio as an agent would: create →
update → get → guarded delete refusal → delete → list for one calendar event, one
email to yourself (search → mark unread → mark read → move to Archive), and a
read-only `unsubscribe_preview`. It reads `token.json` read-only and aborts if the
outlook token is stale rather than refreshing.

## Security notes

Full threat model, credential permissions, incident response and rotation
guidance: [`SECURITY.md`](SECURITY.md).

- Every call runs as the signed-in user; there is no impersonation.
- **Capability scoping:** `TEAMS_MCP_READ_ONLY=1` refuses every mutating op
  before any request; `unsubscribe_apply` is opt-in (disabled by default) and
  needs `confirm:true` plus a matching `expected_sender`; one global 30/minute
  token bucket covers every write. See [Safety modes](#safety-modes).
- Destructive ops need ids from a prior read; `calendar update/delete/respond`
  also require `expected_subject` and abort on mismatch.
- `unsubscribe_apply` acts on at most that message's own https/mailto targets.
  No bulk unsubscribe, no mail deletion, no mailbox rule changes.
- `mail.send`/`draft`/`forward`, `mail.send_draft` (re-checked against the live
  draft), `mail.reply` with `reply_all`, `calendar.create`/`update` and
  `availability` attendees
  resolve names through the directory, then reject recipients outside
  `invite_allowlist` when the list is non-empty; ambiguity is an error, never a
  pick.
- Every write attempt is journaled to `~/.cache/teams-mcp/write-audit.jsonl`
  (ids/counts and outcome only; no bodies, no tokens).
- `mail.attachments`/`mail.get_attachment` are reads (allowed in read-only mode,
  no write token). Downloaded files land under
  `~/.cache/teams-mcp/attachments/` (0700/0600) and are reported by path only —
  the server never opens or executes them.
- Tokens are never logged; API errors redact `Bearer …`.
- `token.json`, `teams-web-token.json`, `teams-state.json`, the browser profile,
  snapshots, the attachment cache, the audit journal and the unsubscribe counter
  are owner-only (0600/0700).
- Teams follows the same gates as mail/calendar: read-only refuses
  `teams.send`/`teams.create`/`teams.delete` before any request; send/create need
  `confirm:true`; create is channel-only (never a chat); posting to a
  chat/channel this server did not create requires the exact target in the
  operator-owned `teams.allowed_targets` (`target_not_allowed` otherwise); and
  delete is channel-only, needs `expected_name` (`name_mismatch`) and only
  removes channels recorded in `teams-state.json`. Channel-message reads use the
  Teams-web Graph token because that is the only credential here with
  `ChannelMessage.Read.All`; see [`docs/auth.md`](docs/auth.md).
