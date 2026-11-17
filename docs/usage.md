# Usage examples

Agent-facing calls for the main workflows. Arguments are JSON passed to the
tool's `arguments`; every datetimes is ISO 8601 (UTC unless a bare local time is
given, which is read in the configured timezone). Outputs shown are abbreviated.

## 1. Am I signed in?

```json
{"tool": "auth", "arguments": {"op": "status"}}
```

```json
{"authenticated": true, "fresh_audiences": ["outlook"],
 "account": "user@example.com", "source": "refresh",
 "audiences": [{"audience": "outlook", "expires_at": "2026-01-15T22:44:13.501Z", "fresh": true,
                "seconds_left": 3860, "scope_count": 74}],
 "cache_format": "v1"}
```

`authenticated` means **some** cached audience is fresh, not that every one is:
`fresh_audiences` lists the usable ones and each row in `audiences` reports its
own `fresh`. `reason` is only emitted when none is fresh ("run auth.login …"). So
an expired Outlook token can coexist with a working session:

```json
{"authenticated": true, "fresh_audiences": ["graph"],
 "audiences": [{"audience": "outlook", "fresh": false, "seconds_left": -120},
               {"audience": "graph", "fresh": true, "seconds_left": 4961}]}
```

Mail/calendar would then need `auth.login`, while directory search and
chat/channel reads and sends (which use the `graph` and Teams-web tokens) keep
working. The Teams-web cache is a separate file, so it appears in
`fresh_audiences` as `"teams_web"` whenever it is fresh, and the payload always
carries its own `teams_web` row (freshness, source, expiry, path). Even with no
`token.json` at all, a fresh Teams-web cache makes `authenticated: true`:

```json
{"authenticated": true, "fresh_audiences": ["teams_web"],
 "teams_web": {"authenticated": true, "fresh": true, "source": "browser", "refreshable": false}}
```

Check the per-audience rows before concluding the account is signed out. `status`
stays offline: both caches are read from disk, never refreshed.

## 2. Find a 30-minute slot with a colleague

```json
{"tool": "availability", "arguments": {
  "op": "suggest", "attendees": ["joel"], "duration_min": 30,
  "start": "2026-11-17T09:00:00", "end": "2026-11-21T17:00:00",
  "working_hours_only": true, "max_candidates": 5}}
```

```json
{"source": "findmeetingtimes", "duration_min": 30,
 "candidates": [{"start": "2026-11-17T16:00:00.000Z",
                 "local": "2026-11-17 10:00 → 2026-11-17 10:30",
                 "free": "2/2", "verified": true, "data_available": true,
                 "source": "findmeetingtimes"}]}
```

Names and aliases from `config.yaml` are resolved automatically; pass addresses if
you prefer. Busy data comes from a cached snapshot and is refreshed only when
stale (own calendar 5 min, others 30 min). Use
`{"op": "freebusy", ...}` for the raw busy blocks and the digit
`availability_view`; `busy` only holds real blocks (`free` intervals are omitted —
an empty list is the free case). Each `freebusy` row carries `data_available`:
false means the service returned no schedule for that mailbox, which is not the
same as free.

An address the service cannot resolve is reported, not treated as free:
`freebusy` adds `unknown: ["typo@…"]` plus `unknown: true` on that row (and no
digits for it), `suggest` excludes it with `excluded_unknown_attendees` and a
`warning`, and refuses with `unknown_attendees` when nothing resolves. Check the
address with `people.search` rather than reading an empty list as availability.
When some attendee data is missing, the result keeps going but says so:
`verified`/`data_available` are false, `data_warning` names the mailboxes, and a
service candidate with no per-attendee data reports `free: "unknown"` instead of
a fabricated `n/n`.

### One exact slot ("can we meet at 3?")

`suggest` ranks candidates; `check` answers a slot you already have in mind:

```json
{"tool": "availability", "arguments": {
  "op": "check", "attendees": ["joel"],
  "start": "2026-11-18T15:00:00", "end": "2026-11-18T15:30:00"}}
```

```json
{"available": false,
 "range": {"start": "2026-11-18 15:00", "end": "2026-11-18 15:30", "timezone": "America/Chicago"},
 "attendees": [{"email": "joel@example.com", "status": "busy", "data_available": true, "unknown": false}],
 "conflicts": [{"email": "joel@example.com", "status": "busy",
                "local": "2026-11-18 14:45 → 2026-11-18 15:15"}],
 "conflict_count": 1, "unknown": []}
```

`start` and `end` are required; omit `attendees` to ask about your own calendar.
`available` is true only when every attendee is free, and an unresolved mailbox
makes it false rather than guessing. This is a read: no write token, allowed in
read-only mode.

### Every open window

```json
{"tool": "availability", "arguments": {
  "op": "free", "attendees": ["joel"], "start": "2026-11-18", "end": "2026-11-21",
  "duration_min": 30, "working_hours_only": true, "max_slots": 10}}
```

```json
{"slots": [{"start": "2026-11-18T16:00:00.000Z",
            "local": "2026-11-18 10:00 → 2026-11-18 10:30"}],
 "count": 1, "truncated": false, "data_available": true, "unknown": []}
```

`free` lists windows from snapshot busy gaps (no `findmeetingtimes`), honours
`buffer_min`, and caps at `max_slots` (≤20). Use it for "what's open this week",
`check` for one slot, and `suggest` when you want a ranked pick.

## 3. Book the chosen slot

```json
{"tool": "calendar", "arguments": {
  "op": "create", "subject": "Planning sync",
  "start": "2026-11-17T16:00:00Z", "end": "2026-11-17T16:30:00Z",
  "attendees": ["joel"], "location": "Teams", "show_as": "busy"}}
```

```json
{"ok": true, "op": "create", "id": "AAMkAD...", "subject": "Planning sync",
 "web_link": "https://outlook.office365.com/owa/?itemid=...",
 "conflict_policy": "warn", "conflicts": []}
```

`IsReminderSet` is never sent (the v2 entity rejects it). Attendees outside
`invite_allowlist` are refused before the request. When `end` is absent,
`duration_min` supplies it — from `start`, or from the default start (next
15-minute boundary) when `start` is omitted too:

```json
{"tool": "calendar", "arguments": {
  "op": "create", "subject": "Planning sync",
  "start": "2026-11-17T16:00:00Z", "duration_min": 30}}
```

Passing `duration_min` together with both `start` and `end` is
`invalid_argument` (pass `start`+`duration_min`, or `start`+`end`).

`conflict_policy` controls what happens when the proposed time overlaps your own
calendar (events marked `free` and, on update, the event itself are ignored):

```json
{"tool": "calendar", "arguments": {
  "op": "create", "subject": "Planning sync",
  "start": "2026-11-17T16:00:00Z", "end": "2026-11-17T16:30:00Z",
  "conflict_policy": "refuse"}}
```

- `warn` (default) books anyway and reports `conflicts` (up to five),
  `conflict_count` and `conflict_warning` when there is an overlap.
- `refuse` returns `error: "conflict_detected"` with the same rows and writes
  nothing; pick another time or re-call with `warn`.
- `ignore` skips the calendar read entirely.
- If the read fails under `warn`, the booking still happens and the result says
  `conflict_check: "unavailable"` — never silently "no conflicts".

Moving an event uses exactly one time form (`calendar.update` also needs
`event_id` and `expected_subject`):

```json
{"tool": "calendar", "arguments": {"op": "update", "event_id": "AAMk...",
 "expected_subject": "Planning sync", "start": "2026-11-18T16:00:00Z"}}
```

```json
{"ok": true, "op": "update", "changed": {
  "start_before": "2026-11-17T16:00:00.000Z", "start_after": "2026-11-18T16:00:00.000Z",
  "duration_before_min": 30, "duration_after_min": 30}}
```

`start` alone keeps the existing duration; `end` alone keeps the start;
`duration_min` derives the missing bound; `shift_min` moves both bounds and keeps
the duration. Combining `shift_min` with a bound, or `duration_min` with both
bounds, is `invalid_argument`. `changed` reports before/after starts and
durations, and a time change runs the same conflict check.

## 4. What is on my calendar?

For "what's my day", use the agenda view — one compact row per event, no
attendees, location or organizer:

```json
{"tool": "calendar", "arguments": {"op": "agenda"}}
```

```json
{"range": {"start": "2026-11-17 00:00", "end": "2026-11-18 00:00", "timezone": "America/Chicago"},
 "count": 3,
 "events": [{"date": "2026-11-17", "start_local": "09:00", "end_local": "09:15",
             "time": "09:00–09:15", "subject": "Standup", "id": "AAMk..."}],
 "truncated": false}
```

The default window is the whole local day (`start` today, next local midnight);
pass `start`/`end` for another day or a multi-day range. `include_ids: false`
drops the id; `top` (default 25, cap 50) plus `truncated`/`next_cursor` page a
long range. It is a read, allowed in read-only mode.

For a range with more detail (show-as, location, web link opt-in) use `list`:

```json
{"tool": "calendar", "arguments": {"op": "list", "start": "2026-11-17", "end": "2026-11-18", "brief": true}}
```

```json
{"count": 3, "events": [{"id": "AAMk...", "subject": "Standup",
                         "local": "2026-11-17 09:00 → 2026-11-17 09:15"}]}
```

`brief: true` keeps `id, subject, start, end, local, show_as` and drops
location/organizer/attendees; `include_web_link: true` adds links. Default range
is today..+7d, cap 50.

## 5. Triage unread mail

```json
{"tool": "mail", "arguments": {"op": "list", "folder": "inbox", "unread_only": true, "top": 10, "preview_chars": 120}}
```

```json
{"messages": [{"id": "AAMk...", "subject": "Q4 numbers", "from": "cfo@example.com",
               "received": "2026-11-16T14:02:11.000Z",
               "received_local": "2026-11-16 08:02", "unread": true,
               "preview": "Please review the attached..."}],
 "next_cursor": "eyJ1Ijoi..."}
```

Every row carries `received_local` (configured timezone), so sorting or asking
"is this today?" needs no conversion. `brief: true` keeps only
`id, subject, from, received, received_local, unread` (+`has_attachments` when
true) and drops the preview; without it, `preview_chars` caps the preview at
40–1000 chars (default `max_results.preview_chars`, 300), and
`include_body_preview: false` removes the preview entirely. `top` accepts up to
50, but the effective page is capped by `max_results.mail` (default 25) — the
config value is the operator's opt-in.

Pass `cursor` back to page. `include_unsubscribe: true` adds
`unsub: {has, one_click}` so bulk senders are identifiable without opening them.
`from=<address>` filters by sender; that page is ordered locally (Exchange
rejects `$orderby` together with a `From` filter). Because a small `$top` there
was observed returning the oldest matches, the scan reads whole pages (50 each,
at most five), sorts them newest-first and returns the newest `top`. It reports
`scanned_pages`; when five pages were not enough to reach the end of the match
set it also returns `truncated: true` and a resumable `next_cursor` — pass it
back to continue scanning, not to "page 2" of a single ordered list.

Folder names and counts come from one read-only call:

```json
{"tool": "mail", "arguments": {"op": "folders"}}
```

```json
{"count": 6, "truncated": false, "folders": [
  {"id": "AAMk...", "name": "Inbox", "total": 4312, "unread": 18, "child_folders": 0}]}
```

`truncated: true` means the page of 50 was full, so more top-level folders may
exist. Use a folder's `id` as `folder` for `mail.list`.

Moves are deliberately small: `mail.move` takes one `message_id` and one
`destination` per call and takes a token from the one global write rate limit
(default 30/minute, shared by every mutating op), so bulk triage cannot happen in
one step; in read-only mode it is refused. There is no bulk move and no mail
deletion op.

Every capped read says so: list/search, `thread`, `calendar.list`,
`calendar.instances` and `people.search` return `count`, `truncated` when more
exist, and a `next_cursor` where the API provides one. A page is not the whole
answer.

A body read keeps its line structure (`body` is cut, not flattened); HTML mail is
converted to text with entities decoded. A bare `start`/`end` of `YYYY-MM-DD`
means local midnight; a bare wall clock is read in the configured zone, and a
time that the zone skips at a DST change is refused by name instead of being
moved an hour.

## 6. Read one message, then reply

```json
{"tool": "mail", "arguments": {"op": "get", "message_id": "AAMk...", "body_format": "text", "body_chars": 2000}}
```

```json
{"id": "AAMk...", "subject": "Q4 numbers", "body": "Please review ...", "body_truncated": false}
```

To read the whole conversation without fetching each message, use `thread`
(rows are compact by default; `include_body: true` adds text):

```json
{"tool": "mail", "arguments": {"op": "thread", "thread_id": "AAQk...", "top": 5, "include_body": true, "body_chars": 800}}
```

```json
{"thread_id": "AAQk...", "count": 2, "truncated": false,
 "messages": [{"id": "AAMk...", "subject": "Q4 numbers", "from": "cfo@example.com",
               "received_local": "2026-11-16 08:02", "unread": true,
               "body": "Please review the attached numbers ...", "body_truncated": false}]}
```

Thread bodies default to 500 chars (`body_chars` max 2 000) with
`body_truncated`; HTML is converted to text before the cut. Pass `cursor` to
continue a long conversation.

### Attachments

An attachment-only message (`has_attachments: true`, empty preview/body) is read
in two steps. First the metadata, without downloading content:

```json
{"tool": "mail", "arguments": {"op": "attachments", "message_id": "AAMk..."}}
```

```json
{"message_id": "AAMk...", "count": 1, "truncated": false,
 "attachments": [{"id": "AAMk...", "name": "Q4.pdf", "content_type": "application/pdf",
                  "size": 27175, "is_inline": false}]}
```

The list is capped at 20 rows (`truncated: true` when more exist) and never
returns content. Then read one attachment by its `id`:

```json
{"tool": "mail", "arguments": {"op": "get_attachment", "message_id": "AAMk...", "attachment_id": "AAMk..."}}
```

Text (`text/*`, `application/calendar*`) comes back inline, UTF-8, cut at 20 000
characters with `text_truncated`:

```json
{"id": "AAMk...", "name": "invite.ics", "content_type": "text/calendar",
 "size": 812, "text": "BEGIN:VCALENDAR...", "text_truncated": false}
```

Any other type is saved under `~/.cache/teams-mcp/attachments/` (directory 0700,
file 0600, sanitised basename plus a random suffix) and reported as a path; the
server does not open or execute it:

```json
{"id": "AAMk...", "name": "Q4.pdf", "content_type": "application/pdf",
 "size": 27175, "saved": true,
 "path": "/home/you/.cache/teams-mcp/attachments/…-Q4-a1b2c3d4.pdf"}
```

Limits and refusals: decoded content over 10 MB returns `note: "too_large"`
without saving; a reference attachment with no bytes returns
`note: "content_unavailable"`. Pass `text: true` to force UTF-8 decoding instead
of saving (useful for JSON or `.eml`, which are text but not `text/*`). Both ops
are read-only: they work in read-only mode and consume no write token.

```json
{"tool": "mail", "arguments": {"op": "reply", "message_id": "AAMk...", "body": "Reviewed — approved."}}
```

`send` takes explicit `to[]`, `subject`, `body` and returns
`{ok, op, recipient_count}` plus a `recipients: {to?, cc?, bcc?}` echo of the
resolved addresses (empty lists omitted), and an `id` when exactly one Sent
Items copy can be matched, otherwise a `note` (Outlook answers 202 with no id, so
nothing is guessed). `forward` needs
`message_id` and `to[]`. `reply`/`forward`/`send` look the sent copy up in Sent
Items
and add `id` when exactly one copy matches; when the copy is not (yet)
identifiable they return a `note` instead — read `sentitems` if you need to act
on it. Add `"reply_all": true` to `reply` to reach everyone on the original;
those recipients are checked against `invite_allowlist` before the send.
`calendar.create`/`update` return `attendees_detail` (email, name, type,
status), the same array `calendar.get` returns.

A safer two-step send — save the message first, review, then send it:

```json
{"tool": "mail", "arguments": {"op": "draft", "to": ["joel@example.com"],
 "subject": "Q4 numbers", "body": "Attached is the summary."}}
```

```json
{"ok": true, "op": "draft", "id": "AAMk...", "web_link": "https://outlook.office365.com/owa/...",
 "note": "saved to Drafts; nothing was sent"}
```

`draft` is a write (it creates an item in Drafts) and is refused in read-only
mode, but it never sends. Give the saved `id` to `send_draft` together with
`confirm: true`:

```json
{"tool": "mail", "arguments": {"op": "send_draft", "draft_id": "AAMk...", "confirm": true}}
```

`send_draft` re-reads the live draft, so a draft edited in Outlook is retested:
it refuses when the draft has no recipients or any recipient is outside
`invite_allowlist` (nothing sent), then sends and reports the Sent Items `id` or
a `note`. Its `recipients` echo therefore lists the draft's actual addresses,
not the ones supplied when the draft was created.

## 7. Preview an unsubscribe before doing anything

```json
{"tool": "mail", "arguments": {"op": "unsubscribe_preview", "message_id": "AAMk..."}}
```

```json
{"message_id": "AAMk...", "unsubscribe": {
  "url": "https://lists.example.com/u/abc", "mailto": "mailto:leave@example.com", "oneClick": true}}
```

Applying is deliberately separate and **disabled by default**:

```json
{"op": "unsubscribe_apply", "message_id": "AAMk...",
 "confirm": true, "expected_sender": "news@example.com"}
```

- If `safety.unsubscribe.enabled` is false (the default) the call returns
  `unsubscribe_disabled` and makes **no request**. An operator enables it in
  `config.yaml` or with `TEAMS_MCP_UNSUBSCRIBE_ENABLED=1`.
- When enabled, without `confirm: true` it returns `{refused: true, preview}` and
  sends nothing. `expected_sender` is required and must match the message's From
  address (case-insensitive); a mismatch refuses with the preview, so a stale id
  cannot unsubscribe a different sender.
- A daily cap (default 5) on **external actions**, not apply calls: an https
  one-click POST and a `mailto` send each consume a slot, so one dual-target
  message costs two. Exceeding the remainder returns `cap_reached` with the next
  local-midnight reset time and sends nothing; the cap is checked before the
  message is fetched.
- `unsubscribe_apply` takes a token from the global write rate limit (default
  30/minute, shared with every other mutating op); a `rate_limited` refusal
  includes `retry in Ns` and makes no request.
- At most that message's own https one-click POST and/or `mailto` send is
  performed — never in bulk. The POST target comes from the message, so it is
  vetted first: https only, no embedded credentials, no redirects followed, and
  the host must resolve to a public address. A refusal is reported as
  `{status: "refused", refused_reason}` with nothing sent. Every attempt is
  journaled in `~/.cache/teams-mcp/write-audit.jsonl`.

## Read-only mode for autonomous agents

```bash
TEAMS_MCP_READ_ONLY=1 node dist/index.js
# or config.yaml: safety: {read_only: true}
```

Every write op then returns `read_only` before any network call: mail
send/draft/send_draft/reply/forward/mark/move/unsubscribe_apply, calendar
create/update/delete/respond, `settings.set` and `auth.logout`. Reads,
`availability` (including `check` and `free`), `calendar.agenda`, `auth.status`
and `auth.login` still work; that includes
`mail.attachments` and `mail.get_attachment`, which only write downloaded bytes
into the local attachment cache. `calendar.agenda`, `availability.check` and
`availability.free` are the supported read-only ways to answer "what's my day"
and "is this slot free". The same environment
also honours `TEAMS_MCP_UNSUBSCRIBE_ENABLED`,
`TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP` and
`TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE`. Recommended for evals and any agent
that was not given an explicit, narrow write task.

Beyond read-only mode, the safety keys themselves are operator-owned: the MCP
`settings.set` refuses `safety.*`, `teams.allowed_targets` and
`invite_allowlist` (returns `settings_invalid`, journals
`safety_key_protected`), so an agent cannot disable read-only, enable
unsubscribe, raise a cap or widen the Teams send allowlist even in write mode. Use `config.yaml` or the shell CLI for those. The write rate limit is
one global token bucket (default 30/minute) covering every mutating op, and the
audit journal also records `subject_mismatch`, allowlist refusals and
`auth.login` attempts.

## 8. Resolve a person before inviting

```json
{"tool": "people", "arguments": {"op": "search", "query": "joel"}}
```

```json
{"ambiguous": true, "needs_disambiguation": "multiple directory matches; choose one by id or email before acting",
 "people": [{"id": "id-1", "name": "Joel Example", "email": "joel@example.com", "title": "Student"}]}
```

Pick an `id`/`email` and use it in `attendees`. Configured aliases
(`aliases: {joel: joel@example.com}` in `~/.config/teams-mcp/config.yaml`)
resolve exactly and skip search.

## Editing settings

```json
{"tool": "settings", "arguments": {"op": "set", "key": "work_hours.start", "value": "08:30"}}
{"tool": "settings", "arguments": {"op": "get"}}
```

Same operations from a shell: `node dist/cli.js settings get|set|validate|reload|path`.

`safety.*`, `teams.allowed_targets` and `invite_allowlist` are operator-owned:
the MCP `settings.set` refuses them with `settings_invalid` (audited as
`safety_key_protected`), while the shell CLI accepts them. Change those from a shell or by editing
`~/.config/teams-mcp/config.yaml`.

## Teams: chats and channels

Two credentials back one tool. Chat reads/writes use the Outlook Web Graph token;
channel history uses the Teams-web token that this server harvests from the
browser profile once and then refreshes headlessly (see [`docs/auth.md`](auth.md)).

Start with the offline status (no network):

```json
{"tool": "teams", "arguments": {"op": "status"}}
```

```json
{"teams_web": {"authenticated": true, "fresh": true, "source": "browser",
  "refreshable": true, "scope_count": 37, "seconds_left": 3861}}
```

`authenticated` means a cache exists; `fresh` (plus `seconds_left`) is the
usability flag, so a stale cache reports `authenticated: true, fresh: false` and
the next channel operation refreshes or re-harvests. `status` itself makes no
network call.

List chats, teams, and one team's channels. `chat_type` filters server-side where
supported, re-checks each row's `chatType` client-side, and scans every page, so
it is a complete answer; `query` matches the topic or a member display name.
`kind=teams` and `kind=channels` accept the same `query` over their names
(case-insensitive, client-side, across pages), which resolves "Robotics" or
"General 2025" without dumping every team/channel:

```json
{"tool": "teams", "arguments": {"op": "list", "kind": "chats", "top": 10}}
{"tool": "teams", "arguments": {"op": "list", "kind": "chats", "chat_type": "group", "top": 25}}
{"tool": "teams", "arguments": {"op": "list", "kind": "chats", "query": "joel"}}
{"tool": "teams", "arguments": {"op": "list", "kind": "teams", "query": "Robotics"}}
{"tool": "teams", "arguments": {"op": "list", "kind": "channels", "team_id": "<team-id>", "query": "General 2025"}}
```

Chat rows carry `last_updated`, plus `last_message_preview` and
`last_message_at` (both omitted when Graph exposes no preview) so same-titled chats are
distinguishable; duplicate ids from a mid-scan membership change are collapsed
to the newest snapshot.

Every answer carries `count`, `has_more` and `truncated`. `has_more` means a
`next_cursor` is available; `truncated` means the call stopped before filling the
requested `top` (page/scan budget), not merely that older rows exist — a page
that returned all `top` rows is not truncated.

Read messages — a chat, or a channel with its team id:

```json
{"tool": "teams", "arguments": {"op": "read", "chat_id": "19:…@thread.v2", "top": 20}}
{"tool": "teams", "arguments": {"op": "read", "team_id": "<team-id>", "channel_id": "19:…@thread.tacv2"}}
```

`top` is "exactly up to this many real messages", not "the newest N": roster
events are filtered before the cap, so ask for what you need.

Rows are compact and text-first: `{id, from, created, text, web_url, reply_to?,
message_type?, has_attachments?}`. HTML becomes plain text, cut at 500 chars by
default (`text_chars`, max 1 000); `replies` keeps a 1 000-char cut. `top`
defaults to 10 (cap 25) on `read`. Roster/system events are filtered out by
default and paging continues until
`top` human messages or no cursor, so "latest 10 messages" means ten messages.
Add `include_system: true` to see them rendered instead
(`{system: true, event: "member added", text: "member added: Name"}`). Continue
a page with `next_cursor`:

```json
{"tool": "teams", "arguments": {"op": "read", "team_id": "<team-id>",
 "channel_id": "19:…@thread.tacv2", "cursor": "<next_cursor>"}}
```

A cursor is bound to the call that minted it — op, surface, ids and filters
(`include_system`, `chat_type`, `query`). Replaying it with different arguments,
or using a cursor from an older build, answers `invalid_cursor` without a
request; repeat the first page with the new arguments instead.

Channel threads have their own op (`message_id` is the parent message):

```json
{"tool": "teams", "arguments": {"op": "replies", "team_id": "<team-id>",
 "channel_id": "19:…@thread.tacv2", "message_id": "1730000000000"}}
```

Sending is deliberately two-step. `text` is limited to 20 000 characters and an
oversized body is `invalid_argument` before any request or rate token. A target
must be one this server created (recorded in `teams-state.json`) **or** an
operator-listed one in `teams.allowed_targets` (`chat:<chat_id>` or
`channel:<team_id>/<channel_id>`). There is no model-facing override: anything
else is `target_not_allowed` and no request is made. An operator adds the exact
target to `config.yaml` (or `TEAMS_MCP_TEAMS_ALLOWED_TARGETS`); server-created
targets need no entry:

```json
{"tool": "teams", "arguments": {"op": "send", "chat_id": "19:…@thread.v2",
 "text": "Running 5 minutes late", "confirm": true}}
```

```json
{"ok": true, "surface": "chat", "target_id": "19:…@thread.v2", "message_id": "1730000000123"}
```

Create a private channel in a team you own (the server checks
`/groups/{team-id}/owners` and never creates a team):

```json
{"tool": "teams", "arguments": {"op": "create", "kind": "channel", "team_id": "<team-id>",
 "name": "AI TEST — teams-mcp", "private": true, "confirm": true}}
```

`teams.create kind=chat` is **not supported**: Graph accepts only `owner` or
`guest` for a chat member (`roles: []`, omitted roles and `"member"` all answer
400), so a created chat would silently make every invitee an owner. Create the
group chat in Microsoft Teams, ask the operator to add it to
`teams.allowed_targets`, then post with `send`. For channels, the whole shape
is validated before any
write token is spent or any request is made: `team_id` and `name` are required,
optional `members` are capped at 20 (`invalid_argument`), a channel without
`private: true` is `private_required` (this server never creates team-wide
channels), and a team where you are not an owner is `not_owner`. Any refusal
after the rate token was taken refunds it, so a rejected `create` does not use
up the write budget.

Created ids are remembered in `~/.config/teams-mcp/teams-state.json`, so a later
`send` to that chat/channel works with `confirm: true` alone. To remove a channel
this server created, pass its exact name as a guard:

```json
{"tool": "teams", "arguments": {"op": "delete", "team_id": "<team-id>",
 "channel_id": "19:…@thread.tacv2", "expected_name": "AI TEST — teams-mcp", "confirm": true}}
```

A successful delete answers `{deleted: true, …}`. If the channel is already gone
(404/410 from the pre-read or the DELETE), the call answers
`{deleted: false, already_absent: true}`, prunes the ledger row and journals
`already_absent` once — it is not an error and is safe to retry. Deleting is
limited to channels in the ledger — existing channels and chats are refused with
`target_not_created`; a name mismatch is `name_mismatch` and deletes nothing. In
read-only mode `teams.send`, `teams.create` and `teams.delete` are refused before
any network call, and every write (including `confirmation_required`/
`target_not_created`/`invalid_argument`/`private_required`/`not_owner` refusals)
is rate-limited and journaled like mail/calendar writes.

## Input validation and limits

Bad input is refused before any provider request, and every refusal is a
one-line tool error (`isError: true`) naming the problem:

- `timezone` must be a valid IANA zone; `start`/`end` must be real calendar
  dates (2026-13-01 and 2026-02-30 are rejected, not normalised).
- `since`/`until` are parsed as ISO 8601 and sent as canonical UTC literals, so
  an OData fragment such as `2026-01-01 or IsRead eq true` is rejected.
- Each `to`/`cc`/`bcc`/`attendees` element must be a single address. Display
  names (`Name <addr>`) are reduced to the addr-spec; comma-separated lists,
  multiple `@`, and CR/LF are `invalid_recipient` (no write is issued).
- `top` is clamped to the configured cap; zero/negative/fractional values cannot
  reach the provider as `$top`.
- `calendar.update` combinations are checked first: `shift_min` cannot be
  combined with `start`/`end`/`duration_min`, and `duration_min` cannot be
  combined with both `start` and `end`; an unknown `conflict_policy` is
  `invalid_argument`. None of these consumes a rate token.
- Availability `start`..`end` is capped at 366 days; `people.resolve` accepts at
  most 50 ids/emails per call.
- Unknown or corrupt token caches produce `auth_required` with a hint; they
  never trigger an unexpected browser login.
