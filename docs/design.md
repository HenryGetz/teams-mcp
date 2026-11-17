# Design: token-efficient Teams/Outlook MCP server

Status: implemented against the real tenant. Outlook REST v2.0 is primary; Graph
is directory-only. Login/refresh lives in `src/auth/*` and writes `token.json`.
An independent QA pass (recorded in the development repository, not shipped)
covered the fixed write-response timezone bug.

## 1. API architecture (recon-determined, not a preference)

| Concern | API | Audience |
| --- | --- | --- |
| Mail | Outlook REST v2 `https://outlook.office.com/api/v2.0/me/messages…` | `outlook` |
| Calendar | Outlook REST v2 `/me/calendarview`, `/me/events` | `outlook` |
| Scheduling | REST v2 `POST /me/calendar/getschedule`, `POST /me/findmeetingtimes` | `outlook` |
| Directory | Graph `GET /users?$search="…"` + `ConsistencyLevel: eventual` | `graph` |

Graph app assignment for mail/calendar is blocked in the tested tenant (403), so no
Graph mail/calendar code path exists. Graph is imported only by
`src/tools/people.ts` via `src/graph/client.ts` + `searchDirectory()`.

REST v2 payloads are PascalCase (`Subject`, `ReceivedDateTime`, `Start.DateTime`,
`InternetMessageHeaders`) and wrapped in `value` with `@odata.*`. The boundary in
`src/outlook/client.ts` camelises keys, drops `@odata.*`, and maps rows onto
compact normalized objects (`normalizeMessage`, `normalizeEvent`). No PascalCase
or `@odata` envelope escapes that module; tools and snapshots see camelCase only.

There is **no delta support** (405 on `/calendar/delta`): calendar snapshots are
time-window `calendarview` polls, and pagination uses `@odata.nextLink`.

## 2. Tool surface: seven tools, `op` dispatch

| Tool | Ops | State |
| --- | --- | --- |
| `auth` | `status` \| `login` \| `logout` | all live (login runs the browser flow) |
| `mail` | `list` \| `search` \| `get` \| `thread` \| `folders` \| `attachments` \| `get_attachment` \| `send` \| `draft` \| `send_draft` \| `reply` \| `forward` \| `mark` \| `move` \| `unsubscribe_preview` \| `unsubscribe_apply` | reads + send/mark/move live-probed; reply/forward/draft/send_draft exercised against a mock REST endpoint; attachment metadata/content read live |
| `calendar` | `list` \| `agenda` \| `get` \| `create` \| `update` \| `delete` \| `respond` \| `instances` | all implemented; respond/instances/agenda also exercised against a mock REST endpoint |
| `availability` | `freebusy` \| `suggest` \| `snapshot` \| `check` \| `free` | all implemented; check/free exercised over snapshot busy fixtures (getSchedule + findmeetingtimes + local search) |
| `people` | `search` \| `resolve` | live (Graph directory) |
| `settings` | `get` \| `set` \| `reload` \| `validate` \| `path` | live |
| `teams` | `status` \| `list` \| `read` \| `replies` \| `send` \| `create` \| `delete` | live (chat + channel read/send/create/delete; see §3) |

Why `op` dispatch: every tool's name, description and JSON schema is prompt
payload on every request. One schema per domain with a short enum costs a
fraction of what a tool-per-operation surface costs, and unknown ops return the
valid set. As measured by `scripts/dev/measure-surface.mjs`, `tools/list` for
all seven tools is ~15.4 KB (~3.85k tokens at bytes/4) plus a 408-byte resource
listing; descriptions were shortened to keep that down. This is schema size, not
a promise about per-call runtime savings. Teams presence remains absent (403 in
the tested tenant); the Teams tool is
split across two tokens because no single token covers both chat and channel
history (see [`docs/auth.md`](auth.md)).

Rejected: one tool per op (schema cost grows with features), a generic Graph
proxy (no validation, no safety boundary), resources for every object (clients
pull resources only on demand and cannot act).

## 3. Tool contracts

Shared rules: ISO 8601 UTC on the wire, local display strings alongside; ids come
from a prior read; compact JSON with short keys and no nulls; one-line errors;
opaque `next_cursor`; caps enforced server-side.

### auth

- `status` — offline read of `token.json`: `{authenticated, expires_at, source,
  account, audiences:[{audience,expires_at,fresh}], path, cache_format}` or
  `{authenticated:false, reason}`. `authenticated` reflects the outlook token
  (primary API) with a 5-minute expiry margin.
- `login` — runs the real sign-in flow (`forceLogin`) with the persistent
  browser profile; audience `outlook` by default, `graph` on request. Writes
  `token.json` (0600) before returning.
- `logout` — deletes `token.json`; browser profile and legacy captures stay.

### mail

- `list` — folder (default `inbox`), `top` ≤ 50 in the schema but capped by
  `max_results.mail` (default 25; the operator opts in to more), `unread_only`,
  `from`,
  `since`/`until`, `cursor`; `$select` limited to the row fields.
  `from=<address>` is special: Exchange rejects `$orderby` with a
  `From/EmailAddress/Address` filter (`InefficientFilter`) and a small `$top`
  without ordering was observed returning the *tail* of the match set, so
  `from` scans up to 5 pages of 50, sorts the union by `ReceivedDateTime`
  newest-first **client-side**, then slices the caller's `top`. The result is a
  newest-N answer, with `scanned_pages`; when the 5-page cap stopped before the
  match set ended it also returns `truncated: true` and a resumable
  `next_cursor` (the underlying nextLink) so the scan can continue. Ordinary
  folder pages keep the existing `$orderby ReceivedDateTime desc` path.
- `search` — REST `$search` (cannot be combined with `$filter`/`$orderby`).
- `get` — one message with `Body` and `InternetMessageHeaders`;
  `body_format=text` converts HTML to text (`htmlToText`: entities decoded,
  `<br>`/`<p>`/`<hr>`/list items become newlines) instead of stripping tags off
  the raw string, and the body is cut — not whitespace-flattened — at 4 000 chars
  by default (`body_chars`, max 20 000) with `body_truncated`.
- `thread` — `ConversationId eq '<id>'`, compact rows, no bodies by default
  (`include_body: true` adds text and selects `Body`; bodies are cut at 500
  chars by default, `body_chars` max 2 000, with `body_truncated`; HTML is
  converted with `htmlToText` before the cut). Every row
  carries `conversation_id`, so any list/search/get result can seed a thread.
  A capped read reports itself: `count`, `truncated` when the service has more,
  and `next_cursor` to fetch the next page (a cursor from another API host is
  rejected with `invalid_cursor`).
  The query deliberately omits `$orderby`: Outlook rejects
  `$orderby ReceivedDateTime` with a `ConversationId` filter (`InefficientFilter`),
  so the page is sorted locally.
- Row: `{id, conversation_id, subject, from, received, received_local, unread,
  flag, preview ≤preview_chars, has_attachments, next_cursor}` plus optional
  `web_link`. `received_local` is the received instant in the configured
  timezone, so ordering and "is it today" need no client-side conversion.
  `preview_chars` (40–1000, default `max_results.preview_chars`, 300) caps the
  preview text; `include_body_preview: false` removes it and wins over the cap.
  `brief: true` keeps only `{id, subject, from, received, received_local,
  unread}` (+`has_attachments` when true) for cheap triage. `web_link` is
  omitted by default (`include_web_link: true`), because it is ~230 B/row of
  tenant URLs. `include_unsubscribe: true` selects `InternetMessageHeaders` for
  the page and adds `unsub: {has, one_click?}`; the verbose `url`/`mailto` are
  only inlined with `include_unsubscribe_targets: true` (≈1.2 KB per 25 rows), so
  `unsubscribe_preview` stays the detail path for one message.
- `unsubscribe_preview` — reads `List-Unsubscribe` / `List-Unsubscribe-Post` and
  returns `{mailto, url, oneClick, raw}`; nothing is posted.
- `send` — `POST /me/sendmail` with `Message{Subject, Body{ContentType, Content},
  To/Cc/BccRecipients, Importance?}` and `SaveToSentItems: true`. REST answers
  **202 with an empty body**; the result is
  `{ok, op:"send", recipient_count, recipients}`
  plus a best-effort `id` only when exactly one Sent Items copy matches (same
  subject, receive window after the 202), otherwise a `note`. An id is never
  invented. Recipients are resolved through aliases/directory and
  rejected when `invite_allowlist` is set and they fall outside it.
- `folders` — `GET /me/mailfolders?$top=50&$select=Id,DisplayName,
  TotalItemCount,UnreadItemCount,ChildFolderCount`, one page. Rows are
  `{id, name, total?, unread?, child_folders?}`; the result is `count` plus
  `truncated` when the page was full.
- `attachments` — `GET /me/messages/{id}/attachments?$select=Id,Name,
  ContentType,Size,IsInline&$top=20`. Returns `{message_id, count, attachments,
  truncated}`, each row `{id, name, content_type, size, is_inline}`. No content
  is requested; the list is capped at 20 with `truncated: true` when more exist.
  This is the read that makes an attachment-only message intelligible.
- `get_attachment` — `GET /me/messages/{id}/attachments/{attachmentId}` with
  `ContentBytes`. `text/*` and `application/calendar*` are decoded as UTF-8 and
  capped at 20 000 chars (`text`, `text_truncated`); `text:true` overrides the
  media-type test and decodes any attachment as UTF-8. Every other type is
  written to `~/.cache/teams-mcp/attachments/` (dir 0700, file 0600) under a
  sanitised, basename-only name plus a random suffix, returning `saved:true` and
  `path`.
  Decoded content over 10 MB returns `note:"too_large"` without saving; a
  reference attachment with no `ContentBytes` returns
  `note:"content_unavailable"`. Files are never opened or executed, and both ops
  are pure reads: no write token, no rate limit, allowed in read-only mode.
- `draft` — `POST /me/messages` with the same `Message` body as `send` and no
  `SaveToSentItems` (that key belongs to the sendmail envelope). Outlook creates
  the message in **Drafts** and returns the object, so the result is
  `{ok, op:"draft", recipient_count, recipients, id?, web_link?, note}`. Nothing is sent,
  but it is a write on every gate: read-only refuses it, it takes a rate token,
  recipients are resolved and allowlist-checked, and the attempt is audited.
- `send_draft` — `POST /me/messages/{id}/send`. Requires `draft_id` and
  `confirm:true` (otherwise `confirmation_required`, nothing sent). The draft is
  re-read immediately before sending (`Id,ConversationId,Subject,
  ToRecipients,CcRecipients,BccRecipients`) because it may have been edited in
  Outlook since `draft` created it: no recipients is `invalid_argument`, and any
  recipient outside `invite_allowlist` is `recipient_not_allowed` — both before
  the send. The Sent Items id/note is reported like `send`.
- `reply` — `POST /me/messages/{id}/reply {Comment}`, or `/replyall` with
  `reply_all:true`; the original is read first
  (existence, sender, `ConversationId`, `Subject`) so `recipient_count` is
  factual. Reply recipients are derived by the service. For reply-all the
  server derives the same set itself (sender + To + Cc, de-duplicated
  case-insensitively) and checks it against `invite_allowlist` **before** the
  send, so no recipient outside the allowlist is reached. After the 202 the sent
  copy is looked up in Sent Items (same conversation, `ReceivedDateTime` window,
  subject match with a prefix-insensitive fallback because Outlook may rewrite
  `FW: X` → `RE: X` on reply). Exactly one match returns `id`; zero or several
  return a `note` instead — a guessed id would be worse than none. Lookups retry
  briefly while Sent Items indexes (400/900/1800 ms).
- `forward` — `POST /me/messages/{id}/forward {Comment?, ToRecipients}`,
  target resolved and allowlist-checked; the source message is read for
  `ConversationId`/`Subject` so the forwarded copy can be identified the same
  way (expected subject `FW: <original>`).
- Write acknowledgements echo the resolved addresses as
  `recipients: {to?, cc?, bcc?}` (empty lists omitted) for `send`, `draft`,
  `send_draft`, `reply` (including `reply_all`) and `forward`. The caller can
  confirm where a message went without a follow-up `mail.get`; `send_draft`
  reports the addresses re-read from the live draft immediately before sending,
  not the ones passed when it was created.
- Recipient resolution (shared with calendar): a value with `@` is an address
  (a `Name <addr>` wrapper is reduced to the bare addr-spec before the allowlist
  check and the write); a configured alias wins; anything else is looked up in
  Graph `/users`. Exactly
  one match is used; several return `ambiguous_recipient` with all candidates and
  `needs_disambiguation`; none returns `recipient_not_found`. Resolution happens
  before the allowlist check and before any write.
- `mark` — `PATCH /me/messages/{id}` with exactly one of `{IsRead}` or
  `{Flag:{FlagStatus}}`; the result echoes `id` and the new `mark` only.
- `move` — `POST /me/messages/{id}/move {DestinationId}`; returns the **new** id
  (`from_id` keeps the old one) because moving folders changes the id.
- `unsubscribe_preview` — read-only parse of `List-Unsubscribe` /
  `List-Unsubscribe-Post`.
- `unsubscribe_apply` — requires `confirm:true`; without it the result is
  `{refused:true, reason, preview}` and nothing is sent. With it, at most two
  targets of the *same* message are acted on: an `https` one-click POST with body
  `List-Unsubscribe=One-Click` and a `mailto:` target via sendmail with subject
  `unsubscribe`. The POST is unauthenticated but vetted: https only, no embedded
  credentials, no redirects followed, 10 s timeout, 64 KB bound on the reply, and
  the host must not be loopback/private/link-local — by literal address or by
  every DNS answer (SSRF guard). Refusals are
  `{status:"refused", refused_reason}`; per-target `{target, type, status,
  http_status?}`, `partial` when one of two failed.

### calendar

- `list` — range defaults 09:00 today .. 17:00 local seven days out, `top` ≤ 50,
  `$select` minimal, `next_cursor` + `truncated`; rows include a `local` range
  and omit `web_link` unless `include_web_link: true`. `brief: true` keeps `id,
  subject, start, end, local, show_as` for cheap scanning. `calendarview`
  responses are UTC and normalised to UTC ISO. A bare `YYYY-MM-DD` bound means
  local midnight and a bare wall clock is read in `timezone` (a DST-gap time is
  rejected by name, not slid an hour). All-day events arrive as UTC midnight with
  an exclusive end date, so their row label is `YYYY-MM-DD (all day)` rather than
  a converted evening.
- `agenda` — the compact day/range view: `[start, end)` defaults to the whole
  local day of `start` (today by default), computed DST-safely; `top` default 25,
  cap 50, with `truncated`/`next_cursor`. Rows are
  `{date, start_local, end_local, time, subject}` plus `id` unless
  `include_ids: false`. No attendees, location, organizer or links — it is the
  cheap "what's my day" read, not a replacement for `list`/`get`.
- `get` — one event with attendee response status and `web_link`.
- `create` — `POST /me/events` with `Subject, Body, Start/End {DateTime,TimeZone},
  Attendees, Location, IsAllDay, ShowAs, Sensitivity`. `IsReminderSet` is never
  sent (it does not exist on the v2 entity and makes the POST 400). Defaults:
  timezone from config, start = next 15-minute boundary, end = start +
  `default_duration_min`; `end > start` is validated before the request.
  `duration_min` supplies the end when `end` is absent (from `start`, or from
  the default start when neither bound is given); supplying it together with
  both `start` and `end` is `invalid_argument` before any request.
  All-day events use local dates: start at 00:00 local, end exclusive on the next
  day. A `start`/`end` without `Z`/offset is parsed in `timezone`. Attendees are
  resolved (names/aliases/addresses) and allowlist-checked.
  The POST response is a *wall clock in the timezone that was sent*, so the
  normaliser converts through the returned `TimeZone` before reporting
  `start`/`end`/`local`; the response therefore matches a subsequent `get`.
- Conflict check (`create` and time-changing `update`, `conflict_policy`
  `warn`|`refuse`|`ignore`, default `warn`): after the read-only, rate-limit and
  allowlist guards, the tool reads the user's own `calendarview` for
  `[start, end)`, keeps events that truly overlap (`eventStart < end &&
  eventEnd > start`), drops `ShowAs=free` rows (they do not occupy the time) and,
  on update, the event with the same id. Rows are capped at 5 with the true
  count in `conflict_count`. A successful create/update adds `conflicts` and
  `conflict_count`/`conflict_warning` when non-zero; `refuse` instead returns
  `conflict_detected` with the same rows and writes nothing (the rate token is
  refunded); `ignore` skips the read and reports `conflict_policy:"ignore"`. If
  the read itself throws under `warn`, the write proceeds and the result carries
  `conflict_check:"unavailable"` — an unreachable check is not silently reported
  as "no conflicts". `conflict_policy` is validated before the rate token, so an
  unknown value is `invalid_argument` with no request.
- `update` — `GET` then subject guard, then `PATCH /me/events/{id}` with only the
  supplied fields; start/end are always sent as a pair. Bound rules: `start`
  alone keeps the event's existing duration (end = new start + old duration, not
  the old end); `end` alone keeps the existing start; `duration_min` derives the
  missing bound (with `start`, with `end`, or alone from the existing start);
  `shift_min` moves both bounds by a signed offset and keeps the duration.
  `shift_min` with `start`/`end`/`duration_min`, and `duration_min` with both
  bounds, are `invalid_argument` before the rate token. Every update returns
  `changed{start_before, start_after, duration_before_min, duration_after_min}`;
  the conflict check runs only when a time field changed, excluding this event.
  New `attendees` are resolved and **allowlist-checked
  before the PATCH** (this was a bypass fixed in QA). Wall-clock PATCH responses
  are converted through their `TimeZone` the same way as create.
- `create`/`update` responses also carry `attendees_detail` (`email, name, type,
  status`), exactly the array `get` returns, so an invitee's state needs no
  re-fetch.
- `delete` — guard, then `DELETE /me/events/{id}` (204 verified); returns
  `{deleted:true, id, subject}`.
- `respond` — guard, then `POST /me/events/{id}/accept|decline|tentativelyaccept`
  with `{Comment?, SendResponse:true}`.
- `instances` — `GET /me/events/{id}/instances?startDateTime&endDateTime`, compact
  rows sorted by start within each page; `truncated` + `next_cursor` when the
  window holds more occurrences than one page (the endpoint returns them
  unordered and caps at `top`).
- Guard: title compared case- and whitespace-insensitively; a mismatch returns
  `subject_mismatch` and issues **no** write.

### availability

- Attendees (`freebusy`, `suggest`, `check`, `free`, `snapshot`) use the same
  resolver as mail and
  calendar: names/aliases/addresses, ambiguity is an error, and resolution
  happens before any REST call.
- `freebusy` — `getSchedule` busy blocks per attendee, merged, plus
  `availabilityView` digits (0 free, 1 tentative, 2 busy, 3 oof, 4 elsewhere) at
  30-minute granularity. Blocks are capped at 20 per attendee with an `omitted`
  marker. A cached snapshot is clipped to the requested `start`/`end`: range,
  busy blocks and the availability-view digits describe only that window, even
  when the snapshot covers a wider week. `busy` never contains `status:"free"`
  intervals (absence is free; the digits still encode it). Every attendee row
  carries `data_available` (from the snapshot's optional `metadata.data_available[]`
  list, falling back to "not in `unknown`" for older snapshots), so an empty
  `busy` array without data is distinguishable from a genuinely free mailbox.
- Unknown mailboxes are data, not free time. Exchange omits `WorkingHours` when
  it cannot resolve a schedule (typo, foreign address), which is the only
  reliable signal — an empty `ScheduleItems` list looks exactly like a free
  person. That flag flows `getSchedule` → `ScheduleAttendee.unknown` →
  `metadata.unknown` → freebusy `unknown: [emails]` + per-row `unknown:true`
  (and no digits for those rows) / suggest `excluded_unknown_attendees` +
  `warning`, with `unknown_attendees` when nothing resolves. `suggest`
  candidates add `verified` and `data_available` (true only when every requested
  attendee has schedule data **and** the service candidate covered each known
  attendee); when any requested attendee lacks data the result also carries a
  `data_warning`. A service candidate with no per-attendee availability gets
  `free: "unknown"` plus a `data_warning` instead of an `n/n` count, and local
  candidates inherit the same `verified` flag.
- `check` — one exact slot: `start`/`end` required (missing either is
  `invalid_argument` before `ensureSnapshot`), attendees default to the
  signed-in user. For each attendee it reports `status` (`free`/`tentative`/
  `busy`/`oof`/`working_elsewhere`/`unknown`), `data_available` and `unknown`;
  the top-level `available` is true only when every attendee is `free`. Busy
  blocks overlapping the slot are returned as `conflicts` (capped at 5, truth in
  `conflict_count`), clipped to the slot and localised. Unknown attendees make
  `available` false — an unverifiable slot is never reported free. No write, no
  rate token; allowed in read-only mode.
- `free` — explicit free windows from the snapshot's busy gaps, not a ranked
  list: `[start,end)` defaults to today..+7d, `duration_min` defaults to
  `default_duration_min` (5–480), `working_hours_only` default true,
  `buffer_min` 0–120, `max_slots` default `max_results.slots` clamped to 20. It
  skips `findmeetingtimes` entirely and returns
  `slots[{start, end, local}]`, `count`, `truncated`, `data_available` and
  `unknown` (windows are computed only from attendees with data). Read-only.
- `suggest` — see §6.
- `snapshot` — build/return the cached snapshot; returns path, local range,
  stale flag and counts, plus the resource URI.
- Organizer is not duplicated: attendees are de-duplicated case-insensitively.

### people

- `search` — Graph `$search` over displayName/mail/userPrincipalName with
  `ConsistencyLevel: eventual`; in-process TTL cache (`refresh` bypasses).
  Results include `id, name, email, title, department`.
- Ambiguity is explicit: more than one match returns every candidate plus
  `ambiguous:true` and a `needs_disambiguation` note. Nothing is picked silently.
- A full page is not a complete answer: `count` plus `truncated:true` and a note
  when the directory returned `top` rows (the flag is cached with the page).
- A configured alias matches exactly, bypassing search: the result carries
  `alias:true` and `resolved_to`.
- `resolve` — one row per input with `found:false` for 404s; aliases resolved
  before the call.

### teams

Two Graph tokens, selected per operation (never per user choice):

- **OWA `graph` token** — chats (`/me/chats`, messages), team and channel
  metadata, channel create, channel send. Chats are never created here.
- **Teams-web token** — channel messages and replies. It is the only credential
  here with `ChannelMessage.Read.All` (`src/auth/teams-web.ts`). It also holds
  `ChannelMember.ReadWrite.All` and `TeamMember.ReadWrite.All`; rosters are
  verified but not exposed as a tool op in this phase.

- `status` — offline read of `teams-web-token.json`: freshness, seconds left,
  source (`browser`/`refresh`/`access-only`), refreshable, scope count, account.
  Sessions are ~60–66 minutes; a refresh happens on demand and is persisted
  before use. A missing/dead refresh token falls back to a silent browser
  harvest (~18–25 s, warm profile, no prompt).
- `list` — `kind=teams` (`/me/joinedTeams`), `kind=chats`
  (`/me/chats?$select=id,topic,chatType,lastUpdatedDateTime,lastMessagePreview`),
  `kind=channels` (`team_id` required). Graph rejects `$top` on joinedTeams/channels, so those
  pages are capped client-side; chats/messages take `$top`. Chats accept
  `chat_type=group|oneOnOne|meeting` (server `$filter`) and `query` (topic or
  member display name, `$expand=members`, client-side); teams and channels accept
  `query` over `displayName` case-insensitively. Every scan continues until `top`
  matches, the provider is exhausted, or the 15-page budget is reached, so "all
  group chats"/"Robotics" is a real answer, not one page. Rows are
  `{id,name,description?,web_url?}` /
  `{id,name,chat_type,last_updated,last_message_preview,last_message_at,members?}`;
  chat rows are deduplicated by id across a scan (membership churn can repeat a
  row), keeping the newest `lastUpdatedDateTime`; `last_message_preview` and
  `last_message_at` are omitted when Graph exposes no preview (tool JSON is
  null-free), so absence means "unknown/none". Caps 25; opaque `next_cursor` with an exact page URL +
  row offset.
- `read` — `chat_id`, or `team_id`+`channel_id`. Message rows are
  `{id,from,created,text,web_url,reply_to?,message_type?,has_attachments?}`;
  HTML is converted to text and cut at **500** chars by default
  (`text_chars`, param max 1 000). `top` defaults to **10** (cap 25): a read pays
  for every message row, so ten answers most questions. Roster events
  (`messageType: unknownFutureValue`, body `<systemEventMessage/>`) are filtered
  out by default **before** `top` is applied, and paging continues until `top`
  human messages or no cursor. With `include_system:true` they are rendered from
  `eventDetail` as `{system:true, event:"member added", text:"member added: Name"}`.
  `top` is "exactly up to this many real messages", not "the latest N": system
  events are skipped before the cap, so a channel full of roster events may cause
  several provider pages to be read to fill it.
  Provider pages are fetched at 50 rows so a system-heavy channel still answers
  in a few requests. Every list/read/replies answer reports `has_more` (a
  `next_cursor` exists) and `truncated` (the call stopped before filling the
  requested `top` — page/scan budget — rather than merely having older rows);
  `next_cursor` itself is always kept for older history.
- `replies` — channel thread (`team_id`, `channel_id`, `message_id`); message
  text keeps the 1 000-char cap (`replies` has no `text_chars` parameter).
- `send` — `confirm:true` + `text` + exactly one target. `text` is capped at
  20 000 characters before any rate token or request (`invalid_argument`). A
  target must be one this server created **or** an operator-listed one:
  `teams.allowed_targets` (`chat:<chat_id>` / `channel:<team_id>/<channel_id>`,
  from config.yaml, the shell CLI, or `TEAMS_MCP_TEAMS_ALLOWED_TARGETS`).
  Anything else is `target_not_allowed` with no request; there is no model-facing
  flag, so an agent cannot widen its own write scope. Chat sends use the OWA
  token, channel sends the Teams-web token. No attachments, no mentions, no bulk.
- `create` — `kind=channel` only (`team_id` + `name` + `private:true`, optional
  extra `members` capped at 20). Chat creation is **refused** with
  `invalid_argument`: live Graph rejects `roles:[]`, omitted roles
  and `"member"` on `POST /chats` ("Each chat member must have at least one role
  specified. Supported roles are 'owner' and 'guest'"), so the only creatable
  chat hands every invitee ownership. This server does not grant chat ownership
  silently; create the chat in Teams and ask the operator to add it to
  `teams.allowed_targets` before posting.
  Channel creation checks `/groups/{team-id}/owners` for the signed-in user first
  and never creates a team. Created ids are written to `teams-state.json` so
  later sends default to them; `confirm:true` is mandatory.
- `delete` — channels only, and only ids recorded in `teams-state.json`
  (`target_not_created` otherwise, with zero requests). Requires `confirm:true`
  and `expected_name`; the live channel is read first and the name compared
  case/whitespace-insensitively (`name_mismatch` refuses without deleting).
  Uses the OWA token (the Teams-web token lacks `Channel.Delete.All` and is
  403). 204 or 404/410 removes the ledger row; chats are never deleted.
- Failure mapping: a Graph 403 becomes `missing_scope` naming the required
  permission (`ChannelMessage.Read.All`, `TeamMember.ReadAll`, …), because
  "Forbidden" does not tell an agent which token to use; a bare 403 (private
  channel not joined) stays `forbidden`; 404/410 become `not_found`.
- Graph rejects `$top` on some collections and returns inconsistent page sizes;
  every collection is cursor-driven. Filtered reads encode the exact page URL and
  row offset, so cutting a page never skips a row. Cursors are versioned and
  bound to the call shape that minted them (op, surface, ids and filters): a
  cursor replayed with a different op/target/filter — or one minted by an older
  build — is `invalid_cursor` with no request, so a filter can never skip rows an
  unfiltered offset hid.

### settings

- `get` (all effective values or a dot path) / `set` / `reload` / `validate` /
  `path`.
- Keys: `timezone`, `work_hours.{days,start,end}`, `default_duration_min`,
  `aliases.<name>`, `invite_allowlist`, `blocked_times`, `snapshot_dir`,
  `cache.*_ttl_minutes`, `test_event_prefix`, `max_results.*`,
  `teams.allowed_targets` (operator-only through the MCP tool).
- `set` merges into the existing document, validates the whole document with zod,
  refuses on any issue (naming the path), writes atomically with mode 0600, and
  returns the effective value plus `env_override` when an env var shadows it.
- No redaction: config holds no secrets. Credentials live in `token.json`.

## 4. Configuration and CLI

File: `${TEAMS_MCP_CONFIG_DIR:-~/.config/teams-mcp}/config.yaml`. Precedence:
**environment > config.yaml > defaults**. Re-read when mtime/size changes.

```yaml
timezone: America/Chicago
work_hours: {days: [mon, tue, wed, thu, fri], start: "09:00", end: "17:00"}
default_duration_min: 30
aliases: {joel: joel@example.com}
invite_allowlist: ["@example.com"]
blocked_times: [{start: "12:00", end: "13:00", reason: lunch}]
snapshot_dir: ~/.cache/teams-mcp/snapshots
cache: {own_calendar_ttl_minutes: 5, busy_ttl_minutes: 30, directory_ttl_minutes: 60}
test_event_prefix: "[test]"
max_results: {mail: 25, calendar: 50, people: 25, slots: 5, preview_chars: 300}
safety:
  read_only: false
  unsubscribe: {enabled: false, daily_cap: 5}
  write_rate_limit_per_minute: 30
```

Paths:

| Path | Purpose |
| --- | --- |
| `~/.config/teams-mcp/token.json` | dual-audience OWA token cache (outlook+graph), 0600 |
| `~/.config/teams-mcp/teams-web-token.json` | Teams-web Graph credential (37 scopes, own refresh lock), 0600 |
| `~/.config/teams-mcp/teams-state.json` | ids of chats/channels this server created, 0600 |
| `~/.config/teams-mcp/browser-profile/` | persistent Chromium profile (`TEAMS_MCP_PROFILE_DIR`) |
| `~/.cache/teams-mcp/snapshots/calendar.yaml` | calendar snapshot, 0600, deletable |
| `~/.cache/teams-mcp/attachments/` | downloaded attachment files, dir 0700 / files 0600, deletable |
| `~/.cache/teams-mcp/write-audit.jsonl` | append-only write journal, 0600, rotates at 5 MB |
| `~/.cache/teams-mcp/unsubscribe-state.json` | daily unsubscribe counter, 0600 |
| `TEAMS_MCP_OUTLOOK_BASE_URL` | Outlook REST root override |

Env precedence applies to safety too: `TEAMS_MCP_READ_ONLY`,
`TEAMS_MCP_UNSUBSCRIBE_ENABLED`, `TEAMS_MCP_UNSUBSCRIBE_DAILY_CAP` and
`TEAMS_MCP_WRITE_RATE_LIMIT_PER_MINUTE` override the YAML. A boolean env value
that is not a recognised true/false word raises `ConfigError` rather than
silently falling back to the file (a typo must not leave writes enabled).

CLI (no MCP client needed):

```
node dist/cli.js settings path
node dist/cli.js settings get [dot.path]
node dist/cli.js settings set <dot.path> <value>
node dist/cli.js settings validate
node dist/cli.js settings reload
```

Exit codes: 0 ok, 1 invalid config on `validate`, 2 usage, 3 refused write.
Values are coerced (`45`→number, `mon,tue`→list, `[..]`/`{..}`→JSON).

## 5. Calendar snapshots

Path `${TEAMS_MCP_CACHE_DIR:-~/.cache/teams-mcp}/snapshots/calendar.yaml`, atomic
write (tmp + rename), 0600.

```text
version: 2
metadata: {fetched_at, range_start, range_end, timezone, ttl_seconds, stale_after, source: outlook, attendees[], unknown[], data_available[]?}
events:  [{id, subject, start, end, all_day?, show_as?, location?, is_online?}]   # own calendar
busy:    {"alice@example.com": [{start, end, status}]}                            # getSchedule
```

`metadata.unknown` is required from version 2. A version-1 document (or any
snapshot lacking the field) fails validation on read and is rebuilt, because an
unrecorded unresolved mailbox otherwise renders as "free all day". The builder
iterates the requested attendees, not the provider response: a row the service
omits entirely is recorded in `unknown` with `busy: []`. `metadata.data_available`
is optional: it lists the mailboxes for which `getSchedule` actually returned
schedule data. When present it is authoritative for `freebusy`'s per-row
`data_available`; when absent (a snapshot written by an older build) the reader
falls back to "not marked unknown", so old files stay readable.

Refresh policy:

- own events: `cache.own_calendar_ttl_minutes` (5);
- others' busy: `cache.busy_ttl_minutes` (30), stored as `ttl_seconds`;
- the binding TTL is the shorter (own) one;
- refresh when absent, stale, range not covered, or `refresh: true`;
- memory → disk → `calendarview` + `getSchedule`; without a source and without
  disk data the tool returns `snapshot_unavailable` rather than inventing data.

No delta: `/calendar/delta` is 405 here, so every refresh is a bounded
time-window poll.

Writes are safe under concurrency: one refresh at a time per path inside a
process (parallel availability calls coalesce, even for `refresh:true`), a
UUID+pid temp name per writer, one retry if the temp vanishes before `rename`,
abandoned temps older than an hour cleaned on write, and an atomic rename so a
reader never sees a partial file and a failed write leaves the previous snapshot
intact. `busy` entries are stored without free intervals, and `metadata.unknown`
records unresolved mailboxes. The snapshot is exposed as the read-only resource
`teams-mcp://calendar/snapshot` (`application/yaml`) as the **full** YAML view
(can be ~10 KB for a week of events) for bulk planning; `freebusy`/`suggest` are
the cheaper path for one scheduling question, and no tool requires the resource.

## 6. Slot finding

`availability op=suggest`:

1. Resolve attendees (name/alias/address) and normalise `[start,end]` to ISO
   instants. `ensureSnapshot` builds/refreshes for exactly that window when the
   cached snapshot is stale or does not cover it. Attendees the service could not
   resolve are dropped from ranking with a warning (`excluded_unknown_attendees`)
   — never ranked as free; all-unknown refuses with `unknown_attendees`.
2. Ask `POST /me/findmeetingtimes` for candidates using the **requested**
   `TimeConstraint` (`MeetingDuration: PTnM`, `IsOrganizerOptional: false`,
   `MinimumAttendeePercentage: 100`, `ReturnSuggestionReasons: true`). Both
   `MeetingTimeSuggestions` and a `value` wrapper are accepted.
3. Always compute local candidates from the snapshot (`suggestSlots`) over the
   same requested window, for blocked-time/working-hours/buffer rules.
4. Drop service candidates outside the window, in the past, outside working hours
   when `working_hours_only` is true, or overlapping `blocked_times`. Drop local
   candidates carrying a `working_hours` conflict under that flag.
5. Service candidates win when any remain; otherwise the local list is returned
   with `source: "local"` and `service_error` when the service refused. With no
   candidates at all the result says so explicitly (`no_candidates`), rather than
   widening the window.

`max_candidates` defaults to `max_results.slots` and is clamped server-side to
20 (requesting more is accepted, not rejected). Candidates are always inside
`[start,end]`; the result's `range` echoes that window, not the snapshot's. The
result reports `returned` (count), `max_candidates` (effective cap) and
`truncated` (true when more in-window candidates existed than were returned);
local search also reports `total_available`. This is explicitly a ranked subset
— use `freebusy` for the full busy view.

`check` answers the complementary question (one exact slot) and `free` returns
the full set of open windows; neither calls `findmeetingtimes`. When the
requested attendees are just the signed-in user and the service returns
candidates with no per-attendee availability (`free: "unknown"`), `suggest`
treats the local snapshot search as authoritative and returns those slots
instead — a self-only "when am I free" question gets real, verified slots rather
than service candidates nobody can vouch for.

Local algorithm:

- candidates advance in absolute milliseconds on the `interval_min` grid inside
  the requested window;
- conflict window per attendee is `[start - buffer, end + buffer)`;
- working hours and blocked times are evaluated in local wall-clock time
  (Intl, configured IANA zone), so DST changes do not shift a 09:00 slot;
- a spring-forward gap has no labels: candidates step 01:30 CST → 03:00 CDT and
  no candidate is ever labelled 02:xx; a fall-back hour is offered twice, as two
  real instants with the same label;
- ranking: fewest conflicts, then earliest; overlapping results are dropped;
- output per candidate: `{start, end, local, free: "n/m" | "unknown", conflicts:
  "<email>: busy; …", verified, data_available, source}` — never busy arrays.
  `verified`/`data_available` are true only when every requested attendee had
  schedule data; otherwise the candidate is surfaced with `data_warning` at the
  result level rather than being claimed free.

## 7. Token store

`token.json` (singular), 0600, atomic rename. Version 1:

```json
{
  "version": 1,
  "account": {"upn": "…", "id": "…"},
  "refresh_token": "…",
  "tokens": {
    "outlook": {"access_token": "…", "expires_at": "ISO", "scope": "…"},
    "graph":   {"access_token": "…", "expires_at": "ISO", "scope": "…"}
  },
  "source": "refresh|browser|headless",
  "updated_at": "ISO"
}
```

- `getAccessToken({audience})` defaults to `outlook`; stale = `expires_at - 5min`.
- `persistTokenUpdate` merges one audience, preserves the other, and persists a
  rotated refresh token atomically *before* returning.
- Read-only legacy support: `tokens.json` OWA/MSAL captures (audience inferred
  from `claims.aud`) and the old single-audience v1 document (mapped to
  `outlook`). Both are normalised in memory and never rewritten.
- The login/refresh implementation is the auth workstream's (`src/auth/*`);
  `getAccessToken` uses it, and tests inject `deps` hooks to stay offline.
- Failure modes are typed: no cache at all → login; a corrupt/unreadable cache
  → `auth_required` with the path to delete (a browser is **not** launched for a
  corrupt file); refresh `invalid_grant` → one sign-in attempt; a transient
  refresh failure → `auth_required`/retry, no login. Cross-process refreshes are
  serialised by `<configDir>/token.lock` (stale after 120 s is taken over).
- The resolved-config memo is keyed by the config file *path* plus mtime/size,
  so switching `TEAMS_MCP_CONFIG_DIR` cannot return the previous dir's paths.

## 8. HTTP contract

`src/http/client.ts` is shared by both APIs:

- injects `Authorization: Bearer …`, `Accept: application/json`, `Content-Type`
  only with a body;
- retries 408/429/5xx and transport failures with bounded exponential backoff
  (`base * 2^attempt`, capped, ±25% jitter); `Retry-After` (seconds or HTTP-date)
  wins;
- 204/205 and empty bodies → `undefined`; non-JSON success → `invalid_response`;
- errors become `ApiError {status, code, message, retryable, retryAfterMs,
  requestId, hint}` with `toLine()`; 401 carries the re-login hint, 403 names the
  permission problem, 405 names unsupported endpoints;
- bearer-shaped text is redacted; headers and bodies are never logged;
- `ownsUrl()` guards `next_cursor` reuse against foreign hosts.

## 9. Token efficiency

1. Minimal `$select`; no `$expand` unless the op needs it.
2. Nulls omitted, deep (`omitNullish`), stable key order.
3. Previews default to 300 chars (`preview_chars` 40–1000, or
   `include_body_preview: false` for none); bodies 4 000 on explicit `get` only
   (`body_chars` can raise it to 20 000); `thread include_body` bodies stop at
   500 (max 2 000).
4. No `@odata`/`value` envelopes: rows are mapped once at the boundary.
5. Opaque cursors, server-side caps.
6. `availabilityView` digits instead of block lists when data is dense.
7. In-process TTL caches: directory, snapshot.
8. One-line errors.
9. `web_link` is opt-in on list rows (`include_web_link`), `brief` trims calendar
   and mail rows, and `include_unsubscribe` adds only `{has, one_click}`; verbose
   `url`/`mailto` need the explicit `include_unsubscribe_targets` flag.
10. Row-per-question views exist so a summary does not need a full record:
    `calendar.agenda` for a day/range, `availability.check` for one slot,
    `availability.free` for open windows, and `mail brief`/`preview_chars` for
    triage. They return fewer fields and shorter strings; the exact byte saving
    depends on the mailbox and window, so no fixed figure is claimed.

## 10. Safety model

Tool-level capability scoping: an early live-account test with an autonomous
agent performed bulk mailbox actions; the capability controls exist to prevent a
repeat.

1. **Read-only mode.** `safety.read_only` / `TEAMS_MCP_READ_ONLY=1` makes every
   mutating op return a `read_only` tool error *before any request*: mail
   send/draft/send_draft/reply/forward/mark/move/unsubscribe_apply; calendar
   create/update/delete/respond; `teams.send`/`teams.create`/`teams.delete`;
   `settings.set`; `auth.logout`. Reads,
   availability (including `check`/`free` and the read-derived snapshot cache),
   attachment
   metadata/content reads (`attachments`, `get_attachment`), `auth.status` and
   `auth.login` remain available — a read-only server still needs a token.
   The guard lives in each handler (not only the registry wrapper), so direct
   handler calls and any future transport get the same guarantee, and tests
   assert `zero fetch calls` per blocked op.
2. **Unsubscribe is opt-in.** `safety.unsubscribe.enabled` defaults to **false**;
   `unsubscribe_apply` refuses with `unsubscribe_disabled` and zero requests.
   When enabled it additionally requires `confirm: true` **and**
   `expected_sender` matching the message's From address (case-insensitive), so
   a stale message id cannot act on a different sender. A daily cap
   (`safety.unsubscribe.daily_cap`, default 5) is persisted in a 0600 state file
   and returns `cap_reached` with the next local-midnight reset. The https/mailto
   targets are the message's own; the SSRF guard from §6 still applies.
3. **Write rate limit.** A per-process token bucket (default 30/minute,
   configurable) covers every mutating op: mail
   send/draft/send_draft/reply/forward/mark/move/unsubscribe_apply, calendar
   create/update/delete/respond, `teams.send`/`teams.create`/`teams.delete`,
   `settings.set` and `auth.login`/`auth.logout`.
   It is checked before the request; a refusal is `rate_limited` with
   `retry in Ns`. Invalid arguments do not consume a token, and reads are never
   limited.
4. **Write audit.** Every write that reaches a tool's write path — allowed,
   policy-refused (unsubscribe confirm/sender/cap) or blocked by read-only /
   rate-limit — appends one JSON line to
   `~/.cache/teams-mcp/write-audit.jsonl` (0600): `ts`, `tool`, `op`, an
   id/recipient-count/destination `target`, and an outcome/error code.
   Bodies, subjects and tokens are never logged. The file rotates to
   `write-audit.jsonl.1` at 5 MB. Audit failures never change the operation.
5. Reads before writes: ids from a prior read; `calendar delete` additionally
   requires `expected_subject` and aborts on mismatch.
6. `invite_allowlist` (when non-empty) rejects outsiders before any write, on
   create **and** update. Names without `@` are directory-resolved first (mail,
   calendar and availability); an ambiguous or missing name is an error, never a
   guess.
7. Owner-only files: `token.json` 0600, profile 0700, snapshots 0600, audit
   journal and unsubscribe counter 0600, attachment cache dir 0700 with files
   0600. Downloaded attachments are never opened or executed.

Scope: these are capability boundaries for a caller that sees only the MCP
surface. Anyone who can edit `config.yaml` or the environment can change them;
this is not a sandbox against the local user.

## 11. Layout

```
src/index.ts             stdio MCP server + shutdown
src/cli.ts               settings CLI
src/config.ts            YAML + env, hot reload
src/settings.ts          settings get/set/validate, atomic 0600
src/safety.ts            read-only gate, unsubscribe policy/cap, rate limit, audit
src/time.ts              timezone/DST helpers
src/format.ts            compact JSON, truncation, projection
src/http/client.ts       shared retry/backoff/error transport
src/outlook/client.ts    REST v2 client + PascalCase boundary + mappers
src/outlook/attachments.ts attachment metadata + content fetch, text decode, 0600 save
src/outlook/calendar.ts  calendarview, getSchedule, findMeetingTimes, source
src/graph/client.ts      Graph directory (/users) only
src/auth/store.ts        dual-audience OWA token cache
src/auth/teams-web.ts    Teams-web credential: MSAL decrypt, refresh, harvest
src/auth/lock.ts         cross-process refresh locks
src/teams/client.ts      Graph client for the Teams-web token + 403 mapping
src/teams/normalize.ts   compact team/channel/chat/message rows
src/teams/state.ts       created-target ledger
src/snapshot/calendar.ts snapshot build/read/write/refresh
src/availability/slots.ts local slot search + availability view
src/tools/*.ts           7 tools + registry + snapshot resource
src/tools/recipients.ts  name/alias/address resolution (mail, calendar, availability)
scripts/dev/measure-surface.mjs  bytes + token estimates for surface/outputs
```

## 12. Live verification

`LIVE_WRITES=1 npm run live-writes` drives the built server over stdio exactly as
an agent would:

- `calendar.create` (+2h, 30 min, no attendees) → `update` subject (PATCH) →
  `get` → wrong-`expected_subject` delete refused → `delete` → `list` shows no
  leftover. All succeeded; the deleted event did not reappear.
- `teams` live: chat/channel reads, `chat_type=group` returned every
  group chat in one call (2 scanned pages, complete), a channel `read top=10`
  returned 10 human messages with system events filtered
  (`truncated:true` + cursor for older history → complete), create private
  channel → send → read-back → **`delete` op** (expected_name guard) → 204 → the
  read afterwards is `not_found` (410), channel absent from the list, ledger row
  pruned. Forced headless refresh: `source browser -> refresh`, 5 165 s left,
  2.2 s. Refusal paths (`confirmation_required`, `target_not_created`) are in the
  audit journal with zero requests.
- One `mail.send` to the signed-in user (202) → `mail.search` found it →
  `mark unread` → `mark read` (verified via `get`) → `move` to `archive`
  returned a **new id** and `get` on that id succeeded.
- `unsubscribe_preview` on a real marketing message returned
  `{url, mailto, oneClick:true}`; `unsubscribe_apply` was deliberately not called.

REST behaviour that differed from the recon notes:

1. `PATCH /me/events/{id}` and `POST /me/messages/{id}/move` work as documented
   but were not probe-tested before; both returned 200 and the move changed the
   message id (as expected, and now handled).
2. `mail.search` over `/me/messages` finds the self-sent message immediately, so
   no Sent Items folder traversal is needed.
3. `calendar.update` requires Start/End as a pair in our client because the patch
   is validated against the existing half; the service itself accepted it.
4. No 429s occurred during the probe; the retry/backoff path remains
   fixture-tested only.

An independent QA pass found:

- Every advertised op was invoked over MCP: reads live, writes bounded
  (one tagged event created → guarded delete refusal → guarded delete; one
  self-email → get → mark → move), and reply/forward/respond/instances plus
  `unsubscribe_apply` against a local mock REST endpoint.
- Write responses for create/update are wall clocks in the request's timezone.
  The normaliser now converts through the returned `TimeZone`; a live create at
  15:00 America/Chicago reports `21:00Z` / `local 15:00`, matching `get`
  (previously it reported `15:00Z` / `09:00`).
- `calendar.update` now applies `invite_allowlist` to new attendees (was a
  bypass); names resolve through Graph with explicit ambiguity errors.
- Cached `freebusy` output is clipped to the requested window.
- `availability.suggest` now builds both service and local candidates from the
  requested `[start,end]` (not the cached snapshot's wider range), filters out
  out-of-window / out-of-hours / blocked service slots, honours
  `working_hours_only`, and clamps `max_candidates` to 20 instead of rejecting
  larger requests. Live: a 09:00–12:00 call returned 6 in-window slots, the next
  day's call returned 2, and `max_candidates: 200` was accepted.
- `include_unsubscribe` now emits `{has, one_click}` only; a 25-row page with
  three bulk rows measured 18,428 B, versus 19,589 B with
  `include_unsubscribe_targets: true` (1,161 B saved by default).
- `mail.get` defaults to a 4,000-char body cap; `body_chars` raises it to 20,000.
- Two processes sharing an expired token refresh once (file lock, one rotated
  refresh token, both calls succeed).

Not verified live: `mail.reply`, `mail.forward`, `calendar.respond`,
`unsubscribe_apply` (mock REST endpoint only; no real recipients).

## 13. Known gaps

- `src/auth/*` is outside the tool surface; `auth status`/`login`/`logout` are the
  integration points. Login/refresh is implemented and covered by tests and a
  real warm-profile login.
- `mail.reply`, `mail.forward`, `calendar.respond`, `unsubscribe_apply` are
  verified against a mock REST endpoint, not the tenant.
- Mail deletion is intentionally not exposed.
- `people.resolve` relies on Graph `/users/{id|upn}`.
- `mail.mark`/`mail.move` act on an id from a prior read without a subject guard
  (calendar writes have one); no mail delete op exists. Both are one object per
  call and share the write rate limit.
- Read-only mode is per process and cannot be toggled through the MCP
  `settings.set` op (that op is itself blocked); an operator must edit
  `config.yaml`, change the environment, or use the shell CLI.
- `unsubscribe_apply`'s daily counter is per machine/state file, not per account;
  two config dirs pointed at the same mailbox have independent caps.
- Snapshot never uses delta; polling windows are capped by the caller's range.
