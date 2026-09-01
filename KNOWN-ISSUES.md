
## send_chat_file: recipients get no permission on the uploaded item (found 2026-08-20)

The tool uploads to the signed-in account's OneDrive and posts a file-reference card, but never
grants the chat's members permission on the drive item. In a one-on-one or meeting chat the
recipient hits "request access", and the request lands in the service account's unread mailbox.
Observed live: a chat member could not open a file the assistant had shared; fixed by hand with
a Graph `POST /me/drive/items/{id}/invite` (roles: read, sendInvitation: false), which worked
instantly.

Fix direction: after upload, enumerate the chat's members (`GET /chats/{id}/members`) and invite
each with read role before posting the card — or create an organization link scoped share if
policy allows. Until then, senders must expect the access-request dead end.

## Inbox poller: two server instances race on the same inbox (found 2026-08-21)

Every session that starts the server gets its own inbox poller, and they all default to the same
`~/.teams-assistant/inbox.jsonl`. Two Claude Code sessions at once means two pollers appending
the same messages and overwriting each other's `inbox-state.json`: duplicated lines, watermarks
jumping backwards, and extra Graph load. A mild version showed up during a smoke run, where an
older standalone polling daemon and the new in-process poller polling together earned a 429 from
Graph.

Fix direction: a lock file next to the inbox (first server takes it, later ones skip the poller
and say so on stderr), or a per-session `TEAMS_INBOX_PATH`. Until then, keep one session per
account, or set `TEAMS_INBOX_DISABLED=1` in the extra ones.

## The throttle gates are per process (0.2.0; per resource family since 0.2.1)

`GraphClient` closes a gate on a 429 — per resource family, or the global gate for an
application-wide throttle — so nothing in that process keeps feeding Graph's penalty window on
that family. Separate processes do not share it: a cron loop running `teams-post` in
a fresh process each time, a second server instance, or an orchestrator's own retry loop around
the CLIs is entirely ungated and will keep the throttle alive while the server believes it is
being quiet. Space such callers yourself (one attempt, then wait the window), or route them
through one long-lived process. A cross-process gate (a lock file beside the token cache) is the
fix direction if this bites again.

**Partially addressed for `/members` in 0.4.1**: a live case of exactly this — two daemons on the
same first-party client id sharing one throttle budget, with a 429's Retry-After never actually
clearing under continuous consumption — made `GET /chats/{id}/members` unusable for @mention
resolution. 0.4.1 adds a disk-persisted per-chat member cache (default TTL 24h,
`TEAMS_MCP_MEMBERS_TTL_SECONDS`) so mention resolution no longer depends on that endpoint on the
common path; see the README's "@mentions" and "Throttle budgets are per client id" sections. The
underlying per-process gate limitation above is unchanged and still applies to every other
endpoint and to a `/members` cache miss itself.

## Retry-After was silent on the CLI send path (fixed 0.4.1)

Through 0.4.0, a Graph 429 on `teams-post`/`teams-reply`/`teams-edit` printed Graph's bare error
message with no indication of how long to wait, even though `GraphError` already carried
`retryAfterSeconds` internally — an operator watching the terminal had no signal to stop
blind-retrying into a worse throttle. Fixed: the CLI error output now appends
`(throttled, retry after Ns)` whenever Graph named a Retry-After.

## Inbox poller backfilled old messages on a lost or fresh watermark sidecar (fixed 0.4.1)

A restart was observed replaying roughly 40 old messages into `inbox.jsonl`. The per-chat
watermark sidecar itself already persisted correctly across an ordinary restart; the actual gap
was the documented "safe" fallback for a chat with no watermark on record (a genuinely fresh
install, or the sidecar lost/corrupted) — it re-read the chat's recent window and delivered
everything in it as new. Fixed: a chat with no known watermark now gets one settling poll that
establishes the watermark and delivers nothing; see the README's "The background inbox" section.

## Stuck-auth mode recoverable only by a process restart (defensive fix 0.4.1)

Observed twice live: the inbox poller failing every cycle while Graph itself was reachable, with
only a process restart clearing it. Code inspection could not conclusively pin the live root cause
to one line, but `RopcTokenProvider` trusting its cached token purely by local clock — with no way
for a live 401 to tell it the token had gone bad server-side — is a plausible mechanism that would
produce exactly this symptom. 0.4.1 ships the defensive fix the diagnosis calls for either way:
after several (default 3) consecutive auth-shaped poll failures, the poller forces the token
provider to drop its cached token and re-authenticate from scratch. If the stuck state recurs
under this fix, that is evidence the live mechanism is something else and needs a fresh
diagnosis — the counter and its threshold are `authFailureThreshold` on `InboxPollerDeps`.

## Quoted replies to old messages fail under the single-message throttle (0.2.1)

`GET /chats/{id}/messages/{id}` is throttled on its own budget. When it is, a quoted reply (which
fetches the original to build the quote card) and an attachment download fall back to the chat's
last 50 messages. A message older than that cannot be fetched until the throttle clears; the
reply fails with `MessageFetchThrottled` and nothing is posted. Post a plain message instead, or
wait the named window.
