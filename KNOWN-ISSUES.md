
## send_chat_file: recipients get no permission on the uploaded item (found 2026-08-20, fixed 0.4.2)

The tool uploaded to the signed-in account's OneDrive and posted a file-reference card, but never
granted the chat's members permission on the drive item. Recipients hit "can't be viewed or
downloaded" (a group chat) or "request access" (a one-on-one/meeting chat), and the request landed
in the service account's unread mailbox. Observed live twice: first 2026-08-20 (fixed by hand with
a Graph `POST /me/drive/items/{id}/invite`, roles: read, sendInvitation: false, which worked
instantly), then again 2026-09-02 confirming the tool itself was still unfixed — a sent file's
permissions listed ONLY `{roles:["owner"], to: the assistant account}` until the same manual invite
was repeated per chat member.

**Fixed 0.4.2**: `GraphTeamsChats.sendFile` now grants each OTHER chat member read access on the
uploaded item (that same `/invite` call) BEFORE posting the chat message, resolved through the
same cache-backed member roster resolveMentions already uses — never a direct call to the
throttled `/chats/{id}/members` endpoint on the send path (see README's "@mentions" section for
why that endpoint is avoided on sends). The assistant's own id is excluded from the grant when it
can be determined (it already owns the item as uploader). An unresolvable/empty roster, or ANY
other real member with no AAD id Graph reported (even in a mixed roster where some other members
ARE resolvable — no partial grants), now fails the send loudly BEFORE the upload; a failed
`/invite` call, or an `/invite` that answers HTTP success but whose own response shows no grant
actually landed for a recipient, fails loudly AFTER the upload (which is then left orphaned in
OneDrive — unavoidable, since the invite needs the uploaded item's id) and BEFORE the chat message
post. See `GraphTeamsChats.sendFile`'s own doc comment (`src/graph/teams-chats.ts`) for the full
failure contract; a card the recipients cannot open is refused rather than ever posted.

**Wire-shape anchor (live verification, 2026-09-02, EPF011 delegated token, OneDrive for
Business)** — the exact contract the code above and its tests are pinned to, captured verbatim
(GUIDs generalized): request `POST /me/drive/items/{id}/invite` body
`{"recipients":[{"objectId":"<aad-user-id>"}],"requireSignIn":true,"sendInvitation":false,"roles":["read"]}`
→ `200`; a subsequent `GET .../permissions` listed read grants with `grantedToV2.user` for all six
invited AAD users, and a human recipient confirmed the Teams file card opened. `grantedToV2.user.id`
is the field `sendFile`'s post-invite grant check (above) reads as the source of truth, rather than
trusting the `200` alone.

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

## Stuck-auth mode recoverable only by a process restart (defensive fix 0.4.1, widened in review round 1)

Observed twice live: the inbox poller failing every cycle while Graph itself was reachable, with
only a process restart clearing it. Code inspection could not conclusively pin the live root cause
to one line, but `RopcTokenProvider` trusting its cached token purely by local clock — with no way
for a live 401 to tell it the token had gone bad server-side — is a plausible mechanism that would
produce exactly this symptom.

**Incident artifact (review round 1), recorded verbatim rather than paraphrased:** the daemon
logged, repeatedly:

```
inbox poll failed: <chat>: fetch failed
```

No status code was visible in the log line. Parallel `curl` probes against Graph answered 200 the
entire time the daemon was stuck. This shape matches none of `isAuthShaped`'s vocabulary (401,
`invalid_grant`, an AADSTS code, "token"+"expir…") — a detector gated on that vocabulary alone
would never have fired on the actual incident.

0.4.1 originally shipped a detector gated on `isAuthShaped`; review round 1 widened it to two
tiers, same threshold (default 3, `authFailureThreshold` on `InboxPollerDeps`): auth-shaped
failures are the fast, well-understood path, and — because the evidence above shows the real
symptom is shapeless — ANY OTHER consecutive poll failure shape now also forces the same remedy
(drop the cached token, re-authenticate from scratch) as a last resort. A spurious forced re-mint
during a genuine network outage costs one extra password grant on the next successful call and
nothing else; staying stuck until a human restarts the process is the more expensive failure mode.
The remedy fires once per failing streak (not once per poll) and only resets once a subsequent
poll actually comes back clean — see `InboxPoller.trackAuthHealth`'s doc comment. If the stuck
state recurs under this fix, that is evidence the live mechanism is something else again and needs
a fresh diagnosis.

## Quoted replies to old messages fail under the single-message throttle (0.2.1)

`GET /chats/{id}/messages/{id}` is throttled on its own budget. When it is, a quoted reply (which
fetches the original to build the quote card) and an attachment download fall back to the chat's
last 50 messages. A message older than that cannot be fetched until the throttle clears; the
reply fails with `MessageFetchThrottled` and nothing is posted. Post a plain message instead, or
wait the named window.
