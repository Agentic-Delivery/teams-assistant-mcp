## A TTL-expired members cache under a throttled /members refresh converted a working cache into a permanent hard dependency on the throttled endpoint (live 2026-09-04, closed 0.5.2)

The 24h members-cache TTL (`DEFAULT_MEMBERS_TTL_MS`, `src/graph/members-cache.ts`) does not bound
staleness risk at the cost of one refresh under contention — it converts a working cached mention
path into a PERMANENT hard dependency on the very endpoint it exists to avoid. Root-caused live
(CTP daemon, `docs/throttling-mitigation.md` §1.2): the CTP agent-team chat's roster
(`fetchedAt` 2026-09-03T07:34:05Z) expired at 07:34:05Z the next day; the first throttled mention
came ~07:45Z, eleven minutes later. The mechanism: TTL expiry turns `MembersCache.get()` into a
miss → the miss triggers a `/members` refresh, the one endpoint sharing the crowded per-tenant
default-GET throttle bucket (§2 of that document) → the refresh 429s, so the fresh roster that
would have rewritten the cache is never obtained → every subsequent `--mention` post repeats the
same three steps forever, because healing the cache requires the very call being refused. Every
`--mention` post into that chat between ~07:45Z and ~08:05Z failed with `THROTTLED: the member
list refresh for mention resolution was throttled; nothing was done ... retry after 62s` while
UNTAGGED posts into the same chat succeeded throughout — proof this was never an account- or
tenant-wide block, only `/members` sharing a bucket `/chats/{id}/messages` does not.

**Closed 0.5.2**, two changes (`docs/throttling-mitigation.md` §4, stage 1 item 2):

1. **Stale-serve on a throttled/transiently-failing refresh.** `MembersCache.getStale()` is a new
   TTL-ignoring read `GraphTeamsChats.resolveMentions` (`teams-chats.ts`) falls back to ONLY after
   a live refresh itself fails with a 429/503/504 — never on an ordinary miss, which still refreshes
   as before. A cache hit resolves from the stale roster (logged once, naming the fallback and the
   roster's age) and only fails the post when the requested name is absent from the stale roster
   too, in which case the original THROTTLED error (naming Retry-After and, since 0.5.2, the
   throttle scope — see the entry below) still surfaces, never swallowed into a false "no such
   member".
2. **The roster now fills and refreshes itself from ordinary chat traffic, at zero Graph cost.**
   Every message the poller already reads carries its sender's AAD id and display name
   (`ChatMessage.fromId`/`from`, sourced from Graph's `from.user.id`/`from.user.displayName` —
   `messages.ts`'s `toChatMessage`); the poller merges those into the same on-disk roster
   (`MembersCache.merge`), refreshing `fetchedAt` on every merge that lands something. A chat that
   stays busy therefore never needs a live `/members` call again — the explicit refresh is the
   fallback for a name never seen in traffic, not the primary path. A message whose sender Graph
   could not fully identify (`from: 'unknown'`, no real displayName) is never merged, so a gap in
   Graph's own response cannot poison the roster with a junk entry.

Together these mean the 2026-09-04 incident's own mechanism cannot recur: TTL expiry no longer
forces a throttled call onto the hot path at all in a chat with any recent traffic, and even a
cold/silent chat degrades to "served from stale data, logged" rather than "hard failure until the
throttle happens to clear on its own." See the README's "@mentions" section for the consumer-
facing description.

## A partial (traffic-harvested) roster was trusted verbatim for send_chat_file's permission grant, and never expired (found and closed the same day, 0.5.2 review round 2, 2026-09-04)

The traffic-harvest mechanism above (mitigation 2) merges (senderId, senderDisplayName) pairs into
the on-disk members cache without any `/members` validation — deliberate, since the whole point is
zero Graph cost. A same-day review found the first cut of that mechanism let `membersForInvite`
(`src/graph/teams-chats.ts`, `send_chat_file`'s permission-grant roster) read that SAME cache entry
as if it were authoritative: a chat member who had never spoken (only harvested from OTHERS'
traffic, or never harvested at all) never appeared in the grant, so `send_chat_file` shared a card
they could not open — exactly the "no dead cards, ever" contract that method's own doc comment
names. Worse, because `merge()` used to stamp `fetchedAt` (the SAME field a real `/members` fetch
stamps) on every call, a harvested-only entry never expired, so a member who later LEFT the chat
kept their file-read grant indefinitely — the 24h TTL that used to force a periodic re-check never
fired for a roster that had never been fetched to begin with. The same partial roster also
resurrected the throttled-`/me` incident (entry below): `send_chat_file`'s self-id
roster-membership re-check saw the assistant absent from a roster that had simply never harvested
its own messages, and forced a live `/me` on every send into such a chat regardless of a valid
persisted self-id cache.

**Closed same-day, 0.5.2 (`src/graph/members-cache.ts`):** every cache entry now carries explicit
PROVENANCE — `complete: true` (and a real `fetchedAt`) for a roster confirmed by an actual
`/members` fetch, `complete: false` (`fetchedAt: 0`, `harvestedAt` instead) for one assembled only
from traffic. `membersForInvite` now reads the COMPLETE-only side of the cache
(`getComplete`/`getStaleComplete`) — a PARTIAL entry reads as a plain miss, forcing exactly the
`/members` call that used to be skipped (refusing the send loudly, THROTTLED, if that call itself
is throttled — no partial grants, ever). Mention resolution is unaffected: `resolveMentions` still
reads `get`/`getStale`, which serve a PARTIAL roster exactly as before, since a harvested-only
roster missing a name simply falls through to a live refresh there too — the hazard was specific
to a PERMISSION GRANT trusting incomplete data, not to mentions. On-disk shape stays
0.5.1-compatible: `fetchedAt: 0` (not omitted) on a partial entry keeps it shape-valid for an old
daemon's own `isValidEntry` check, which then reads it as instantly expired via its own TTL
arithmetic — never trusted as complete by an older process either. No operator action needed on
upgrade: a harvested-only entry already on disk simply reads as a miss on the next
`send_chat_file`, which pays one real `/members` call and moves on.

## send_chat_file: a throttled /me refused every CLI attempt, one process at a time (live 2026-09-03, mitigated 0.5.1)

`resolveSelfId`'s in-memory memo (`teams-chats.ts`) protects nothing across process boundaries,
and every standalone CLI invocation (`teams-send-file`) is a fresh process. Live 2026-09-03: eight
consecutive `teams-send-file` attempts over 20 minutes each answered the pre-upload refusal from
the 0.4.2 fix above — "the assistant's own account id could not be determined (the /me lookup
failed)" — because Graph was throttling `/me` for this app+user (two daemons share the first-party
client id), and each fresh process paid, and lost, the same throttled call. The account's own AAD
id never changes, so there was never a reason to re-resolve it live every time.

**Mitigated 0.5.1**: the signed-in account's own id is now persisted to disk
(`.self-id-cache.json`, next to the token cache — `FileSelfIdCache`, `src/graph/self-id-cache.ts`,
modeled on the existing members cache), with no TTL, and read back BEFORE any `/me` call is
attempted — a warm cache costs zero `/me` calls, so the first process to resolve it (server or any
CLI invocation) is the only one that ever pays the live lookup. `TEAMS_MCP_SELF_ID` is a
last-resort operator override (GUID-validated) for a throttle bad enough that even that first
resolution cannot land; it wins outright over the cache and the memo (subject only to sendFile's
roster-membership re-check — see below) and is never itself persisted, so removing it later
reverts cleanly to live/cached resolution. See `GraphTeamsChats.resolveSelfId`'s own doc comment
(`src/graph/teams-chats.ts`) for the full resolution order.

**Two cheap defences against a WRONG cached/overridden id (review round 2)**, closing the
realistic ways one reaches this far — a `TEAMS_MCP_SELF_ID` typo that still happens to look
GUID-shaped, or a stale `.self-id-cache.json` surviving `TEAMS_MCP_USERNAME` being repointed at
the same instance dir (a consuming project switching service accounts without also switching
`TEAMS_MCP_TOKEN_CACHE`) — both dangerous the same way: a wrong id that happens to equal a REAL
other chat member's id would silently exclude THAT PERSON from the grant instead of the assistant:
- `FileSelfIdCache` stamps the resolving account's `TEAMS_MCP_USERNAME` into every entry it writes
  and treats a stored entry from a DIFFERENT username (or a legacy entry with no username field at
  all) as a plain miss, not a wrongly-trusted id.
- `sendFile` never trusts a resolved self id that is not a member of the SAME chat roster it
  already fetched for the grant — a value nobody in that roster has forces exactly one live `/me`
  re-check (bypassing override/memo/cache), correcting the persisted cache on success.

**Known limitation, genuinely out of scope for 0.5.1**: a wrong id that happens to equal a REAL
OTHER member's id in the SAME chat passes both defences above undetected — the roster-membership
check sees a real member and is satisfied, and the username stamp only catches a different
account, not a same-account typo that happens to collide with a real member's own GUID. This is
the account-swap-adjacent case originally named here; nothing invalidates it, the cache is trusted
until the file is deleted by hand. The CLI-per-invocation `/me` cost is not eliminated either way,
only paid once instead of every time: the FIRST `teams-send-file` invocation against a cold cache
(a fresh install, or after the cache file is deleted) still pays one live `/me` call, same as
before this fix.

## The inbox poller starves ad-hoc reads on the same mailbox (measured 2026-09-02, mitigated 0.5.0)

Graph's read quota is per mailbox (client id + signed-in user together — see README "Throttle
budgets are per client id"), and a running inbox poller consumes it continuously. Measured live
2026-09-02: with the daemon polling, ad-hoc single-message GETs (`/chats/{id}/messages/{id}`)
answered 429 with `retry-after: 62` on EVERY attempt across 20+ minutes of patient
Retry-After-honouring backoff. Waiting is not enough: the waiting caller and the poller draw on
one budget, and the poller spends the replenishment the moment each window reopens. Attachment
downloads — reads by nature, and bursty — were dead on arrival next to a running poller.

**Mitigated 0.5.0**, three layers:

1. **Quota yield** (`src/inbox-yield.ts`): the attachment tools and the `teams-attachments` CLI
   write `inbox-yield.json` next to the inbox before touching Graph; every poller checks it at
   the top of each cycle and skips polling while it stands (logged once per yield, counted as a
   clean cycle so backoff never doubles); the file is removed when the reads finish, failure
   included, and its deadline (3 min default, 10 min hard cap) bounds what a crashed reader can
   silence. Works in-process and cross-process alike — the poller never cares who wrote it.
2. **The retry sleep cap moved off Microsoft's number**: `MAX_RETRY_SLEEP_MS` was a round 60s —
   one second UNDER the 62s window Graph actually names on this family, so every honest single
   retry was refused as "too long to sleep" by exactly that margin. It is 90s since 0.5.0.
3. **`TEAMS_INBOX_POLL_SECONDS`**: the poll interval (default 30s) is now a knob, for
   deployments where the same account routinely serves ad-hoc reads.

Not fixed, by design: two ad-hoc readers in ONE process can release each other's yield early
(cost: one contended poll cycle), and the poller checks the yield once per cycle, so a yield
written mid-cycle waits for the next one.


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
same on-disk cache resolveMentions uses, but read through its COMPLETE-only side — see the
"A partial (traffic-harvested) roster..." entry below for why sendFile never trusts the same
PARTIAL roster mention resolution may use — never a direct call to the
throttled `/chats/{id}/members` endpoint on the send path (see README's "@mentions" section for
why that endpoint is avoided on sends). The assistant's own id is excluded from the grant when it
can be determined (it already owns the item as uploader) — since 0.5.1, via the persisted self-id
cache first, a live `/me` only when that cache is cold (see the entry above); when it CANNOT be
determined by ANY of those means (a `/me` outage with a cold cache and no `TEAMS_MCP_SELF_ID`
override — see the second wire-shape anchor below), the send refuses BEFORE the upload rather than
falling back to an "invite everyone including self" default that would have orphaned an upload per
attempt for the whole outage. An unresolvable/empty roster, or ANY other
real member with no AAD id Graph reported (even in a mixed roster where some other members ARE
resolvable — no partial grants), also fails the send loudly BEFORE the upload; a failed `/invite`
call, or an `/invite` that answers HTTP success but whose own response shows no grant actually
landed for a recipient, fails loudly AFTER the upload (which is then left orphaned in OneDrive —
unavoidable, since the invite needs the uploaded item's id) and BEFORE the chat message post. See
`GraphTeamsChats.sendFile`'s own doc comment (`src/graph/teams-chats.ts`) for the full failure
contract; a card the recipients cannot open is refused rather than ever posted.

**Wire-shape anchor 1 — the GET readback (live verification, 2026-09-02, EPF011 delegated token,
OneDrive for Business)** — the exact contract the code above and its tests are pinned to, captured
verbatim (GUIDs generalized): request `POST /me/drive/items/{id}/invite` body
`{"recipients":[{"objectId":"<aad-user-id>"}],"requireSignIn":true,"sendInvitation":false,"roles":["read"]}`
→ `200`; a subsequent `GET .../permissions` listed read grants with `grantedToV2.user` for all six
invited AAD users, and a human recipient confirmed the Teams file card opened. `grantedToV2.user.id`
is the field `sendFile`'s post-invite grant check (above) reads as the source of truth, rather than
trusting the `200` alone.

**Wire-shape anchor 2 — the POST response itself (live verification, 2026-09-02, EPF011 delegated
token, OneDrive for Business, re-grant on the same test file)** — anchors the grant check against
the `/invite` response BODY directly, not just the follow-up GET above. Request body
`{"recipients":[{"objectId":"<johan-aad-id>"},{"objectId":"<epf011-owner-aad-id>"}],"requireSignIn":true,"sendInvitation":false,"roles":["read"]}`
(note: the OWNER included as a recipient, to settle whether an owner-as-recipient is silently
dropped) → `200` with (GUIDs generalized):
```json
{
  "@odata.context": ".../$metadata#Collection(microsoft.graph.permission)",
  "value": [
    {
      "id": "...",
      "roles": ["read"],
      "grantedToV2": {
        "user": {
          "@odata.type": "#microsoft.graph.sharePointIdentity",
          "displayName": "...",
          "email": "...",
          "id": "<johan-aad-id>"
        }
      },
      "grantedTo": { "user": { "...": "same id" } }
    },
    {
      "id": "...",
      "roles": ["read"],
      "grantedToV2": {
        "user": {
          "@odata.type": "#microsoft.graph.sharePointIdentity",
          "displayName": "...",
          "email": "...",
          "id": "<epf011-owner-aad-id>"
        }
      },
      "grantedTo": { "...": "same shape" }
    }
  ]
}
```
Two facts this settles: (a) the `POST` response itself — not only the later GET — carries
`grantedToV2.user.id` per recipient, which is what `sendFile`'s post-invite check reads directly
off the `/invite` call's own return value; (b) an owner included as a recipient IS echoed with its
own grant entry, not silently dropped — the worst-case "owner-not-echoed" shape a defensive review
raised turned out not to match live behaviour, though the code still refuses loudly rather than
assume that in general (see the failure contract above).

`grantedToIdentitiesV2[].user.id` (an array, sibling to `grantedToV2` on the same permission entry)
is a documented Graph variant some permission kinds use instead — several recipients' grants folded
under ONE permission resource rather than one `grantedToV2` entry each. Not captured in the two live
snapshots above (both used one recipient per permission entry), but `sendFile`'s grant check accepts
it as an equivalent echo defensively, so that shape alone degrades to acceptance rather than a false
outage if a future capture shows Graph choosing it.

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

(Cross-reference, added 2026-09-04: the "cache miss itself" residual risk this paragraph names is
what the 2026-09-04 incident actually hit — see the top entry of this file for the 0.5.2 close.)

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
