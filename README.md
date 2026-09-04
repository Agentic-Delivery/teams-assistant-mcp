# teams-assistant-mcp

An MCP server that lets a Claude Code agent read and post in a fixed set of Microsoft Teams group
chats. It signs in as an ordinary user account with a username and password, so from Teams' point
of view there is a person in the chat, not a bot.

It is not built to scale. One server per project, one account, one short list of chats. That was
the point: get a working two-way channel between an agent and the people in a pilot without
waiting on an app registration, admin consent, or a Teams app manifest.

This repo ships the server on its own. Consuming projects install the compiled output; none of
them carry this source tree.

## Companion skill: teams-styling

The repo also ships a Claude Code plugin, `teams-styling`
([plugins/teams-styling](plugins/teams-styling)), that teaches an agent how to compose
messages people actually want to read: when to style and when to stay plain, the HTML
vocabulary Teams verifiably renders (verified by screenshot against the real client, not
inferred from docs), known rendering quirks, and a live status-board pattern for making
agent progress visible in the chat. Install it by adding this repo as a Claude Code plugin
marketplace. Styled output needs the raw-HTML path the skill assumes: pass `format: 'html'`
to `send_chat_message`/`edit_chat_message` (default `'text'` still escapes everything), or
`--html` to `teams-post`/`teams-edit`. The caller owns entity-escaping `<`, `>`, `&` inside
their own content on that path — the server posts it verbatim.

## Getting started

**[SETUP.md](SETUP.md)** is the zero-to-working manual: requesting the account, `.env`, finding
chat ids, the allowlist, wiring the server into a Claude Code agent, and the tool reference. The
rest of this README is background on how and why the thing works.

## How it works

```
Claude Code  --stdio-->  teams-assistant-mcp  --HTTPS-->  Microsoft Graph  -->  Teams
                                |
                           allowlist  (refuses any chat id not on it)
                                |
                          TokenProvider  (ROPC today, swappable)
```

The server speaks MCP over stdio and exposes sixteen tools:

| Tool | What it does |
|---|---|
| `list_chats` | The allowlisted chats, annotated with whether the account can actually see each one |
| `read_chat_messages` | Messages from one chat, oldest first, with a watermark for the next call |
| `send_chat_message` | Posts to a chat whose allowlist entry has `canPost: true`; `format: 'text'` (default) escapes and renders, `format: 'html'` posts raw HTML verbatim; optional `mentions` — see "@mentions" below |
| `send_chat_image` | Posts a PNG/JPEG that renders inline, from a local path or base64 bytes |
| `send_chat_file` | Uploads a local file to the account's OneDrive (`TEAMS_MCP_UPLOAD_DIR`, default `ai-test`) and shares it into the chat, granting every other chat member read access on the uploaded item |
| `reply_chat_message` | Posts a quoted reply to a specific message — chats have no reply threads, so this is the quote card the Teams UI produces; optional `mentions` |
| `edit_chat_message` | Replaces the text of a message this account sent (Graph refuses anyone else's); same `format` and `mentions` options as `send_chat_message` |
| `react_to_chat_message` | Puts an emoji reaction on a message — the receipt gesture for "seen, being handled" |
| `delete_chat_message` | Soft-deletes a message this account sent — the reversible kind; no hard delete offered |
| `get_chat_attachment` | Downloads one attachment to a local file and returns the path — shared files, and pasted images which appear as `inline-image-N` |
| `list_chat_attachments` | The attachments on one message — id, name, contentType, downloadable — without transferring any content |
| `download_chat_attachments` | Downloads everything downloadable on one message in one call, optionally narrowed by a case-insensitive name filter and redirected to a chosen directory — see "Downloading attachments" below |
| `poll_chats` | Reads every allowlisted chat in one call, carrying a watermark per chat |
| `pin_chat_message` | Pins a message — see "Pinning" below for the one-pin-per-chat behaviour |
| `unpin_chat_message` | Unpins a message; refuses if it is not the one currently pinned |
| `list_pinned_messages` | Lists what is currently pinned, with a plain-text preview |

Besides the tools, the server runs a background inbox poller — see below.

All posting tools pass the same allowlist `canPost` gate. Editing and deleting only work on the
account's own messages — that is Graph's rule for delegated calls, and the server surfaces
Graph's refusal verbatim rather than pre-checking it.

Watermarks are exclusive ISO timestamps. Pass back what the previous call returned and you get
only what arrived since. When nothing is new no watermark comes back, so the caller keeps the one
it already had.

Graph returns message bodies as HTML even for plain typed text, so `messages.ts` flattens that to
text. Mentions keep their visible name and lose the markup. This is not a general HTML renderer
and does not try to be.

## @mentions

A plain `<at id="N">Name</at>` tag in the body does NOT notify anyone — Graph only rings the
bell when the message also carries a parallel `mentions` array entry with the person's real AAD
user id. `send_chat_message`, `reply_chat_message` and `edit_chat_message` all take an optional
`mentions: string[]` — display names to actually notify, resolved case-insensitively as an
unambiguous substring against the chat's current member list ("Shiv" matches "Garg, Shivankit").
A name matching zero or more than one member is refused with a clear error; nothing is ever
silently dropped.

How the name gets placed depends on `format`:
- `format: 'text'` (default): the mention's name must literally occur somewhere in your message
  text — that occurrence becomes the notifying tag. Just write the person's name where you mean
  to mention them ("Shiv can you review this?").
- `format: 'html'`: place a literal `@{Name}` token at each spot to mention (e.g. `@{Shiv}`); the
  server swaps it for the `<at>` tag. Every name in `mentions` needs a matching token and every
  token needs a declared mention, or the call is refused.

Only mention people who need to act — an owner, a question addressed to them. Tagging everyone
named in a message is how notifications stop meaning anything; see the `teams-styling` skill's
"Mentions" section for the full doctrine.

Mention resolution reads a per-chat member cache (disk-persisted next to the token cache, TTL
`TEAMS_MCP_MEMBERS_TTL_SECONDS`, default 24h) instead of calling Graph's `/chats/{id}/members` on
every send — that endpoint shares a throttle budget across every process signed in with the same
client id (see "Throttle budgets are per client id" below), and a chat's membership in this fixed
pilot allowlist effectively never changes. A cache hit resolves with zero Graph calls; a miss (no
entry, an expired one, or a name the cached roster does not have) refreshes once and re-checks — a
name still unresolved after that gets the usual clear error.

**A cache hit never becomes a hard dependency on the throttled endpoint (0.5.2).** Live-diagnosed
2026-09-04 (KNOWN-ISSUES.md): a `/members` refresh 429ing on an expired cache used to mean the
cache could never be rewritten, since rewriting it needed the very call being refused — the roster
was hostage to the throttle forever, not just for one bad attempt. So when a refresh answers 429
(or any transient 503/504), resolution falls back to whatever this chat's stale (past-TTL) roster
still holds on disk, logging one line naming the fallback, and only fails the post when the
requested name is absent from that stale roster too — in which case the original throttled error
still surfaces, naming Graph's Retry-After. A 429 is never swallowed into a false "no such member".

**The roster also fills itself from ordinary chat traffic, at zero Graph cost, as a PARTIAL entry
— never used for a file-share permission grant (0.5.2, updated in the same release after a
review found the first cut of this trusted a partial roster for that grant).** Every message the
background inbox poller reads already carries its sender's AAD id and display name
(`from.user.id`/`from.user.displayName` on the Graph chat message); the poller merges those into
the same on-disk roster the cache above serves mentions from (`MembersCache.merge`,
`src/graph/members-cache.ts`), stamping that chat's `harvestedAt` (never `fetchedAt` — that field
is reserved for a REAL `/members` fetch) on every merge that lands something. A chat that stays
busy therefore never needs a live `/members` call for MENTION resolution — the explicit refresh
above is the fallback for a name never seen in traffic, not the primary path. A message whose
sender Graph could not fully identify (an id with no display name — the `from: 'unknown'` fallback
below) is never merged, so a gap in Graph's own response can't poison the roster with a junk entry.

A roster built only from traffic is PARTIAL, not COMPLETE: it only knows who has spoken, not who
is silently in the chat. `send_chat_file`'s permission grant (below) never trusts a partial roster
— it always forces one real `/members` call for a chat that has no COMPLETE roster on disk yet
(refusing the send loudly, never granting a partial list, if that call itself is throttled), and a
COMPLETE roster's own freshness is judged only against its real `fetchedAt`, unaffected by
intervening traffic. **Operators: no action is needed on a restart or an upgrade to 0.5.2 — a
harvested roster on disk from an earlier build simply never drives a file grant; the very next
`send_chat_file` into that chat pays one real `/members` call and moves on.**

## Retry-After

Any Graph 429 in the send/reply/edit path never auto-retries — that has always been the rule for
writes (`GraphClient.post`/`ReliableTeamsChats`). Both consumer-facing surfaces now name the wait
whenever Graph named one — the CLIs (`teams-post`, `teams-reply`, `teams-edit`) in their stderr
output, and the MCP tools' error result text for whichever agent is driving the server directly:
`Too many requests (throttled, retry after 62s)`. One shared renderer (`retryAfterSuffix` in
`src/graph/graph-client.ts`) backs both, so neither surface can drift out of sync or silently lose
the wait time (0.4.1 review round 2: the MCP tool path originally missed it entirely).

### Throttle-scope diagnostics (0.5.2)

Graph's throttling is keyed by four independent, differently-scoped buckets (per app, per
app-per-tenant, per resource, per user — `docs/throttling-mitigation.md` §2.1), and a 429 alone
does not say which one closed. Every 429 `GraphClient` sees is now logged on one line naming
`x-ms-throttle-scope`, `x-ms-throttle-information` (when Graph sent one) and `retry-after` —
whether or not the read loop goes on to retry successfully, so a throttle that clears on its own
retry is still measured, not just the ones that end in a thrown error. The scope also travels
structurally on `GraphError.throttleScope` and is folded into the `THROTTLED: ...` text the
mention-resolution error surfaces (the exact error both CLIs and the MCP tool path print
verbatim), e.g. `THROTTLED: the member list refresh for mention resolution was throttled; ...
[throttle scope: Tenant_Application]`. The next incident is measured, not inferred.

## Pinning

`pin_chat_message` / `unpin_chat_message` / `list_pinned_messages` manage a chat's pinned
message(s). **A Teams chat effectively holds only ONE pin**: pinning a second message silently
REPLACES the first while Graph still reports the POST as a success (verified live 2026-08-25 and
reconfirmed 2026-08-26) — there is no "add another pin". `pin_chat_message` returns the resulting
`pinnedMessages` list so the replacement is visible rather than assumed. Listing an empty
collection is itself an undocumented quirk: Graph answers a bare 404 where a normal empty
collection would be `200` with `value: []` (verified live 2026-08-26); the server reads that 404
as "nothing pinned" and returns an empty list rather than surfacing it as a failure.

## The background inbox

Reading chats through tools means the agent has to remember to poll, and every session that
wanted to be woken by incoming messages ended up hand-rolling its own polling daemon. So the
server does it: alongside the MCP transport it polls every allowlisted chat (30s interval,
in-process, no child processes) and appends each new message as one JSON line to a stable,
session-independent file:

```
~/.teams-assistant/inbox.jsonl        the inbox (override with TEAMS_INBOX_PATH)
~/.teams-assistant/inbox-state.json   watermark sidecar, lives next to the inbox
```

One line per message: `{"chat","id","from","at","text","attachments"}` — `text` capped at 2000
characters, `attachments` a count. Messages posted by the signed-in account itself are skipped
(resolved via `/me`, so the assistant's own posts never echo back as inbox events), as are
deleted stubs and empty system events.

The sidecar remembers the delivered watermark and newest message id per chat (written atomically —
temp file then rename, so a crash mid-write cannot itself corrupt it), so a server restart never
re-emits old messages. A chat with no watermark on record — a genuinely fresh install, or a lost or
corrupted sidecar — starts from NOW: that first poll only establishes the watermark and delivers
nothing, rather than backfilling whatever already existed in the chat (0.4.1; a live restart had
backfilled roughly 40 old messages under the previous "re-read the recent window once" fallback).

If the poller sees several consecutive poll failures (default threshold 3), it forces the token
provider to drop its cached token and re-authenticate from scratch, rather than waiting on a
process restart to notice a token gone bad server-side (0.4.1; observed live: only a restart
recovered from this). Two tiers, same threshold: a 401 or an error naming `invalid_grant`/an
AADSTS code/token expiry is the fast, recognisable path; any OTHER consistent run of failures also
triggers the same remedy as a last resort, because the actual live incident's own log line —
`inbox poll failed: <chat>: fetch failed`, repeated, no status code, while Graph itself answered
200 to parallel probes — matched neither shape (see KNOWN-ISSUES.md for the verbatim evidence). The
remedy fires once per failing streak and logs what actually happened at each step (requested, then
still failing, or recovered), not just the initial intent.

A failing poll appends `{"error": "...", "at": "...", "consecutiveFailures": N}` to the same file.
That line is the difference between "the chats are quiet" and "auth is dead" — a watcher must
never have to guess which silence it is looking at. An identical failure repeating poll after poll
is written once, not once per poll — re-surfacing only at 10 and 50 consecutive failures with a
running count (`"still failing after N polls: ..."`), so a five-hour outage does not hide behind
one stale line and a watcher reading only the file's tail can still tell it is ongoing. A recovery
after ten or more failed polls writes `{"recovered": true, "at": "...", "afterFailures": N}`; a
short blip (fewer than ten) stays quiet, same as before. When *everything* fails (auth death,
network gone) the interval backs off, doubling to a 10-minute cap and snapping back on recovery —
a 429's `Retry-After`, when Graph names one, floors the next delay instead of the doubled guess,
so the poller never comes back sooner than Graph itself asked and feeds the penalty window it is
trying to back off from. A single failing chat — usually one the account has not been added to
yet — is surfaced but does not slow the healthy chats down. The poller never crashes the server;
every poll is fully caught.

The recommended consumption pattern: arm a file watcher (Claude Code's `Monitor`, `tail -F`,
inotify) on the inbox path at session start and react per line. Do not poll the tools for new
messages any more.

Three knobs: `TEAMS_INBOX_PATH` moves the inbox (the sidecar and the quota-yield file follow
it), `TEAMS_INBOX_POLL_SECONDS` changes the 30s poll interval (the poller is the main consumer
of the per-mailbox Graph read budget — see "Downloading attachments" for the measured
consequences), and `TEAMS_INBOX_DISABLED=1` switches the poller off entirely for consumers that
only post. The poller also honours the quota-yield file the attachment tools write — while it
stands, cycles are skipped (logged once per yield, not per cycle) and the backoff does not
double, so polling resumes at the normal cadence the moment the yield lifts. A yielded cycle is
recorded in the health file below as `yielded: true`, not as a fresh success.

## Supervising the daemon

On 2026-08-21 two Claude Code sessions running the server at once meant two pollers on the same
default inbox path, duplicating lines and racing on `inbox-state.json`; a related incident killed
one project's poller in a host reboot while a second project's daemon (same `dist/index.js` path,
different env) survived, and a pgrep-based liveness check matched the survivor — the dead pipeline
read as "up" and an allowlisted chat sat undelivered for 1.5 hours. pgrep is the wrong tool here on
principle: it answers "does some process matching this pattern exist", when the question is "is
THIS inbox's pipeline delivering". Two files beside the inbox answer that properly:

```
~/.teams-assistant/poller-health.json   liveness snapshot, rewritten after every poll
~/.teams-assistant/poller.lock          single-instance lock, one poller per inbox
```

**The health file** is the liveness contract. After every poll — clean, failed, or skipped for the
quota yield — the poller atomically rewrites it (tmp + rename, so a reader never sees a torn
snapshot):

```json
{ "pid": 1234, "inboxPath": "...", "lastAttemptAt": "...", "lastSuccessAt": "...",
  "ok": true, "consecutiveFailures": 0, "backoffMs": 30000 }
```

How a watcher should judge it: `lastAttemptAt` older than a couple of `backoffMs` (the delay the
poller itself announced before its next attempt) means the poller is dead or wedged, whatever the
process table says. A fresh `lastAttemptAt` with `ok: false` means the opposite failure: the
process is up but its polls are failing — process-up and pipeline-up are different claims, and the
file distinguishes them where pgrep cannot. `lastSuccessAt` survives failures (and yields), so it
also says how long an outage has run. A cycle skipped for the quota yield carries `yielded: true`
and does NOT advance `lastSuccessAt` — it is deliberate politeness, not a poll that actually ran,
and a watcher must not read it as either a fresh success or a failure.

**The lock** makes "one poller per inbox" checked instead of assumed. On start the server takes
`poller.lock`, keyed to the resolved inbox path (`dirname(inboxPathFor(env))` — the same helper
the poller and the CLIs use, so a `TEAMS_INBOX_PATH` override moves the lock with everything else
it already moves). If a live pid already holds it, the newcomer logs one line on stderr — naming
the lock path it lost — and **serves every other MCP tool normally, running no poller** for this
process; a dead holder's lock (the reboot case) is taken over automatically. This matters because
`dist/index.js` is also what an MCP tools registration runs (see "Installing into a
project"), against the same daemon `.env` more often than not: exiting on a contended lock would
silently drop every Teams tool for that session the moment a daemon was already polling, for a
reason no MCP client surfaces. The lock's only job is "one poller per inbox", and not starting a
second poller fully satisfies it — a process manager watching for a poller should watch the
health file above, not this process's exit code.

**The trap**: the lock is keyed per inbox PATH, not per config directory. An instance that sets
its own `TEAMS_MCP_CONFIG` / `TEAMS_MCP_TOKEN_CACHE` but leaves `TEAMS_INBOX_PATH` unset still
falls back to the same default inbox path as any other such instance, and therefore silently
contends for the same lock — from the losing instance's own environment, nothing looks shared. **A
second instance for a second account MUST set its own `TEAMS_INBOX_PATH`** (see `env.example`).

## Downloading attachments

Three tools cover the read side of files. `read_chat_messages` (and `teams-read`) show attachment
metadata on every message that carries any — name, contentType, id — so a reader can SEE there is
something to fetch; the inbox daemon's per-line `attachments` count is the coarse version of the
same signal. `list_chat_attachments` gives the same metadata for one message on demand, including
whether each entry is downloadable (a quoted-reply quote card is not). `download_chat_attachments`
fetches the message once and downloads everything downloadable on it — shared files and pasted
inline images alike — optionally narrowed by a case-insensitive name filter; a filter that matches
nothing is an error naming what the message does carry, never an empty success.
`get_chat_attachment` remains for grabbing a single attachment by id.

Two download mechanics, both handled: a file shared into a chat lives in SharePoint/OneDrive and
its `contentUrl` wants browser cookies, so the download goes through Graph's `/shares/u!{id}`
facade (unpadded URL-safe base64 of the URL — `shareIdFor` in `src/graph/teams-chats.ts`), which
the Graph bearer token CAN read; an image pasted into a message body is hosted content on the
message itself and comes from the `hostedContents` endpoint.

Where files land: `TEAMS_MCP_DOWNLOAD_DIR`, else a directory under the OS tmpdir; the download
tools also take an explicit output directory per call. Names are sender-controlled data and are
sanitized before writing (path components stripped, reserved and control characters replaced,
prefixed with the message id), and an existing file is NEVER silently overwritten — a collision
gets a `-1`/`-2`/… suffix (`src/downloads.ts`).

Each downloaded entry's result carries `bytes` — the number of bytes actually downloaded and
written, never null — so a caller can verify a download completed without re-stat-ing the file
itself (`get_chat_attachment`, `download_chat_attachments`, and the `teams-attachments` CLI all
report it the same way). `list_chat_attachments`/`--list` stay metadata-only, as before: nothing
is downloaded, so there is no byte count to report.

Permissions: with the shipped setup — a Microsoft first-party client id and the
`https://graph.microsoft.com/.default` scope — SharePoint downloads work out of the box, no extra
scope, no app registration, no admin consent (live-verified 2026-09-02 with this package's own
token). If `TEAMS_MCP_CLIENT_ID` is ever pointed at a custom app registration instead, that
registration needs the delegated `Files.Read.All` (or `Sites.Read.All`) Graph permission, which
may require admin consent — and a 403 from the download says exactly that, with Graph's own
error text preserved.

Throttling — and why retries alone are NOT enough: these reads share the per-mailbox Graph
budget with the inbox poller, and the poller spends that budget continuously. Measured live
2026-09-02: with a poller running, ad-hoc single-message GETs answered 429 with `retry-after: 62`
on every attempt across 20+ minutes of patient backoff — the waiting caller and the poller share
one budget, and the poller consumes it again the moment each window reopens. So the attachment
tools COORDINATE instead of just waiting: before touching Graph they write a quota-yield file
(`inbox-yield.json`, next to the inbox), every inbox poller checks it at the top of each cycle
and skips polling while it stands, and the file is removed the moment the reads finish — on
failure too. The yield's own deadline (3 minutes by default, 10 minutes hard cap) bounds how
long a crashed reader can silence the inbox. This works identically whether the poller runs in
the same process (the MCP server's own tools) or in another one (the `teams-attachments` CLI
next to a running daemon); the poller never cares who wrote the file. See `src/inbox-yield.ts`.

On top of the coordination, binary downloads retry a 429/503/504 through the same `GraphClient`
read loop as every other GET — Retry-After honoured, bounded, gate-aware (see "Send reliability"
below) — and the retry sleep cap is 90s, deliberately above the 62s window Graph actually names
on this family (the old round 60s cap sat one second UNDER Microsoft's own number, refusing
every honest retry by exactly that margin). Multiple attachments on one message download
sequentially, not in parallel. If ad-hoc reads matter routinely in a deployment, also consider
slowing the poller itself down: `TEAMS_INBOX_POLL_SECONDS` (default 30) trades inbox latency
for read-budget headroom.

## The standalone CLIs

Nine small commands ship beside the server for scripts, cron jobs and background monitors that
need Teams without a running MCP session: `teams-post <chatId> [--html] [--mention "Name"]...`
(text on stdin), `teams-reply <chatId> <messageId> [--mention "Name"]...` (text on stdin),
`teams-edit <chatId> <messageId> [--html] [--mention "Name"]...` (new text on stdin), `teams-react
<chatId> <messageId> <emoji>`, `teams-read <chatId> [--limit N] [--since ISO]`, `teams-pin
<chatId> <messageId>`, `teams-unpin <chatId> <messageId>`, `teams-send-file <chatId> <path>
[more paths...] [--caption "text"]` and `teams-attachments <chatId> <messageId> [--list]
[--name <filter>] [--out <dir>]`. Same allowlist, same auth, same
code paths as the server tools — including the send reliability below. `--html` on
`teams-post`/`teams-edit` posts stdin as raw Teams-subset HTML, verbatim — the caller is
responsible for entity-escaping their own `<`, `>`, `&`; see the `teams-styling` plugin for the
verified vocabulary. `--mention
"Name"` (repeatable) @mentions that person, same resolution and placement rules as the
`mentions` tool parameter — see "@mentions" above. `teams-send-file` uploads and shares one or
more files in one call (one `send_chat_file` per path); `--caption` (optional, anywhere in argv)
is shown above the FIRST file's card only, never repeated on every card. `teams-attachments`
downloads every downloadable attachment on a message into `--out` (else
`TEAMS_MCP_DOWNLOAD_DIR`, else a tmpdir) and prints the absolute paths; `--list` prints the
metadata instead and downloads nothing; `--name` narrows the download by case-insensitive
substring — see "Downloading attachments" above for the sanitization, collision and permission
story. `teams-read` includes each message's attachment metadata whenever there is any, so a
script reading a chat can tell there is something to fetch.

Their output contract exists because of a real incident (2026-08-24): an ad-hoc wrapper's
caller grepped for a success token the wrapper never printed, read eleven successful posts as
eleven throttles, and re-posted a broadcast ten times. So: success is exactly one JSON line on
stdout and exit 0; failure is prose on stderr and a non-zero exit (2 usage, 3 allowlist,
1 anything else). **Branch on the exit code, never on output text.** `teams-send-file` with
several paths STREAMS one JSON line per sent file as EACH one lands, rather than buffering until
the whole batch finishes: if a later file fails partway through, the earlier files' lines are
already on stdout and the exit code is still non-zero — the already-printed lines are proof those
files landed, so a caller must not blindly re-run the whole batch and re-send them.

## Send reliability: readback before retry

Every send (server tools and CLIs alike) goes through `ReliableTeamsChats`, which treats a
failure report as a claim about the response path, not about the chat. On any send error it
reads the chat back first — if this attempt's copy is standing, that copy is returned as the
success and nothing is re-sent. Only a readback that finds nothing leads to a retry, waiting
out any `Retry-After` the throttle named. Reads (GETs) retry ONCE inside `GraphClient` on
429/503/504, after the named wait; writes never auto-retry at that layer. A 204 or empty body on
a write is a success, not a parse error.

**The throttle gates (since 0.2.0, per resource family since 0.2.1).** Graph escalates its
penalty window when a caller keeps sending while throttled — and retries that look reasonable
one call at a time add up to exactly that. So a 429 closes a gate for the full `Retry-After`
(30 s when none is named): every request on that gate until the window passes fails fast,
locally, as `GraphError` code `LocallyThrottled`, without touching the network. Gates are keyed
per resource family (`/chats/{id}/messages/{id}` and `/chats/{id}/messages` are two gates —
Graph throttles them separately, and 2026-08-25 proved it: the single-message fetch was
refused for hours while the message list and plain posts on the same chat stayed healthy);
an application-wide throttle (error code `ApplicationThrottled` — the one code seen live; nothing
speculative) closes the global gate every request checks. A quoted reply and an attachment download both start with
that single-message fetch, so under its throttle they fall back to scanning the chat's last
50 messages; a reply to something older than that fails with `MessageFetchThrottled` —
nothing posted, the wait named. A throttled send waits the
window out BEFORE reading the chat back, then retries once — unless the named window exceeds
what one call may sleep (60 s), in which case it fails immediately with the 429 and its
Retry-After, no theatre first. If a send's outcome is genuinely
unknown (the response path died) and the gate then blocks the readback, the send reports
`UnknownOutcome` naming the original failure — never "not sent". The gates are per process: see
KNOWN-ISSUES.

## Auth, and the fact that ROPC is temporary

Sign-in uses the OAuth password grant (ROPC) against Entra ID, with one of Microsoft's own
first-party client ids. That is why no app registration and no admin consent are needed: the
tenant already trusts Microsoft's clients. The code defaults to the Teams client id (the
`TEAMS_FIRST_PARTY_CLIENT_ID` constant in `src/config.ts`), but on some tenants the token that
id returns carries no chat scopes at all, so `env.example` tells you to set the Office client id
instead — the one proven to hand back a fully scoped token. Both ids are published by Microsoft;
SETUP.md has the details and where to look them up.

Two things about this grant. It cannot be used by an account with MFA or an MFA-requiring
Conditional Access policy, which is why the account has to be excluded from both. And Microsoft
has ROPC on the way out, so at some point it will stop working.

So authentication sits behind one interface, `TokenProvider` in `src/auth/token-provider.ts`, with
exactly one method. `RopcTokenProvider` is one implementation of it. The Graph client, the tools
and the allowlist know nothing about how the token was obtained. Replacing ROPC with device code
means writing a second implementation and changing the one line in `src/index.ts` that constructs
it. Device code is the obvious successor: same client id, same scopes, one interactive sign-in per
refresh-token lifetime instead of a password on every cold start.

The access token and its refresh token are cached on disk between restarts. That file holds a live
credential, so it is written mode 0600 and its name is in `.gitignore`.

## The allowlist

`teams-mcp.config.json` lists the chats the server may touch:

```json
{
  "assistantDisplayName": "Assistant (AI)",
  "allowedChats": [
    { "id": "19:....@thread.v2", "label": "Pilot chat", "canPost": true },
    { "id": "19:....@thread.v2", "label": "Leadership", "canPost": false }
  ]
}
```

Every chat id entering a Graph call goes through the allowlist first. A chat that is not listed is
refused for both reading and posting, and the refusal happens before any network call. `canPost`
defaults to `false` when omitted, so a chat added carelessly gets read access only.

An empty list is a startup error rather than an open door, because an empty list is almost always
a misconfigured file.

This matters because the account's token is broad. It carries every delegated Teams scope the
first-party client id grants, which means the identity itself can reach every chat the account is
a member of. The allowlist is the only thing narrowing that down, so treat editing it as a
governance action, not a config tweak.

`list_chats` shows allowlisted chats only. It will not tell you about other chats the account is
in, which also means it is no help for finding a chat id in the first place. That job belongs to
`npm run discover-chats`, which a human runs and which prints every chat the account can see
(SETUP.md step 4 lists the other ways to find an id).

## The account: one identity per consuming project

Each project that installs this server gets its own dedicated cloud-only account. Not a person's
identity, not an admin account, and not shared with another project. Sharing one account across
projects would mean one allowlist edit in project A silently widening what project B's agent can
reach, and one leaked password burning every project at once.

The account needs:

- an Office 365 licence with the Teams service plan enabled (without it every `/chats` call fails
  with 403 "Failed to get license information for the user"; auth succeeds, the licence is the
  blocker, and no scope or permission change fixes it)
- MFA off, and exclusion from any Conditional Access policy that requires MFA
- a display name that says it is an AI, for example `Assistant (AI)`
- membership in each allowlisted chat, added by a human the normal way

The display name is not cosmetic. Anyone in the chat should be able to see at a glance that the
thing writing is an assistant. `npm run probe` compares the account's real `displayName` against
`TEAMS_MCP_DISPLAY_NAME` (default `Assistant (AI)`) and warns on a mismatch.

## Configuration

Everything sensitive comes from the environment, per installation. Nothing in this repo carries a
credential, an account name, or a tenant id, and nothing ever should.

| Variable | Required | Notes |
|---|---|---|
| `TEAMS_MCP_TENANT_ID` | yes | Entra ID tenant of the assistant account |
| `TEAMS_MCP_USERNAME` | yes | Assistant account UPN |
| `TEAMS_MCP_PASSWORD` | yes | |
| `TEAMS_MCP_CONFIG` | yes | Path to the allowlist config. The server will not start without one |
| `TEAMS_MCP_CLIENT_ID` | no | Defaults to the first-party Teams client id; see SETUP.md for why you usually want the Office one. Also the knob for Graph throttle isolation — see "Throttle budgets are per client id" below |
| `TEAMS_MCP_TOKEN_CACHE` | no | Defaults to `.token-cache.json` in the working directory. The members cache (see "@mentions" below) and the self-id cache (below) live next to it |
| `TEAMS_MCP_MEMBERS_TTL_SECONDS` | no | How long a chat's cached member list is trusted before a mention resolution refreshes it (bounds a COMPLETE roster's `fetchedAt` and a PARTIAL, traffic-harvested roster's `harvestedAt` alike — see "@mentions" below). `send_chat_file`'s permission grant judges freshness against `fetchedAt` alone, ignoring intervening traffic, and never uses a PARTIAL roster at all. Defaults to 24h (86400) |
| `TEAMS_MCP_SELF_ID` | no | Last-resort operator seed for the signed-in account's own AAD id (`resolveSelfId`), used internally so `send_chat_file` (row above) can exclude the assistant from its own read-access grant on the uploaded item. Normally unnecessary: the id is resolved once from `/me` and persisted next to the token cache with no TTL, so only the first process on a fresh install pays the live lookup. Must be GUID-shaped or it is ignored |
| `TEAMS_MCP_DOWNLOAD_DIR` | no | Where attachment downloads land (`get_chat_attachment`, `download_chat_attachments`, `teams-attachments`). Defaults to a temp directory |
| `TEAMS_MCP_UPLOAD_DIR` | no | OneDrive folder where `send_chat_file` parks uploads. Defaults to `ai-test` |
| `TEAMS_MCP_DISPLAY_NAME` | no | Overrides the expected display name for the probe |
| `TEAMS_INBOX_PATH` | no | Where the background inbox JSONL lands. Defaults to `~/.teams-assistant/inbox.jsonl`. **Set this to a distinct path for every additional server instance on the same host** — the poller's single-instance lock (`poller.lock`) and health file (`poller-health.json`) are keyed to this path, not to `TEAMS_MCP_CONFIG`, so two instances that both leave it unset silently share one lock and only one of them polls (see "Supervising the daemon") |
| `TEAMS_INBOX_POLL_SECONDS` | no | Inbox poll interval, default 30. Raising it frees per-mailbox Graph read budget for ad-hoc reads — see "Downloading attachments" |
| `TEAMS_INBOX_DISABLED` | no | Set to `1` to not run the background inbox poller at all |

### Throttle budgets are per client id

Microsoft Graph's throttle budget for the `/chats`, `/messages` and `/members` families is keyed
by application (client id) *and* signed-in user together. Two long-lived processes signed in as
the same account with the same `TEAMS_MCP_CLIENT_ID` (the default first-party Teams id, if neither
overrides it) SHARE one budget — a 429 either one earns closes the gate for both, and one daemon's
traffic can keep the other permanently throttled even though it made none of the offending calls
itself (verified live, 0.4.1: two daemons on the same default client id, one bystander). Giving
each long-lived instance its own `TEAMS_MCP_CLIENT_ID` (any other Microsoft first-party id — see
SETUP.md's client-id section) decouples their budgets. This is documentation only: the env
override already exists (`config.ts`'s `pick(env, ['TEAMS_MCP_CLIENT_ID'])`), nothing new to wire.

`env.example` lists the same variables with comments, and SETUP.md walks through filling them in
(including why the password wants single quotes and the paths want to be absolute). Put real
values in a gitignored `.env`, never in this repo and never in the consuming repo.

## Building and testing

```bash
npm ci
npm run lint       # oxlint
npm run test:run   # vitest
npm run build      # tsc -> dist/
npm pack           # versioned tarball of dist/ plus the example config
```

CI runs the same four gates on every push to main and publishes the `npm pack` tarball as a
build artifact. Releases are the tarballs of tagged versions; there is no npm registry publish
yet, which is the natural next step once a second consumer appears.

Tested without a licensed account: allowlist enforcement through a real MCP client, the token
provider (caching, expiry skew, refresh-then-password fallback, concurrent callers, no password in
error text), config and allowlist parsing, the 0600 token cache, HTML-to-text, Graph message
mapping, watermark diffing, and the Graph client's handling of the licence 403. The end-to-end
cases that need Graph to return an actual chat are listed as pending in
`src/graph/live.awaiting-licence.test.ts`. They are deliberately not written against a mocked
Graph, since a mock cannot answer the one question they exist for: whether Graph accepts these
calls from a real assistant identity.

## Installing into a project

SETUP.md covers the whole path, from requesting the account to a passing smoke check. The short
version: build this repo, fill in `.env` and `teams-mcp.config.json` here, and run
`npm run install-local` — it prints the exact `claude mcp add` command (and `.mcp.json`
equivalent) for your checkout.

## Licensing

This project is licensed under the [Elastic License 2.0](LICENSE). In plain
language: you are free to use it, copy it, modify it, and adapt it to make it
work for your own organization — including commercially, inside any company.
The one thing you may not do is sell it or offer it to others as a hosted or
managed product. We share it with the community for free and would like it to
stay that way.
