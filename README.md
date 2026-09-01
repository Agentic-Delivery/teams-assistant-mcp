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

The server speaks MCP over stdio and exposes fourteen tools:

| Tool | What it does |
|---|---|
| `list_chats` | The allowlisted chats, annotated with whether the account can actually see each one |
| `read_chat_messages` | Messages from one chat, oldest first, with a watermark for the next call |
| `send_chat_message` | Posts to a chat whose allowlist entry has `canPost: true`; `format: 'text'` (default) escapes and renders, `format: 'html'` posts raw HTML verbatim; optional `mentions` — see "@mentions" below |
| `send_chat_image` | Posts a PNG/JPEG that renders inline, from a local path or base64 bytes |
| `send_chat_file` | Uploads a local file to the account's OneDrive (`TEAMS_MCP_UPLOAD_DIR`, default `ai-test`) and shares it into the chat |
| `reply_chat_message` | Posts a quoted reply to a specific message — chats have no reply threads, so this is the quote card the Teams UI produces; optional `mentions` |
| `edit_chat_message` | Replaces the text of a message this account sent (Graph refuses anyone else's); same `format` and `mentions` options as `send_chat_message` |
| `react_to_chat_message` | Puts an emoji reaction on a message — the receipt gesture for "seen, being handled" |
| `delete_chat_message` | Soft-deletes a message this account sent — the reversible kind; no hard delete offered |
| `get_chat_attachment` | Downloads one attachment to a local file and returns the path — shared files, and pasted images which appear as `inline-image-N` |
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
name still unresolved after that gets the usual clear error. A 429 on that one refresh call is
never swallowed into a false "no such member": it surfaces as its own throttled error naming
Graph's Retry-After.

## Retry-After

Any Graph 429 in the send/reply/edit path never auto-retries — that has always been the rule for
writes (`GraphClient.post`/`ReliableTeamsChats`). The CLIs (`teams-post`, `teams-reply`,
`teams-edit`) now also name the wait in their error output whenever Graph named one:
`Too many requests (throttled, retry after 62s)`, so an operator staring at a failed send does not
blind-retry into a worse throttle.

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

A failing poll appends `{"error": "...", "at": "..."}` to the same file. That line is the
difference between "the chats are quiet" and "auth is dead" — a watcher must never have to
guess which silence it is looking at. An identical failure repeating poll after poll is written
once, not once per poll. When *everything* fails (auth death, network gone) the interval backs
off, doubling to a 10-minute cap and snapping back on recovery; a single failing chat — usually
one the account has not been added to yet — is surfaced but does not slow the healthy chats
down. The poller never crashes the server; every poll is fully caught.

The recommended consumption pattern: arm a file watcher (Claude Code's `Monitor`, `tail -F`,
inotify) on the inbox path at session start and react per line. Do not poll the tools for new
messages any more.

Two knobs: `TEAMS_INBOX_PATH` moves the inbox (the sidecar follows it), and
`TEAMS_INBOX_DISABLED=1` switches the poller off entirely for consumers that only post.

## The standalone CLIs

Seven small commands ship beside the server for scripts, cron jobs and background monitors that
need Teams without a running MCP session: `teams-post <chatId> [--html] [--mention "Name"]...`
(text on stdin), `teams-reply <chatId> <messageId> [--mention "Name"]...` (text on stdin),
`teams-edit <chatId> <messageId> [--html] [--mention "Name"]...` (new text on stdin), `teams-react
<chatId> <messageId> <emoji>`, `teams-read <chatId> [--limit N] [--since ISO]`, `teams-pin
<chatId> <messageId>` and `teams-unpin <chatId> <messageId>`. Same allowlist, same auth, same
code paths as the server tools — including the send reliability below. `--html` on
`teams-post`/`teams-edit` posts stdin as raw Teams-subset HTML, verbatim — the caller is
responsible for entity-escaping their own `<`, `>`, `&`; see the `teams-styling` plugin for the
verified vocabulary. `--mention
"Name"` (repeatable) @mentions that person, same resolution and placement rules as the
`mentions` tool parameter — see "@mentions" above.

Their output contract exists because of a real incident (2026-08-24): an ad-hoc wrapper's
caller grepped for a success token the wrapper never printed, read eleven successful posts as
eleven throttles, and re-posted a broadcast ten times. So: success is exactly one JSON line on
stdout and exit 0; failure is prose on stderr and a non-zero exit (2 usage, 3 allowlist,
1 anything else). **Branch on the exit code, never on output text.**

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
| `TEAMS_MCP_TOKEN_CACHE` | no | Defaults to `.token-cache.json` in the working directory. The members cache (see "@mentions" below) lives next to it |
| `TEAMS_MCP_MEMBERS_TTL_SECONDS` | no | How long a chat's cached member list is trusted before a mention resolution refreshes it. Defaults to 24h (86400) |
| `TEAMS_MCP_DOWNLOAD_DIR` | no | Where `get_chat_attachment` writes. Defaults to a temp directory |
| `TEAMS_MCP_UPLOAD_DIR` | no | OneDrive folder where `send_chat_file` parks uploads. Defaults to `ai-test` |
| `TEAMS_MCP_DISPLAY_NAME` | no | Overrides the expected display name for the probe |
| `TEAMS_INBOX_PATH` | no | Where the background inbox JSONL lands. Defaults to `~/.teams-assistant/inbox.jsonl` |
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
