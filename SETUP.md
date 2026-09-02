# Setting this up for your project

This is the zero-to-working manual. At the end a Claude Code agent in your project can read and
post in a fixed set of Teams chats, signed in as a dedicated account. Most steps are minutes; the
account request in step 1 is the long pole, so start it first and do the rest while you wait.

You need Node 20.12 or newer (the scripts use `util.parseEnv` and the server is launched with
`node --env-file`), git, and the `claude` CLI in the project that gets the tools.

## 1. Request a dedicated account — one per project, always

The server signs in with a plain username and password. That account is the identity your agent
has in Teams, and getting it right is the only step you cannot fix later without asking IT twice.

Never reuse a person's account, and never share one assistant account between projects. Two
projects means two accounts, two `.env` files, two isolated identities and sets of rooms. With a
shared account, an allowlist edit made for project A silently widens what project B's agent can
reach, and one leaked password burns every project at once.

Ask the account/licensing owner for:

- A member account (not a guest) **with an Office 365 licence that includes the Teams service
  plan**. Without the licence, sign-in succeeds and then every chat call fails with 403 "Failed
  to get license information for the user". No scope or permission change fixes that; only the
  licence does.
- **No MFA**, and exclusion from any Conditional Access policy that blocks legacy/password
  sign-in. The mechanism is ROPC, the OAuth password grant: the server exchanges username and
  password for a token directly, with no browser step where a second factor could happen. MFA
  breaks it by design.
- A display name that says what it is, for example `Assistant (AI)`. Everything the agent
  posts appears under this name. There is no per-message override: Graph silently ignores
  attempts to set a different sender on delegated posts (verified 2026-08-20). People in the
  chat must be able to see at a glance that they are talking to an AI.

Once the account exists, a human adds it to each chat it should work in, the normal way. The
allowlist you configure later only narrows what the account could reach; it cannot grant access
to a chat the account is not a member of.

## 2. Clone and build

```bash
git clone <this repo> teams-assistant-mcp
cd teams-assistant-mcp
npm ci
npm run build
```

`npm run test:run` should be green if you want to check the checkout (one file reports as
skipped — those are live-Graph cases waiting on a licensed account, not a problem with yours).

## 3. Fill in .env

```bash
cp env.example .env
```

The `.env` stays in this repo's root (it is gitignored) and is the one place real credentials
live. Per variable:

**`TEAMS_MCP_TENANT_ID`** — the Entra ID tenant the account lives in.
`az account show --query tenantId -o tsv` prints it if you are logged in with the Azure CLI;
otherwise it is on the overview page of [entra.microsoft.com](https://entra.microsoft.com)
(Identity → Overview → Tenant ID).

**`TEAMS_MCP_USERNAME`** — the account's UPN, e.g. `teams-assistant@yourtenant.com`.

**`TEAMS_MCP_PASSWORD`** — single-quote it:

```bash
TEAMS_MCP_PASSWORD='the-actual-password'
```

Node's env-file parser strips the quote envelope and takes the content literally, and a shell
sourcing the file treats single quotes the same way, so `$`, `#`, spaces and double quotes inside
the password all survive. The classic trap is a password that itself contains a quote: a quoted
value ends at the first matching quote inside it, silently truncating the password. If yours
contains a single quote, wrap it in double quotes instead; if it contains both kinds, request a
password without quotes.

**`TEAMS_MCP_CLIENT_ID`** — which of Microsoft's own client ids the sign-in claims to be.
Neither needs an app registration or admin consent. Left unset, the server uses the Teams
first-party id (the `TEAMS_FIRST_PARTY_CLIENT_ID` constant in `src/config.ts`); that one
authenticates everywhere, but on some tenants the token it returns carries no chat scopes at all
and every `/chats` call fails 403 even with a licensed account. The Microsoft Office first-party
id worked fully on the tenant this was developed against and is the recommended value. Microsoft
publishes the application ids of its first-party apps — look up "Microsoft Office" in that list
and set:

```bash
TEAMS_MCP_CLIENT_ID=YOUR_CLIENT_ID   # the Microsoft Office first-party application id
```

If chat calls fail 403 with a licensed account, run `npm run probe` and look at the `scopes`
line of the token before suspecting anything else.

`TEAMS_MCP_CLIENT_ID` is also the throttle-isolation knob: Graph's throttle budget for the chat
endpoints is keyed by client id + signed-in user together, so two long-lived instances signed in
as the same account with the same client id share one budget — a 429 either one earns can starve
the other. Running more than one long-lived instance against the same account (a second server, a
standalone polling daemon)? Give each its own `TEAMS_MCP_CLIENT_ID` (any other Microsoft
first-party id from the same published list) so their throttle budgets are separate. See the
README's "Throttle budgets are per client id" section.

**`TEAMS_MCP_TOKEN_CACHE` and the member cache** — the per-chat @mention member cache
(`TEAMS_MCP_MEMBERS_TTL_SECONDS`, default 24h) is written next to whatever `TEAMS_MCP_TOKEN_CACHE`
resolves to, as `.members-cache.json` in the same directory. Both files hold per-instance state, so
this is another reason two long-lived instances for the same account want separate working
directories (and separate `TEAMS_MCP_TOKEN_CACHE` paths), not just separate client ids.

**`TEAMS_MCP_CONFIG`** — absolute path to the allowlist file (step 5). The server refuses to
start without it. Absolute, because the agent starts the server from the consuming project's
directory, so a relative path resolves against the wrong repo.

**`TEAMS_MCP_TOKEN_CACHE`** — absolute path for the cached refresh token, e.g.
`<repo>/.token-cache.json`. The file holds a live credential (written mode 0600) and that name is
gitignored here; leave the variable unset and the default lands in the agent's working directory
instead, i.e. inside your project.

**`TEAMS_INBOX_POLL_SECONDS`** — inbox poll interval, default 30. The poller is the main
consumer of the account's per-mailbox Graph read budget; raise this if the same account also
serves ad-hoc reads and attachment downloads (README "Downloading attachments" has the
measured story).

**`TEAMS_MCP_DOWNLOAD_DIR`** — where attachment downloads land (`get_chat_attachment`,
`download_chat_attachments`, `teams-attachments`). Defaults to a directory
under the OS tmpdir, which may be cleaned between sessions.

**`TEAMS_MCP_UPLOAD_DIR`** — OneDrive folder (under the account's drive root) where
`send_chat_file` parks uploads before sharing them. Defaults to `ai-test`.

**`TEAMS_MCP_DISPLAY_NAME`** — what `npm run probe` expects the account's real display name to
be. The probe warns on a mismatch; it is the check that the "says it is an AI" rule from step 1
actually holds.

## 4. Find your chat ids

A Graph chat id looks like `19:...@thread.v2` (group and meeting chats) or
`19:..._...@unq.gbl.spaces` (one-to-one). Three ways to get one, in the order that works:

**Copy link in Teams (primary).** Right-click the chat in the Teams chat list and choose *Copy
link*. Paste it anywhere. The id is the segment between `/l/chat/` and `/conversations`,
verbatim:

```
https://teams.microsoft.com/l/chat/19%3A0123456789abcdef0123456789abcdef%40thread.v2/conversations?tenantId=...
                                   └──────────────────── the chat id ────────────────┘
```

When it arrives URL-encoded like above, decode `%3A` to `:` and `%40` to `@` — nothing else
changes. Example ids (all fake):

- group chat: `19:0123456789abcdef0123456789abcdef@thread.v2`
- meeting chat: `19:meeting_MDAwMDAwMDAtMDAwMC0wMDAwLTAwMDAtMDAwMDAwMDAwMDAw@thread.v2`
  (meetings keep their chat; the id works like any other)
- one-to-one: `19:aaaaaaaa-1111-2222-3333-444444444444_bbbbbbbb-5555-6666-7777-888888888888@unq.gbl.spaces`
  (the two GUIDs are the two members)

**`npm run discover-chats`.** Signs in with the `.env` credentials and lists every chat the
account can see, with id, type, topic and members. Needs step 1 complete: the account must have
its licence and be a member of at least one chat. This is also the sanity check that the
credentials work at all.

**The Teams web URL bar (fallback).** Open [teams.microsoft.com](https://teams.microsoft.com) in
a browser and click into the chat; the address bar contains the chat id, URL-encoded the same way
as the copy-link form.

## 5. Write the allowlist

```bash
cp teams-mcp.config.example.json teams-mcp.config.json
```

One entry per chat the agent may touch:

```json
{
  "assistantDisplayName": "Assistant (AI)",
  "allowedChats": [
    { "id": "19:0123456789abcdef0123456789abcdef@thread.v2", "label": "Pilot chat", "canPost": true },
    { "id": "19:fedcba9876543210fedcba9876543210@thread.v2", "label": "Leadership — read only", "canPost": false }
  ]
}
```

Decide `canPost` per chat, deliberately. A chat the agent should only observe gets `false`; when
the field is omitted it defaults to `false`, so a carelessly added chat is read-only. Set
`assistantDisplayName` to the account's actual display name: it is what the server tells the
agent its own name in the chat is, and lying here defeats the point of the honest name from
step 1. (The probe's mismatch warning checks `TEAMS_MCP_DISPLAY_NAME` from `.env`, not this
field, so keep the two in sync with the account.)

An empty `allowedChats` is a startup error, not an open door. The allowlist is the only thing
narrowing down what the account's token could reach, so treat edits to this file as a governance
action.

## 6. Connect it to a Claude Code agent

```bash
npm run install-local
```

builds and prints the registration command for your checkout, something like:

```bash
claude mcp add teams-assistant -- node --env-file=/home/you/repo/teams-assistant-mcp/.env /home/you/repo/teams-assistant-mcp/dist/index.js
```

Run that inside the project that should get the tools. One thing to understand about this line:
`dist/index.js` reads `process.env` and nothing else — the server has no dotenv and never opens
`.env` itself. Node's `--env-file` flag is what loads the file, which is why the flag (with an
absolute path) has to be part of the command. It is also why paths *inside* `.env` must be
absolute: the agent starts the server from your project's directory, not from this repo.

The equivalent for a project-scoped `.mcp.json` (install-local prints this too):

```json
{
  "mcpServers": {
    "teams-assistant": {
      "command": "node",
      "args": [
        "--env-file=/home/you/repo/teams-assistant-mcp/.env",
        "/home/you/repo/teams-assistant-mcp/dist/index.js"
      ]
    }
  }
}
```

The paths are machine-specific. If your team commits `.mcp.json`, agree on a checkout location,
or let each teammate register locally with their own install-local output.

Smoke checklist:

1. Start a new Claude Code session in the project; `/mcp` shows `teams-assistant` connected.
   On a good start the server logs `teams-assistant-mcp ready: N allowed chat(s), auth=ropc` to
   stderr; on failure the reason lands there too, with a missing config file and an empty
   allowlist as the usual suspects.
2. Ask the agent to call `list_chats`. Your allowlisted chats come back, each with
   `visibleToAccount: true`. A `false` there means the account is not a member of that chat —
   the allowlist cannot fix that, only adding the account to the chat can.
3. Optionally, have it `send_chat_message` to a `canPost` chat and watch the message arrive
   under the assistant's name.

That's it. Done.

## 7. How it behaves day to day

Nothing pushes. Teams does not notify this server; the agent finds new messages by calling
`poll_chats` with the watermarks object the previous call returned. Every 15–30 seconds is a
reasonable cadence — faster mostly burns Graph calls, much slower makes the chat feel dead. The
usual pattern is a background sentinel: a cheap loop (a subagent, a cron-ish task, anything) calls
`poll_chats` on a timer and carries the watermarks between rounds. When something new arrives it
wakes the main agent with the messages instead of the agent sitting in the poll loop itself. The
watermarks object is the entire state; lose it and you re-read recent history, which is annoying
but harmless.

Everything the agent posts appears under the account's display name. There is no way to vary the
sender per message.

Edited messages do not come back through the watermark: an edit keeps the original
`createdDateTime`, so a poll will not see it again. The only trace is `lastModifiedDateTime` on
the message when you re-read the chat. If catching edits matters, re-read instead of relying on
the poll.

Deleted messages read back with `isDeleted: true` and an emptied body. The agent's own
`delete_chat_message` is the soft kind (Teams shows the "This message was deleted" stub) and
`undo_delete_chat_message` puts the original back; there is no hard delete here. Both act on
the account's own messages only unless told `force` — the typical use is withdrawing something
posted in the wrong chat, not tidying up after other people.

Replies are quote-cards. Chats have no real threads (that is a channels feature), so
`reply_chat_message` posts a normal message carrying a quote of the original at the bottom of the
chat. Use it when a bare message would lose which question it answers.

The allowlist is the hard safety boundary. The account's token can reach every chat the account
is in; the server refuses anything not listed, before any network call. Read-only chats stay
read-only because their entry says `canPost: false`, not because anyone remembers to be careful.

## 8. Tool reference

The seventeen tools, as registered in `src/server.ts`. All results are JSON text; errors
(including allowlist refusals) come back as readable text, not transport failures.

| Tool | What it does | Input |
|---|---|---|
| `list_chats` | The allowlist intersected with what the account can see; `visibleToAccount: false` flags missing membership | none |
| `read_chat_messages` | Messages from one chat, oldest first, with a watermark for the next call | `chatId`, `since?` (ISO watermark, exclusive), `limit?` (1–200, default 50) |
| `send_chat_message` | Posts to a chat with `canPost: true` | `chatId`, `text`, `format?` (`'text'` default, escapes and renders; `'html'` posts raw Teams-subset HTML verbatim — caller escapes their own `<`, `>`, `&`), `mentions?` (display names to @mention — see below) |
| `reply_chat_message` | Posts a quote-card reply to a specific message | `chatId`, `replyToMessageId`, `text`, `mentions?` |
| `edit_chat_message` | Replaces the text of a message this account sent; Teams shows "Edited" | `chatId`, `messageId`, `newText`, `format?`, `mentions?` (same as `send_chat_message`) |
| `react_to_chat_message` | Puts an emoji reaction on a message in an allowlisted chat | `chatId`, `messageId`, `emoji` |
| `delete_chat_message` | Soft-deletes a message this account sent (restorable); refuses somebody else's message unless `force` | `chatId`, `messageId`, `force?` |
| `undo_delete_chat_message` | Restores a soft-deleted message this account sent; same ownership rule | `chatId`, `messageId`, `force?` |
| `send_chat_image` | Posts a PNG/JPEG that renders inline | `chatId`, `path?` or `base64?` (exactly one), `mime?` (required with base64), `text?` |
| `send_chat_file` | Uploads to the account's OneDrive and shares into the chat as a file card, granting every other chat member read access on the uploaded item | `chatId`, `path`, `text?` |
| `get_chat_attachment` | Downloads one attachment to `TEAMS_MCP_DOWNLOAD_DIR` and returns the path; pasted images appear as `inline-image-N` | `chatId`, `messageId`, `attachmentId?` (default: first) |
| `list_chat_attachments` | One message's attachment metadata — id, name, contentType, downloadable — nothing transferred | `chatId`, `messageId` |
| `download_chat_attachments` | Downloads everything downloadable on one message (quote cards skipped); names sanitized, collisions suffixed, never overwrites | `chatId`, `messageId`, `nameFilter?` (case-insensitive substring), `outputDir?` |
| `poll_chats` | Reads every allowlisted chat in one call, carrying a watermark per chat | `watermarks?` (chatId → ISO watermark), `limit?` |
| `pin_chat_message` | Pins a message — REPLACES whatever was pinned before it (a chat effectively holds one pin); returns the resulting pinned list | `chatId`, `messageId` |
| `unpin_chat_message` | Unpins a message; refuses if it is not the one currently pinned | `chatId`, `messageId` |
| `list_pinned_messages` | Lists what is currently pinned (at most one entry in practice), with a plain-text preview | `chatId` |

Editing only works on the account's own messages; that is Graph's rule for delegated calls, and
the server passes Graph's refusal through verbatim rather than pre-checking it. Deleting and
restoring are pre-checked on purpose: the message is fetched and its author compared with the
signed-in account before anything is sent, and anyone else's message is refused — with the
author named — unless `force` is true. Pass `force` only when a human has asked for it; see
README's "Withdrawing a message" for the permission (`Chat.ReadWrite`) and the throttle story.

`mentions` (`send_chat_message`/`reply_chat_message`/`edit_chat_message`) is a list of display
names to actually NOTIFY, not just reference — resolved case-insensitively as an unambiguous
substring against the chat's member list ("Shiv" matches "Garg, Shivankit"; zero or multiple
matches is refused, never silently dropped). With `format: 'text'` (default), each mention's
name must occur in the message text — that occurrence becomes the notifying tag. With `format:
'html'`, place a literal `@{Name}` token at each spot to mention instead; every declared mention
needs a matching token and vice versa. See the `teams-styling` skill's "Mentions" section for
when to tag someone versus just naming them.

## The standalone CLIs

Beside the server, `npm run build` produces ten commands under `dist/cli/` (also exposed as
package bins): `teams-post [--html] [--mention "Name"]...`, `teams-reply [--mention "Name"]...`,
`teams-edit [--html] [--mention "Name"]...`, `teams-react`, `teams-read`, `teams-pin`,
`teams-unpin`, `teams-send-file <chatId> <path> [more paths...] [--caption "text"]`,
`teams-attachments <chatId> <messageId> [--list] [--name <filter>] [--out <dir>]` (download all,
or `--list` for metadata only — see README's "Downloading attachments") and `teams-delete
<chatId> <messageId> [--undo] [--force]` (soft-delete one of the account's own messages, `--undo`
to restore it, `--force` to skip the own-message check — see README's "Withdrawing a message").
Same env
vars, same allowlist, same send-reliability code paths as the server tools. `--html` on
`teams-post`/`teams-edit` posts stdin as raw HTML verbatim, same contract as `format: 'html'`
above; `--mention "Name"` (repeatable) works the same as the `mentions` tool parameter.
`teams-send-file` sends one or more local files (`--caption`, optional, shown on the first file
only), streaming one JSON line to stdout as EACH file lands rather than buffering until the whole
batch finishes — see README's "The standalone CLIs" for why a mid-batch failure still needs the
earlier lines visible. Success is exactly one JSON line on stdout and exit 0; failure is stderr
plus a non-zero exit (2 usage, 3 allowlist, 1 anything else). Branch on the exit code, never on
output text — see README's "The standalone CLIs" for the incident that made this a rule.
