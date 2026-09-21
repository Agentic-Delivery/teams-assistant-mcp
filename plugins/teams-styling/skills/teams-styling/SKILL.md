---
name: teams-styling
description: "Styling doctrine for agent-posted Teams messages: when to style and when not to, the empirically verified HTML vocabulary Teams renders, known quirks, and the live status-board pattern for agent-progress transparency. Read it BEFORE composing the message body — do not skip because the message 'looks short'; more than two sentences is not short. Triggers on: posting, replying to, or editing ANY Teams chat message; 'post.mjs', 'reply.mjs', 'edit.mjs', 'teams-post', 'teams-reply', 'teams-edit'; the flags '--html' and '--text'; 'send_chat_message', 'reply_chat_message', 'edit_chat_message'; any message longer than two sentences; 'Teams message', 'chat message', 'status board', 'findings table', 'status update', 'report to the channel', 'alert'."
---

# Teams message styling

Companion skill to `teams-assistant-mcp`. Everything in the vocabulary below was verified
against the real Teams client (desktop, dark theme) by screenshot review — not inferred from
Microsoft documentation.

**Capability note**: styled output requires posting Graph `contentType: "html"` content
*unescaped*. The server's standard send path deliberately escapes everything (plain text
in, plain text out); the raw-HTML path this skill assumes is `format: 'html'` on
`send_chat_message`/`edit_chat_message` (MCP tools) and `--html` on `teams-post`/`teams-edit`
(standalone CLIs) — shipped since teams-assistant-mcp 0.3.0. On that path YOU own
entity-escaping `<`, `>`, `&` inside your own content (quirk 3 below); the server posts it
verbatim. Never post markdown syntax into Teams (it renders as literal `**asterisks**`).
Threaded replies (`reply_chat_message` `format: 'html'` / `teams-reply --html`) carry the same
verbatim contract since reply HTML parity — style a reply exactly as you would a new message.

## When to style — and when not to

Styling is for making structured content scannable, not for decoration. Default is plain
text; reach for styling when the content has structure a reader must scan:

| Content | Format |
|---|---|
| Short conversational reply, answer to a question (1–2 sentences) | Plain text — no styling |
| Medium answer to a person (3–8 sentences, usually with an ask) | HTML: one `<b>` lead line, `<div>&nbsp;</div>` air between ideas, bullets for parallel items, the question/ask on its own last line — never a plain-text paragraph block |
| Comparison across items with SHORT cells (ids, labels, verdicts, numbers) | `<table>` — never plain-text pipe walls |
| Findings with prose explanations | Headed sections, NOT a table: one `<b>` heading line per item, `<ul>` bullets under it, a bold `→ Fix:`/`→ Action:` closing bullet |
| Long report | Summary message with a table + link to the full document; not the whole report inline |
| Progress on multi-agent or long-running work | Live status board (pattern below) |
| Warning / blocker / decision needed | One `<b>` lead line + ⚠️/🔴 emoji; mention the human owner by name |
| Code, identifiers, config values | `<code>` inline, `<pre>` for blocks |
| Anything with a pass/fail dimension | Emoji + bold as primary signal; color only as reinforcement (never color alone — color-blind readers, theme drift) |

Restraint rules: at most one `<h2>` title per message (h1 is oversized in chat); style the
signal, not every word; a message that is all bold has no bold. If in doubt, plain text.
More than two sentences ⇒ not plain text. Plain text is for one- or two-sentence replies only.

## Mentions

teams-assistant-mcp 0.4.0 added real, notifying @mentions — `mentions: string[]` on
`send_chat_message`/`reply_chat_message`/`edit_chat_message` (`--mention "Name"`, repeatable, on
the CLIs). A plain `@Name` typed into the text or an `<at>` tag with no matching Graph `mentions`
array entry is just decoration — it does NOT ring anyone's bell. Getting someone's actual
attention requires this feature, not typing their name.

**WHEN to tag** — @-mention a person whenever:
- they are named as a decision owner or action owner ("Mika owns the migration"),
- the message asks them something directly ("Johan, does this match what you expected?"),
- they specifically need to be notified now, not just find the message later.

**WHEN NOT to tag** — do NOT mention someone every time their name comes up in the narrative
("as Mika mentioned yesterday...", "this affects Bob's team too"). A name appearing in prose is
not the same as that person needing to act. Tag at most the people who actually need to DO
something with the message — notification fatigue is real: an inbox full of pings for messages
that needed no response trains people to stop reading pings at all, which defeats the one thing
a mention is for. One clear action-owner tag beats five drive-by name-checks.

**The placeholder contract for `format: 'html'`**: since raw HTML is posted verbatim (no
escaping, no rewriting), the server cannot infer where in your markup a mention belongs — you
mark the spot yourself with a literal `@{Name}` token (e.g. `@{Mika}`), matched case-insensitively
against the `mentions` list you also pass. The server replaces every occurrence of each token
with the real `<at id="N">...</at>` tag and the id it references in the parallel `mentions`
request field — place `@{Mika}` more than once if you genuinely mean the same person at several
spots, they all get the same id. Every name you declare in `mentions` needs AT LEAST ONE
`@{Name}` token somewhere in the html, and every token needs a matching declared name — an
orphaned mention (declared but never placed) or an orphaned token (placed but never declared,
usually a typo) is refused rather than silently posted as literal `@{Mika}` text with no
notification behind it. `format: 'text'` needs no placeholder — write the person's plain name
where you mean to mention them (every whole-word occurrence gets tagged, but never a name that's
only part of a URL or another word — see mentions.ts if you need the exact matching rules) and
the server finds it for you.

## Verified rendering vocabulary — all of these WORK

| Construct | Notes |
|---|---|
| `<b>` `<i>` `<u>` `<s>` and combinations | render exactly as expected |
| `<code>` | inline code gets a bordered monospace box |
| `<pre>` | monospace block with border, leading indentation preserved (see quirk 1) |
| `style="color:#..."` inline CSS | honored — semantic color works (red=fail, green=pass) |
| `<ul>` `<ol>` | proper bullets/numbering; inline styles inside items fine |
| `<blockquote>` | rendered as a bordered callout box |
| `<h1>` `<h2>` `<h3>` | real size hierarchy |
| `<a href="...">` | named hyperlinks work |
| `<table border="1" style="border-collapse:collapse;">` with per-cell `style="padding:4px 8px;"` | full borders, bold `<th>` header row, padding respected |
| emoji (✅ ⏳ 🔴 ⚠️ ⚙️) | render in color |
| Graph PATCH message edit | updates in place; original timestamp kept; small grey "Edited" label appears above the message |

## Known quirks

1. **Runs of multiple spaces mid-line in `<pre>`** render a stray visible glyph (¬-like) —
   likely NBSP substitution. Line-leading indentation is safe; avoid mid-line space runs.
2. **🏁 (checkered flag)** renders as a monochrome glyph, not a colored emoji. Prefer ✅/🎉
   for completion. Verify uncommon emoji in a test chat before relying on them.
3. **Manual escaping is on you** on the raw-HTML path: `<`, `>`, `&` inside content (code
   snippets, generics, XML) must be entity-escaped (`&lt;` `&gt;` `&amp;`) or they are
   swallowed as markup.
4. **Tables reflow with the reader's pane width** — there is no fixed layout. A table that
   looks fine full-screen collapses into one-word-per-line columns in the default narrow
   chat pane (screenshot-verified both ways). You cannot control the reader's window, so
   design for the narrow case: cells hold at most one short sentence; multi-sentence prose
   means you wanted headed sections, not a table.
5. **Inline `<code>` boxes carry heavy padding** and visibly break line rhythm when
   frequent — several per sentence turns prose into confetti, and inside narrow table
   cells they force ugly wraps. Reserve them for a few short identifiers per message;
   never put a long file path in a code box inside a table cell (it gets clipped) — long
   paths go in plain text or on their own line.
6. **One pin per chat.** Pinning a second message (Graph `POST /chats/{id}/pinnedMessages`)
   silently REPLACES the first while the API reports success — verify with a GET of the
   collection, never trust the POST. Consequence: a standing fixture must be ONE combined
   message (see the dashboard pattern below), not one pin per concern.
7. **`<p>` and list elements stack large margins** — fine for prose, but they wreck dense
   dashboards (headers float away from their lists, tails glue together). For dashboard-like
   messages use plain `<div>` per line with a `<div>&nbsp;</div>` spacer between sections;
   numbered items as literal "1. " text in divs, not `<ol>`. Screenshot-verified both ways.

## The live status-board pattern (agent-progress transparency)

Purpose: anyone glancing at the chat sees what the agent team is doing *right now*,
without scroll spam.

- On dispatching more than one agent, or any work with an ETA over ~15 minutes: post ONE
  board message — `<b>⚙️ Agent team status</b> (updated HH:MM)` plus a `<ul>` with one line
  per agent: status emoji (⏳ running / ✅ done / 🔴 failed), agent role, task, ETA or
  duration.
- Update by EDITING that same message (Graph PATCH), never by re-posting.
- **Pace the edits**: update on meaningful state change (agent finished/failed, ETA
  slipped), not on every tick. Leave each version visible long enough for a human to read
  it — minimum ~1 minute between edits. This is also what keeps you clear of Graph rate
  limits.
- The final edit marks everything ✅ and names where the report landed; the detailed report
  is its own separate message, styled per the vocabulary above.
- Boards carry progress, never decisions: anything needing a human decision is a normal
  message that names its owner.

**Standing-dashboard variant** (recommended for a channel the team lives in): instead of a
board per mission, keep ONE permanent pinned message — status section on top, a backlog/plan
mirror below (see verified-delivery's ways-of-working for the ownership doctrine) — edited
in place forever, div-layout per quirk 7, single pin per quirk 6. When nothing runs, the
status section must SAY so with a timestamp ("🟢 idle — last mission <x> completed <when>");
a board frozen on ⏳ reads as broken. Where a one-click full view of the master document
helps, chats also accept a website tab via Graph:
`POST /chats/{id}/tabs` with `"teamsApp@odata.bind": ".../appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"`
and `configuration.contentUrl`/`websiteUrl` pointing at the document's web view (verified
working with an Azure DevOps file view; the target must tolerate iframe rendering).

## Wiring it into a project

A skill teaches *how*; it does not make an agent reach for it. The consuming project's
conduct/profile must carry the mandate:

> All formatted channel output (tables, reports, status boards, alerts) follows the
> `teams-styling` skill; agents load it before composing any such message.

**That line alone was measured insufficient.** Earlier versions of this section claimed one line
is enough; an audit of 21 sessions in the guidewire factory (2026-09-21) falsified it. Of 188
sends that this skill's own table says must NOT be plain text, 58 went out styled — 31 %
compliance — with the mandate line in force the whole period, reinforced by an auto-loaded
project memory rule and two prior verbal corrections from the owner. Over-styling was 0 of 12:
the failure is pure omission, not disagreement. Split by whether this skill had loaded earlier
in the session, compliance was 13 % before a load and 63 % after (n=373) — the rule works while
it is in context and is absent otherwise. The mandate line is necessary and nowhere near
sufficient; deterministic enforcement is required alongside it, at two points:

- **A send-path guard in the tool** — `teams-assistant-mcp` refuses a multi-sentence body sent
  without `--html`/`format: 'html'` unless the caller passes an explicit `--text` override, from
  the version that ships the guard onward. This is the one enforcement that reaches every
  consumer automatically and holds regardless of model, session or harness.
- **A `PreToolUse` deny hook on `Bash`** — matches `(post|reply|edit)\.mjs` and the
  `teams-post`/`teams-reply`/`teams-edit` CLI names, denies when neither `--html` nor `--text` is
  on the send's own line and the body reads as structured text, and names this skill in the
  denial text together with the correct argument order (the chat id is positional-first, so
  `post.mjs --html <chat>` parses `--html` as the chat id). This is the only fix that carries the
  rule into context at the failing instant. **From 0.2.0 this plugin ships that hook** —
  `hooks/enforce-teams-styling.sh`, registered by `hooks/hooks.json`, carrying the same
  classifier as the CLI guard so it never denies what the tool would allow. No per-project hook
  installation is needed any more.

With the mandate line but neither enforcement point, expect agents to fall back to plain-text
walls roughly two times in three.

**How the shipped hook takes effect.** Plugin hooks are registered **when a session starts** (or
on `/reload-plugins`) — never live. Installing or upgrading to 0.2.0 mid-session does nothing to
that session; the hook fires from the next one. Plugin hooks **merge** with any project or user
`settings.json` hooks rather than replacing them: every matching hook runs, and a single deny
from any of them blocks the call. A project that hand-installed the earlier copy of this hook is
therefore safe running both — it should remove its local copy and the matching `settings.json`
entry once it has confirmed in a fresh session that the plugin hook fires, so there is one source
of truth. The hook needs `jq` on `PATH`, and its own test battery lives beside it at
`hooks/tests/hook-teams-styling-cases.sh` (19 cases; run it from the installed cache copy to
verify an upgrade).

**Where that line lives — and nothing of this package beside it.** No package artifact goes
into a customer's repository: this skill is read from the plugin cache, and everything a
project writes against it — the mandate line above, a project overlay of this skill, the
project's own channel conventions (board wording, pinned-message ids, the mention roster),
the settings that register this plugin — lives in the consumer's `.claude/` directory, which
is the clone of the provider's private per-project repository and is never tracked by the
consumer repository. The line names a plugin, so it is method, and it goes in the private
`.claude/CLAUDE.md` or the profile, not in a customer-visible document. The check, before
composing anything: `git ls-files .claude` in the consumer repository prints nothing and
`.gitignore` carries `.claude/`; a hit is a blocker handled first, as the delivery package's
consumer-repository hygiene reference describes (tip removal, history purge with the
repository admin, both logged). Ruling of 2026-09-04, after one consumer was found carrying
agent overlays, memory and profile in the customer's repository for two months.
