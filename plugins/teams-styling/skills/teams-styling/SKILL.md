---
name: teams-styling
description: "Styling doctrine for agent-posted Teams messages: when to style and when not to, the empirically verified HTML vocabulary Teams renders, known quirks, and the live status-board pattern for agent-progress transparency. Load before posting anything beyond a short conversational reply to a Teams chat — reports, findings tables, status updates, alerts."
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

## When to style — and when not to

Styling is for making structured content scannable, not for decoration. Default is plain
text; reach for styling when the content has structure a reader must scan:

| Content | Format |
|---|---|
| Short conversational reply, answer to a question | Plain text — no styling |
| Comparison across items with SHORT cells (ids, labels, verdicts, numbers) | `<table>` — never plain-text pipe walls |
| Findings with prose explanations | Headed sections, NOT a table: one `<b>` heading line per item, `<ul>` bullets under it, a bold `→ Fix:`/`→ Action:` closing bullet |
| Long report | Summary message with a table + link to the full document; not the whole report inline |
| Progress on multi-agent or long-running work | Live status board (pattern below) |
| Warning / blocker / decision needed | One `<b>` lead line + ⚠️/🔴 emoji; mention the human owner by name |
| Code, identifiers, config values | `<code>` inline, `<pre>` for blocks |
| Anything with a pass/fail dimension | Emoji + bold as primary signal; color only as reinforcement (never color alone — color-blind readers, theme drift) |

Restraint rules: at most one `<h2>` title per message (h1 is oversized in chat); style the
signal, not every word; a message that is all bold has no bold. If in doubt, plain text.

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

## Wiring it into a project

A skill teaches *how*; it does not make an agent reach for it. The consuming project's
conduct/profile must carry the mandate — one line is enough:

> All formatted channel output (tables, reports, status boards, alerts) follows the
> `teams-styling` skill; agents load it before composing any such message.

Without that line, expect agents to fall back to plain-text walls.
