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
in, plain text out), so styling needs the raw-HTML send path. If your install doesn't have
one, plain text is the fallback — never post markdown syntax into Teams (it renders as
literal `**asterisks**`).

## When to style — and when not to

Styling is for making structured content scannable, not for decoration. Default is plain
text; reach for styling when the content has structure a reader must scan:

| Content | Format |
|---|---|
| Short conversational reply, answer to a question | Plain text — no styling |
| Findings / triage / comparison across items | `<table>` — never plain-text pipe walls |
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
