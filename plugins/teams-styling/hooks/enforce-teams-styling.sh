#!/bin/sh
# PreToolUse hook on Bash: HARD-ENFORCE Teams message styling at composition time.
#
# SHIPPED BY THE teams-styling PLUGIN SINCE 0.2.0
#   Registered through the plugin's own hooks/hooks.json, so no consuming project has to install
#   a copy. Plugin hooks are read WHEN A SESSION STARTS (or on /reload-plugins) - never live, so
#   an install or upgrade mid-session takes effect in the NEXT session. They MERGE with any
#   project/user settings hooks; every matching hook runs and any single deny wins, so a project
#   that still has its own copy is safe until it removes it.
#   Layers around this one: the CLI guard in teams-assistant-mcp >= 0.7.2 is the authoritative
#   backstop - notably for bodies this hook cannot inspect (piped from a file) - and `--text` on
#   the send's own line is the deliberate-plain escape that both layers honour.
#
# WHAT IT ENFORCES
#   A Teams send (post.mjs / reply.mjs / edit.mjs, or the teams-post / teams-reply /
#   teams-edit CLI names) whose body reads as structured text must go out styled: `--html`.
#   The hook inspects only what the Bash command string itself reveals - heredoc content, or a
#   quoted printf/echo argument upstream of the pipe. A body piped in from a FILE is not
#   inspectable here and is allowed through: the CLI-side guard in teams-assistant-mcp is the
#   backstop for that path.
#   Each send is judged on its own line, so one command carrying a styled send and a plain
#   one is still caught.
#
# THE CLASSIFIER IS THE CLI'S, PORTED (2026-09-21)
#   teams-assistant-mcp's CLI guard (`structuredTextReason` in src/cli/common.ts, shipped in
#   0.7.2) is the AUTHORITATIVE layer; this hook is
#   defence-in-depth in front of it and must never deny what the CLI would allow. The rules
#   below are that function's final semantics, transliterated to sh/awk:
#     - >= 3 sentence terminators. A `.`/`!`/`?` (run) counts only when the NEXT character is
#       whitespace or the end of the body - which is what keeps "v0.7.2", "findings.md",
#       "30.9 percent" and URLs from counting, with no maintained list.
#     - MINUS the short fixed abbreviation list whose final dot IS followed by whitespace and
#       still is not a sentence end: e.g. / i.e. / kl. (case-insensitive) and No. (case
#       SENSITIVE, so an ordinary sentence ending in "no." is not exempted). Both word-boundary
#       guarded, as in the CLI's \b.
#     - OR a blank line.
#     - OR a real pipe-table row: a line that both STARTS and ENDS with `|` (not merely a line
#       containing pipes - a shell pipeline is not a table).
#     - OR >= 4 non-empty lines (a bulleted list has no terminators at all).
#     - OR a body longer than 400 characters.
#
# THE FINDING IT IMPLEMENTS
#   Styling audit of 21 agent sessions, 2026-09-21. Compliance with the teams-styling mandate
#   was 13.1% in sessions where the skill had not been loaded versus 63.2% after a load
#   (n=373). The mandate existed in four written places and not one of them fires at the
#   instant the Bash command is composed. This hook is the only control that does not depend
#   on model discipline, and it carries the rule into context exactly when the mistake is
#   being made.
#
# ESCAPE HATCH
#   `--text` on the send's own line means "plain on purpose" and is allowed through
#   unconditionally, as is `--html`.
#
# Contract: reads the PreToolUse JSON on stdin; exit 0 = allow, exit 2 + stderr = deny.
# (Exit 1 is NON-blocking in the hook protocol and must never be used for the deny path.)
# Side-effect free; requires `jq` on PATH.
set -u

input=$(cat)
cmd=$(printf '%s' "$input" | jq -r '.tool_input.command // ""' 2>/dev/null) || exit 0
[ -z "$cmd" ] && exit 0

# Cheap gate: anything that is not a Teams send leaves immediately.
case "$cmd" in
  *post.mjs*|*reply.mjs*|*edit.mjs*|*teams-post*|*teams-reply*|*teams-edit*) ;;
  *) exit 0 ;;
esac

AWK_JUDGE=$(cat <<'AWKEOF'
# --- the CLI classifier, ported (see header) -------------------------------------------------
function isword(c) { return (c ~ /[A-Za-z0-9_]/) }

# SENTENCE_TERMINATOR: /[.!?](?=\s|$)/g. A run of . ! ? scores once, and only when what follows
# it is whitespace or the end of the body.
function count_terminators(s,    i, j, n, nx, cnt) {
  n = length(s); i = 1; cnt = 0
  while (i <= n) {
    if (substr(s, i, 1) ~ /[.!?]/) {
      j = i
      while (j <= n && substr(s, j, 1) ~ /[.!?]/) j++
      nx = (j > n) ? "" : substr(s, j, 1)
      if (nx == "" || nx ~ /[ \t\n\r]/) cnt++
      i = j
    } else i++
  }
  return cnt
}

# ABBREVIATION_TERMINATOR_CI (e.g. / i.e. / kl., case-insensitive) and _CS (No., case-sensitive).
# Each needs the same "followed by whitespace or end" shape as a terminator, plus the CLI's \b
# in front: the preceding character must not be a word character.
function count_abbrev(s,    i, n, low, cnt, prev, nx, hit, w) {
  low = tolower(s); n = length(s); cnt = 0
  for (i = 1; i <= n; i++) {
    hit = 0; w = 0
    if (substr(low, i, 4) == "e.g." || substr(low, i, 4) == "i.e.") { hit = 1; w = 4 }
    else if (substr(low, i, 3) == "kl.") { hit = 1; w = 3 }
    else if (substr(s, i, 3) == "No.") { hit = 1; w = 3 }
    if (!hit) continue
    prev = (i == 1) ? "" : substr(s, i - 1, 1)
    if (prev != "" && isword(prev)) continue
    nx = (i + w > n) ? "" : substr(s, i + w, 1)
    if (nx == "" || nx ~ /[ \t\n\r]/) cnt++
  }
  return cnt
}

function structured(b,    nterm, nl, ne, i, t, tt, BL) {
  # A blank line inside the body is structure: it must be styled.
  if (b ~ /[^\n][ \t]*\n[ \t]*\n/) return 1
  nterm = count_terminators(b) - count_abbrev(b)
  if (nterm < 0) nterm = 0
  if (nterm >= 3) return 1
  nl = split(b, BL, "\n")
  ne = 0
  for (i = 1; i <= nl; i++) {
    t = BL[i]
    # A REAL pipe-table row: starts AND ends with a pipe. "cat x | grep y | head" does not.
    if (t ~ /^[ \t]*\|.*\|[ \t]*$/) return 1
    tt = t; gsub(/[ \t\r]/, "", tt)
    if (tt != "") ne++
  }
  if (ne >= 4) return 1              # MANY_LINES_THRESHOLD
  if (length(b) > 400) return 1      # LONG_BODY_THRESHOLD
  return 0
}

function judge(L,    sl, shimstart, head, tail, hd, delim, i, j, k, t, pre, tmp, cnt, len,
                     ch, q, c, body, probe, run, fmt) {
  sl = line[L]
  if (!match(sl, shim)) return 0
  shimstart = RSTART
  body = ""
  head = substr(sl, 1, shimstart - 1)

  # Case 1: a heredoc opened on the shim line feeds it the body - either side of the shim
  # token (`cat <<EOF | node .../post.mjs id` and `node .../post.mjs id <<EOF`). On the tail
  # side, stop at the first pipe/separator: a later `python3 - <<EOF` in the same command is
  # somebody else's heredoc, not this send's body.
  tail = substr(sl, shimstart)
  if (match(tail, /[|;&]/)) tail = substr(tail, 1, RSTART - 1)
  hd = ""
  if (match(head, /<<-?[ \t]*["']?[A-Za-z_][A-Za-z0-9_]*["']?/)) hd = substr(head, RSTART, RLENGTH)
  else if (match(tail, /<<-?[ \t]*["']?[A-Za-z_][A-Za-z0-9_]*["']?/)) hd = substr(tail, RSTART, RLENGTH)
  if (hd != "") {
    delim = hd
    sub(/^<<-?[ \t]*/, "", delim)
    gsub(/["']/, "", delim)
    for (i = L + 1; i <= n; i++) {
      t = line[i]
      sub(/^[ \t]+/, "", t); sub(/[ \t]+$/, "", t)
      if (t == delim) break
      body = body line[i] "\n"
    }
  } else {
    # Case 2: the body is a quoted argument upstream of the pipe into the shim (printf,
    # echo). Take everything before the LAST pipe preceding the shim - that drops the
    # `| node --env-file=... .../post.mjs` tail and its dotted paths.
    pre = head
    k = 0
    for (j = length(pre); j >= 1; j--) if (substr(pre, j, 1) == "|") { k = j; break }
    if (k > 0) pre = substr(pre, 1, k - 1)

    # A quoted body may start on an earlier line. If the double quotes in `pre` are
    # unbalanced, walk backwards until they balance (bounded).
    tmp = pre
    cnt = gsub(/"/, "&", tmp)
    if (cnt % 2 == 1) {
      for (i = L - 1; i >= 1 && i > L - 200; i--) {
        pre = line[i] "\n" pre
        tmp = line[i]
        cnt += gsub(/"/, "&", tmp)
        if (cnt % 2 == 0) break
      }
    }

    # Pull out the contents of every quoted run. A run that is PURE format string ("%s\n",
    # "%s\n%s\n") is dropped rather than kept: it carries no sentence terminators, but since
    # the line-count rule arrived it would otherwise inflate the body by a phantom line and
    # let the hook deny a three-item printf the CLI allows. Literal \n escapes inside a kept
    # run become real newlines, so a genuinely multi-line printf body is counted as one.
    len = length(pre); i = 1
    while (i <= len) {
      ch = substr(pre, i, 1)
      if (ch == "\"" || ch == "'") {
        q = ch; i++
        run = ""
        while (i <= len) { c = substr(pre, i, 1); if (c == q) break; run = run c; i++ }
        fmt = run
        gsub(/%[-+ 0#]*[0-9]*(\.[0-9]+)?[a-zA-Z]/, "", fmt)
        gsub(/\\[nrt]/, "", fmt)
        gsub(/[ \t]/, "", fmt)
        if (fmt != "") {
          gsub(/\\n/, "\n", run)
          body = body run "\n"
        }
      }
      i++
    }
  }

  # Nothing inspectable (body piped from a file, or a redirect) -> allow; C1 is the backstop.
  probe = body
  gsub(/[ \t\n]/, "", probe)
  if (probe == "") return 0

  return structured(body)
}
{ line[NR] = $0 }
END {
  n = NR
  shim = "(post|reply|edit)\\.mjs|teams-(post|reply|edit)"
  for (L = 1; L <= n; L++) {
    if (line[L] !~ shim) continue
    if (line[L] ~ /--html|--text/) continue   # explicit format decision for THIS send
    if (judge(L)) { print "STYLE"; exit }
  }
  print "OK"
}
AWKEOF
)

verdict=$(printf '%s' "$cmd" | awk "$AWK_JUDGE")
[ "$verdict" = "STYLE" ] || exit 0

cat >&2 <<'MSG'
BLOCKED by the teams-styling hook: this Teams send reads as structured text - more than two
sentences, or a blank line, a pipe-table row, four or more lines, or past 400 characters - and
it carries neither --html nor --text. (Same classifier as the teams-* CLI guard, which would
refuse it one layer further in.)

The rule - the teams-styling skill: anything beyond two sentences is NOT sent as plain text.
Walls of plain text are the single most-corrected defect in agent-posted channel output.

Load the skill before composing the body:  Skill(teams-styling:teams-styling)

Correct shape - the chat id is POSITIONAL AND FIRST, the flags come after it:

  cat <<'EOF' | node --env-file=<env> <path>/post.mjs <chatId> --html
  <b>One-line headline</b>
  <div>&nbsp;</div>
  <div>Lead sentence.</div>
  <ul><li><b>Point.</b> Detail.</li></ul>
  EOF

  (`post.mjs --html <chatId>` parses --html AS the chat id and is refused - flags AFTER the id.)

To send this plain on purpose, add --text.
MSG
exit 2
