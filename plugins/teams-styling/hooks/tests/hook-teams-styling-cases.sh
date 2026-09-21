#!/bin/sh
# Test battery for ../enforce-teams-styling.sh, the hook the teams-styling plugin ships.
#
# The hook is defence-in-depth in front of teams-assistant-mcp's CLI guard
# (`structuredTextReason`, src/cli/common.ts). The contract these cases hold it to: it must
# catch structured plain-text sends at composition time, and it must NEVER deny a body the
# CLI would allow. Cases 1-11 are the battery the hook shipped with (2026-09-21); 12-19 were
# added when the CLI's final classifier was ported into it (abbreviations, real pipe-table
# rows, the line-count and length rules).
#
# Run:  sh plugins/teams-styling/hooks/tests/hook-teams-styling-cases.sh
#       (exit 0 = all cases as expected)
# Against an INSTALLED plugin cache copy, run it from inside that cache's hooks/tests/ - the
# path below resolves the hook relative to this file, so it follows the copy it ships with.
set -u

HOOK="$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)/enforce-teams-styling.sh"
pass=0; fail=0

# run <expect: ALLOW|DENY> <name> <<'EOF' ... command string ... EOF
run() {
  expect="$1"; name="$2"
  cmd=$(cat)
  out=$(printf '%s' "$cmd" | jq -Rs '{tool_input:{command:.}}' | sh "$HOOK" 2>/dev/null)
  rc=$?
  if [ "$rc" -eq 2 ]; then got="DENY"; else got="ALLOW"; fi
  if [ "$got" = "$expect" ]; then
    pass=$((pass+1)); printf 'ok    %-4s %s\n' "$got" "$name"
  else
    fail=$((fail+1)); printf 'FAIL  expected %s got %s: %s\n' "$expect" "$got" "$name"
  fi
}

SHIM='node --env-file=$HOME/.teams-assistant/.env $HOME/.teams-assistant/post.mjs 19:abc'

run DENY "1  multi-sentence plain heredoc" <<EOF
cat <<'MSG' | $SHIM
The upgrade branch is pushed. Compile is green. I will deploy after the review.
MSG
EOF

run ALLOW "2  same body, --html" <<EOF
cat <<'MSG' | $SHIM --html
The upgrade branch is pushed. Compile is green. I will deploy after the review.
MSG
EOF

run ALLOW "3  same body, --text (deliberate plain)" <<EOF
cat <<'MSG' | $SHIM --text
The upgrade branch is pushed. Compile is green. I will deploy after the review.
MSG
EOF

run ALLOW "4  not a Teams send at all" <<'EOF'
git commit -F - <<'MSG'
The upgrade branch is pushed. Compile is green. I will deploy after the review.
MSG
EOF

run ALLOW "5  body piped from a file (not inspectable; CLI guard is the backstop)" <<EOF
cat workspaces/reply.txt | $SHIM
EOF

run ALLOW "6  two sentences, plain" <<EOF
cat <<'MSG' | $SHIM
The upgrade branch is pushed. Compile is green.
MSG
EOF

run DENY "7  two sentences but a blank line" <<EOF
cat <<'MSG' | $SHIM
The upgrade branch is pushed.

Compile is green.
MSG
EOF

run DENY "8  quoted printf body, three sentences" <<EOF
printf '%s' "Branch pushed. Compile green. Deploying now." | $SHIM
EOF

run DENY "9  one styled send and one plain send in the same command" <<EOF
cat <<'A' | $SHIM --html
<b>Styled.</b>
A
cat <<'B' | $SHIM
Branch pushed. Compile green. Deploying now.
B
EOF

run ALLOW "10 one-liner with a version, a filename and a decimal" <<EOF
cat <<'MSG' | $SHIM
Shipped v0.7.2 to the server, notes in findings.md, coverage 30.9 percent.
MSG
EOF

run ALLOW "11 abbreviation does not count as a sentence end (e.g.)" <<EOF
cat <<'MSG' | $SHIM
Send it e.g. tomorrow morning. Thanks.
MSG
EOF

run DENY "12 five-line bulleted body, no terminators at all" <<EOF
cat <<'MSG' | $SHIM
- branch pushed
- compile green
- review open
- deploy queued
- retro tomorrow
MSG
EOF

run ALLOW "13 three short lines, no terminators (under the line threshold)" <<EOF
cat <<'MSG' | $SHIM
- branch pushed
- compile green
- review open
MSG
EOF

run DENY "14 a real pipe-table row" <<EOF
cat <<'MSG' | $SHIM
| centre | status |
MSG
EOF

run ALLOW "15 shell pipes inside the prose are not a table row" <<EOF
cat <<'MSG' | $SHIM
Run cat log | grep ERROR | head on the box and send me what it prints.
MSG
EOF

run DENY "16 one long sentence past 400 characters" <<EOF
cat <<'MSG' | $SHIM
The vendor re-cut for ClaimCenter and ContactManager is still outstanding at roughly one hundred and forty-one hours against three to four hours for the other two centres on the same case and roughly twenty-two hours for the August case, which means the constraint this week is the same constraint as last week and nobody has yet given it an owned row anywhere in the backlog or the decision queue we maintain
MSG
EOF

run ALLOW "17 an unrelated second heredoc later in the command is not this send's body" <<EOF
cat <<'MSG' | $SHIM
Status posted.
MSG
python3 - <<'PY'
print("one. two. three. four.")
PY
EOF

run ALLOW "18 'No.' is exempt (capital N)" <<EOF
cat <<'MSG' | $SHIM
Item No. 42 shipped. Done.
MSG
EOF

run DENY "19 lowercase 'no.' is NOT exempt" <<EOF
cat <<'MSG' | $SHIM
I said no. Then I left. Bye.
MSG
EOF

printf '\n%s passed, %s failed\n' "$pass" "$fail"
[ "$fail" -eq 0 ]
