#!/bin/sh
# SessionStart hook - the trigger the SKILL.md's own "Wiring it into a project" section says
# every consumer needs. A skill teaches HOW; this line makes the agent reach for it: without
# an always-in-context mandate, sessions demonstrably fall back to plain-text walls after
# restarts/compactions (observed across consumer projects, 2026-09-02).
cat <<'JSON'
{"hookSpecificOutput":{"hookEventName":"SessionStart","additionalContext":"Teams formatting mandate (teams-styling plugin): before composing ANY styled Teams message - a multi-item post, a table, a findings report, a status-board edit - load the teams-styling skill first (Skill tool: teams-styling:teams-styling). Its verified rendering vocabulary is NOT in context after a session start or compaction. Plain text is only for one-sentence conversational replies. Never post markdown syntax into Teams - it renders literally."}}
JSON
