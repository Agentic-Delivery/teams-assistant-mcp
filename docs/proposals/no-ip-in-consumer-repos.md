# Proposal: no package artifact in a consumer repository (teams-styling 0.1.5)

Status: Applied on the branch for teams-styling 0.1.5 (2026-09-04); merge pending the compatibility review and the owner's word
Filed: 2026-09-04, from the same incident as verified-delivery 2.12.0 (`proposals/no-ip-in-consumer-repos.md` in that repository carries the full record)
Applies to: `plugins/teams-styling/skills/teams-styling/SKILL.md` (§ *Wiring it into a project*), both manifests

## Why this package too

The incident was found in the delivery package's consumer, but the shape is the same here: a
project that adopts this skill writes things against it — the mandate line, an overlay with
its own board wording, the pinned-message ids, the mention roster — and every one of those
names the plugin or carries its conventions. Left to habit they end up in the customer's
repository beside the code. The ruling of 2026-09-04 covers anything belonging to a plugin and
anything that carries the method, so it covers these.

## What changed

The *Wiring it into a project* section, the one place this skill tells a project where to
write something, gains one paragraph: the mandate line and everything written against this
skill live in the consumer's `.claude/`, which is the clone of the provider's private
per-project repository and is never tracked by the consumer repository; the line names a
plugin, so it goes in the private `.claude/CLAUDE.md` or the profile, not in a customer-visible
document; the check is `git ls-files .claude` empty and `.gitignore` carrying `.claude/`, run
before composing anything, with a hit handled first as the delivery package's
consumer-repository hygiene reference describes. This skill has no session-start hook and no
checklist of its own, so the check sits in that section and nowhere else.

Version 0.1.4 → 0.1.5 in `.claude-plugin/marketplace.json` and the plugin's `plugin.json`,
following the package's rule that every content change bumps the version.
