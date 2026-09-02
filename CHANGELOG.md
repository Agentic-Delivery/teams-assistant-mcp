# Changelog

Versions are the tagged tarballs (see README, "Building and testing"). Earlier entries were
written from the commit history when this file was started in 0.6.0, so they are short.

## 0.6.0 — 2026-09-02

Withdrawing a message, done properly. The occasion: the agent posted into a customer chat what
belonged in an internal one, and the only way out was editing the message down to a placeholder.

- `delete_chat_message` now checks ownership before it acts: the message is fetched and its
  author's AAD id compared with the signed-in account's (`/me`). Anyone else's message is refused
  with a `MessageOwnershipError` naming the author, nothing sent; so is a message whose author
  cannot be verified because `/me` failed. New `force` input skips the check (fetch included) —
  for when a human has said so, never otherwise.
- New `undo_delete_chat_message` tool — Graph's `undoSoftDelete`, same ownership rule and
  `force` escape.
- New `teams-delete <chatId> <messageId> [--undo] [--force]` CLI, same JSON-line/exit-code
  contract as the rest of the family (2 usage, 3 allowlist, 1 anything else). A machine on the
  `~/.teams-assistant/*.mjs` shim convention adds `delete.mjs` for it.
- A 403 on either action now names the missing delegated permission (`Chat.ReadWrite`) and the
  custom-app-registration cause, Graph's own text kept — the treatment the attachment
  download's 403 got in 0.5.0. With the shipped first-party client id nothing changes: the
  `.default` token already carries the permission.
- The ownership read goes through the same throttle-aware single-message fetch a quoted reply
  uses (list-scan fallback on a 429). It does not take the inbox quota yield — one GET with a
  cheaper fallback, the same posture as `reply_chat_message`.

## 0.5.0 — 2026-09-02

- `download_chat_attachments` and `list_chat_attachments` tools; `teams-attachments` CLI;
  attachment metadata in `read_chat_messages` and `teams-read`.
- Quota coordination: the inbox poller yields to attachment reads via a yield file, after live
  measurement showed retries alone never got through a running poller.
- Binary downloads retry throttles like every other read; the retry sleep cap is 90 s, above the
  62 s window Graph names on that family.
- A 403 on a SharePoint download names `Files.Read.All` / `Sites.Read.All` and admin consent.

## 0.4.2 — 2026-09-02

- `send_chat_file` grants every other chat member read access on the uploaded item before
  posting (the "can't be viewed" card bug), with the grant verified from the `/invite` response
  rather than the HTTP status.
- New `teams-send-file` CLI, streaming one JSON line per file as each lands.

## 0.4.1 — 2026-09-01

- Operational hardening: per-chat member cache for mention resolution, Retry-After named on both
  the CLI and MCP surfaces through one renderer, stuck-auth recovery for the inbox poller, and a
  fresh watermark no longer backfills old messages.

## 0.4.0 — 2026-08-26

- Real @mentions (`mentions` on send/reply/edit, `--mention` on the CLIs) and pin management
  (`pin_chat_message`, `unpin_chat_message`, `list_pinned_messages`, `teams-pin`, `teams-unpin`).

## 0.3.0 — 2026-08-25

- Opt-in raw-HTML path: `format: 'html'` on send/edit, `--html` on `teams-post`/`teams-edit`;
  the `teams-styling` companion skill.

## 0.2.2 — 2026-08-25

- `get_chat_attachment` without an id picks the first real file, never the quote card.

## 0.2.1 — 2026-08-25

- Throttle gates keyed per resource family; the single-message fetch falls back to the list
  under a 429.

## 0.2.0 — 2026-08-25

- Readback-before-retry sends (`ReliableTeamsChats`), the standalone CLIs and their output
  contract, the throttle gate.
