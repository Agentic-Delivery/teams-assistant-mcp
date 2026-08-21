
## send_chat_file: recipients get no permission on the uploaded item (found 2026-08-20)

The tool uploads to the signed-in account's OneDrive and posts a file-reference card, but never
grants the chat's members permission on the drive item. In a one-on-one or meeting chat the
recipient hits "request access", and the request lands in the service account's unread mailbox.
Observed live: a chat member could not open a file the assistant had shared; fixed by hand with
a Graph `POST /me/drive/items/{id}/invite` (roles: read, sendInvitation: false), which worked
instantly.

Fix direction: after upload, enumerate the chat's members (`GET /chats/{id}/members`) and invite
each with read role before posting the card — or create an organization link scoped share if
policy allows. Until then, senders must expect the access-request dead end.

## Inbox poller: two server instances race on the same inbox (found 2026-08-21)

Every session that starts the server gets its own inbox poller, and they all default to the same
`~/.teams-assistant/inbox.jsonl`. Two Claude Code sessions at once means two pollers appending
the same messages and overwriting each other's `inbox-state.json`: duplicated lines, watermarks
jumping backwards, and extra Graph load. A mild version showed up during a smoke run, where an
older standalone polling daemon and the new in-process poller polling together earned a 429 from
Graph.

Fix direction: a lock file next to the inbox (first server takes it, later ones skip the poller
and say so on stderr), or a per-session `TEAMS_INBOX_PATH`. Until then, keep one session per
account, or set `TEAMS_INBOX_DISABLED=1` in the extra ones.
