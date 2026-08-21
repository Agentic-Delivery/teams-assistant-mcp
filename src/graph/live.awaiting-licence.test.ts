import { describe, it } from 'vitest';

/**
 * The end-to-end tests, and why they are not written yet.
 *
 * Everything below needs a real Teams chat and a signed-in account that Graph will actually let
 * into /chats. As of 2026-08-19 no such account exists: the account used during development
 * authenticates and gets a Graph token with the Teams delegated scopes, but has no Office 365
 * licence, so every /me/chats and /chats/{id}/messages call comes back
 * 403 "Failed to get license information for the user".
 *
 * Writing these against a mocked Graph would produce a green suite that proves nothing about the
 * one thing still unverified: whether Graph accepts these calls from this identity. So they stay
 * unwritten and visibly pending until a licensed account lands.
 *
 * When the account exists: run `npm run probe` first to confirm /me/chats returns chats and to
 * get the chat ids, put those in the allowlist config, then implement these against a throwaway
 * chat containing nobody but the assistant account and whoever is testing.
 */
describe.skip('live Graph (awaiting a Teams-licensed account)', () => {
  it.todo('lists the pilot chat with its real id and members');
  it.todo('reads messages from the pilot chat and returns a usable watermark');
  it.todo('returns nothing on a second read with the previous watermark');
  it.todo('posts a message that appears in Teams under the assistant display name');
  it.todo('downloads a file attachment shared in the pilot chat');
  it.todo('is refused by Graph, not just by us, if the account is removed from the chat');
});
