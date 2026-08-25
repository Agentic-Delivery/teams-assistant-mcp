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
  // Partial live evidence already exists for this one: on 2026-08-24 an ad-hoc script sent raw
  // unicode glyphs (👍) as reactionType against this exact endpoint shape and the reactions
  // rendered in Teams. Still worth a proper live test: the documented vocabulary historically
  // named reactions (like, heart, …), and the glyph acceptance should be pinned, not folklore.
  it.todo('setReaction accepts a raw emoji glyph as reactionType and it renders in Teams');
  it.todo('a send whose response is dropped mid-flight is recovered by the readback, not duplicated');
  it.todo('is refused by Graph, not just by us, if the account is removed from the chat');
});
