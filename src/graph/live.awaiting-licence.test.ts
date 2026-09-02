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
  // Partial live evidence already exists for this one too, from TWO manual captures (both
  // 2026-09-02, EPF011-delegated token, OneDrive for Business — see KNOWN-ISSUES.md's
  // "send_chat_file" entry for both full snapshots): (1) POST /me/drive/items/{id}/invite with
  // {"recipients":[{"objectId":"<aad-user-id>"}],"requireSignIn":true,"sendInvitation":false,
  // "roles":["read"]} answered 200, a subsequent GET .../permissions listed read grants with
  // grantedToV2.user for all six invited AAD users, and a human recipient confirmed the Teams
  // file card opened; (2) a re-grant on the same file, this time including the OWNER as a
  // recipient, showed the /invite response BODY ITSELF (not just the later GET) carries
  // grantedToV2.user.id per recipient, and that an owner-as-recipient IS echoed with its own
  // grant entry rather than silently dropped. Not yet captured live: the grantedToIdentitiesV2
  // variant sendFile's grant check also accepts defensively (documented Graph behaviour for some
  // permission kinds, never observed in either capture above). Still worth a proper live test
  // end-to-end through send_chat_file itself, against a throwaway chat, rather than resting on
  // manual captures.
  it.todo('send_chat_file grants every other chat member read access and their file card opens in Teams');
  // Partial live evidence already exists for this one: on 2026-08-24 an ad-hoc script sent raw
  // unicode glyphs (👍) as reactionType against this exact endpoint shape and the reactions
  // rendered in Teams. Still worth a proper live test: the documented vocabulary historically
  // named reactions (like, heart, …), and the glyph acceptance should be pinned, not folklore.
  it.todo('setReaction accepts a raw emoji glyph as reactionType and it renders in Teams');
  it.todo('a send whose response is dropped mid-flight is recovered by the readback, not duplicated');
  it.todo('the message LIST returns the same body and attachments as the single-message GET (the fetchMessage fallback rests on it)');
  it.todo('is refused by Graph, not just by us, if the account is removed from the chat');
});
