import { describe, expect, it } from 'vitest';
import { ChatAllowlist, ChatNotAllowedError } from './allowlist.js';

const pilot = { id: '19:pilot@thread.v2', label: 'Pilot', canPost: true };
const readOnly = { id: '19:readonly@thread.v2', label: 'Leadership', canPost: false };

describe('chat allowlist', () => {
  it('lets an allowlisted chat be read and posted to', () => {
    const allowlist = new ChatAllowlist([pilot]);

    expect(allowlist.assertReadable(pilot.id).label).toBe('Pilot');
    expect(allowlist.assertPostable(pilot.id).label).toBe('Pilot');
  });

  it('refuses a chat that is not on the list', () => {
    const allowlist = new ChatAllowlist([pilot]);

    expect(() => allowlist.assertReadable('19:someone-elses-chat@thread.v2')).toThrow(
      ChatNotAllowedError,
    );
    expect(() => allowlist.assertPostable('19:someone-elses-chat@thread.v2')).toThrow(
      ChatNotAllowedError,
    );
  });

  it('refuses posting to a read-only chat while still allowing reads', () => {
    const allowlist = new ChatAllowlist([readOnly]);

    expect(allowlist.assertReadable(readOnly.id).canPost).toBe(false);
    expect(() => allowlist.assertPostable(readOnly.id)).toThrow(ChatNotAllowedError);
  });

  it('names the operation in the refusal so the caller knows which control fired', () => {
    const allowlist = new ChatAllowlist([readOnly]);

    try {
      allowlist.assertPostable(readOnly.id);
      expect.unreachable('expected a refusal');
    } catch (error) {
      expect(error).toBeInstanceOf(ChatNotAllowedError);
      expect((error as ChatNotAllowedError).operation).toBe('post');
      expect((error as ChatNotAllowedError).chatId).toBe(readOnly.id);
    }
  });

  it('ignores surrounding whitespace but never case', () => {
    const allowlist = new ChatAllowlist([pilot]);

    expect(allowlist.isAllowed(`  ${pilot.id}  `)).toBe(true);
    expect(allowlist.isAllowed(pilot.id.toUpperCase())).toBe(false);
  });

  it('refuses to construct an empty allowlist', () => {
    expect(() => new ChatAllowlist([])).toThrow(/empty/i);
  });

  it('refuses a duplicate or blank chat id', () => {
    expect(() => new ChatAllowlist([pilot, { ...pilot, label: 'again' }])).toThrow(/twice/i);
    expect(() => new ChatAllowlist([{ id: '   ', label: 'blank', canPost: true }])).toThrow(
      /empty id/i,
    );
  });
});
