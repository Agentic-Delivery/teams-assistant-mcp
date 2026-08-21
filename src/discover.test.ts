import { describe, expect, it } from 'vitest';
import { formatChat, formatChatList } from './discover.js';
import type { ChatSummary } from './graph/teams-chats.js';

const group: ChatSummary = {
  id: '19:0123456789abcdef0123456789abcdef@thread.v2',
  topic: 'Pilot chat',
  chatType: 'group',
  members: ['Alice Anderson', 'Bob Brown', 'Assistant (AI)'],
};

describe('formatChat', () => {
  it('puts the id on its own line so it can be copied whole', () => {
    const [firstLine] = formatChat(group).split('\n');
    expect(firstLine).toBe('19:0123456789abcdef0123456789abcdef@thread.v2');
  });

  it('shows type, topic and members', () => {
    const text = formatChat(group);
    expect(text).toContain('type     group');
    expect(text).toContain('topic    Pilot chat');
    expect(text).toContain('members  Alice Anderson, Bob Brown, Assistant (AI)');
  });

  it('marks a missing topic instead of printing an empty field', () => {
    expect(formatChat({ ...group, topic: '' })).toContain('topic    (none)');
  });

  it('caps the member list and counts the rest', () => {
    const crowded = {
      ...group,
      members: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    };
    expect(formatChat(crowded)).toContain('members  A, B, C, D, E, F, +2 more');
  });

  it('says so when Graph returned no member names', () => {
    expect(formatChat({ ...group, members: [] })).toContain('members  (none visible)');
  });
});

describe('formatChatList', () => {
  it('separates chats with a blank line', () => {
    const second: ChatSummary = {
      id: '19:aaaa-1111_bbbb-2222@unq.gbl.spaces',
      topic: 'Alice Anderson',
      chatType: 'oneOnOne',
      members: ['Alice Anderson', 'Assistant (AI)'],
    };
    const text = formatChatList([group, second]);
    expect(text).toContain('@thread.v2\n');
    expect(text.split('\n\n')).toHaveLength(2);
    expect(text).toContain('19:aaaa-1111_bbbb-2222@unq.gbl.spaces');
  });

  it('explains an empty result rather than printing nothing', () => {
    expect(formatChatList([])).toMatch(/No chats are visible/);
  });
});
