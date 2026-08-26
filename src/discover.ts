import type { ChatSummary } from './graph/teams-chats.js';

/**
 * Text formatting for `npm run discover-chats`. Kept apart from the script so the shape of the
 * output is unit-testable without a Graph sign-in.
 */

const MAX_NAMES = 6;

export function formatChat(chat: ChatSummary): string {
  const names = chat.members.map((member) => member.displayName);
  const shown = names.slice(0, MAX_NAMES).join(', ');
  const overflow = names.length - MAX_NAMES;
  const members =
    names.length === 0 ? '(none visible)' : overflow > 0 ? `${shown}, +${overflow} more` : shown;
  return [
    chat.id,
    `  type     ${chat.chatType}`,
    `  topic    ${chat.topic || '(none)'}`,
    `  members  ${members}`,
  ].join('\n');
}

export function formatChatList(chats: ChatSummary[]): string {
  if (chats.length === 0) {
    return (
      'No chats are visible to this account. Auth worked, so either the account has not been ' +
      'added to any chat yet, or the licence is missing (the probe tells those apart).'
    );
  }
  return chats.map(formatChat).join('\n\n');
}
