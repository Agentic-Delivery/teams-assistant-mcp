export interface AllowedChat {
  /** Graph chat id, e.g. 19:abc...@thread.v2 */
  id: string;
  /** Human label, only used in tool output so the orchestrator can tell chats apart. */
  label: string;
  /** false means the server may read this chat but must refuse to post in it. */
  canPost: boolean;
}

export class ChatNotAllowedError extends Error {
  constructor(
    readonly chatId: string,
    readonly operation: 'read' | 'post',
  ) {
    super(
      `Chat ${chatId} is not on the allowlist for ${operation}. ` +
        `Add it to the allowlist config (allowedChats) if this is intended.`,
    );
    this.name = 'ChatNotAllowedError';
  }
}

/**
 * The control that makes a broadly-scoped service account acceptable.
 *
 * The account's token carries every delegated Teams scope its first-party client id grants, so
 * the identity itself can reach every chat the account is a member of. This class is the only
 * thing standing between that and the world: every chat id entering a Graph call passes through
 * assertReadable or assertPostable first, and an empty allowlist is a startup error rather than
 * an open door.
 */
export class ChatAllowlist {
  private readonly byId: Map<string, AllowedChat>;

  constructor(entries: readonly AllowedChat[]) {
    if (entries.length === 0) {
      throw new Error(
        'Chat allowlist is empty. The server refuses to start without at least one allowed chat — ' +
          'an empty list is almost always a misconfigured config file, not an intent to allow nothing.',
      );
    }
    this.byId = new Map();
    for (const entry of entries) {
      const id = entry.id.trim();
      if (id === '') {
        throw new Error('Chat allowlist contains an entry with an empty id.');
      }
      if (this.byId.has(id)) {
        throw new Error(`Chat allowlist contains ${id} twice.`);
      }
      this.byId.set(id, { ...entry, id });
    }
  }

  /** Chat ids are Graph thread ids and are compared exactly — no case folding. */
  private lookup(chatId: string): AllowedChat | undefined {
    return this.byId.get(chatId.trim());
  }

  assertReadable(chatId: string): AllowedChat {
    const entry = this.lookup(chatId);
    if (!entry) {
      throw new ChatNotAllowedError(chatId, 'read');
    }
    return entry;
  }

  assertPostable(chatId: string): AllowedChat {
    const entry = this.lookup(chatId);
    if (!entry || !entry.canPost) {
      throw new ChatNotAllowedError(chatId, 'post');
    }
    return entry;
  }

  isAllowed(chatId: string): boolean {
    return this.lookup(chatId) !== undefined;
  }

  labelFor(chatId: string): string | undefined {
    return this.lookup(chatId)?.label;
  }

  entries(): AllowedChat[] {
    return [...this.byId.values()];
  }
}
