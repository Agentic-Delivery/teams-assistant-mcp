import { beforeEach, describe, expect, it } from 'vitest';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { ChatAllowlist } from './allowlist.js';
import { buildServer } from './server.js';
import { applyWatermark, toChatMessage } from './messages.js';
import type {
  AttachmentPayload,
  ChatSummary,
  OutboundFile,
  OutboundImage,
  TeamsChatsPort,
} from './graph/teams-chats.js';

const PILOT = '19:pilot@thread.v2';
const WATCHED = '19:watched@thread.v2';
const OUTSIDE = '19:not-ours@thread.v2';

/**
 * Test double for the Graph secondary port. It records what reached Graph, which is how the
 * allowlist tests prove a refusal happened before the network rather than after it.
 */
class FakeTeamsChats implements TeamsChatsPort {
  readonly sent: Array<{ chatId: string; text: string }> = [];
  readonly reads: string[] = [];
  messages: Record<string, Array<{ id: string; created: string; text: string }>> = {
    [PILOT]: [
      { id: 'm1', created: '2026-08-19T08:00:00Z', text: 'first' },
      { id: 'm2', created: '2026-08-19T09:00:00Z', text: 'second' },
    ],
    [WATCHED]: [{ id: 'w1', created: '2026-08-19T07:00:00Z', text: 'watching' }],
  };
  failingChats = new Set<string>();

  async listChats(): Promise<ChatSummary[]> {
    return [
      { id: PILOT, topic: 'Pilot', chatType: 'group', members: ['Alice', 'Bob'] },
      { id: OUTSIDE, topic: 'HR private', chatType: 'group', members: ['HR'] },
    ];
  }

  async readMessages(chatId: string, since?: string) {
    this.reads.push(chatId);
    if (this.failingChats.has(chatId)) {
      throw new Error('chat unreachable');
    }
    const mapped = (this.messages[chatId] ?? []).map((message) =>
      toChatMessage(
        {
          id: message.id,
          createdDateTime: message.created,
          from: { user: { id: 'u', displayName: 'Alice' } },
          body: { contentType: 'text', content: message.text },
        },
        chatId,
      ),
    );
    return applyWatermark(mapped, since);
  }

  async sendMessage(chatId: string, text: string) {
    this.sent.push({ chatId, text });
    return toChatMessage(
      { id: 'sent-1', chatId, createdDateTime: '2026-08-19T10:00:00Z', body: { content: text } },
      chatId,
    );
  }

  readonly replies: Array<{ chatId: string; replyToMessageId: string; text: string }> = [];
  readonly edits: Array<{ chatId: string; messageId: string; newText: string }> = [];
  readonly deletes: Array<{ chatId: string; messageId: string }> = [];

  async replyToMessage(chatId: string, replyToMessageId: string, text: string) {
    this.replies.push({ chatId, replyToMessageId, text });
    return toChatMessage(
      { id: 'reply-1', chatId, createdDateTime: '2026-08-19T10:00:00Z', body: { content: text } },
      chatId,
    );
  }

  async editMessage(chatId: string, messageId: string, newText: string) {
    this.edits.push({ chatId, messageId, newText });
  }

  async deleteMessage(chatId: string, messageId: string) {
    this.deletes.push({ chatId, messageId });
  }

  readonly sentImages: Array<{ chatId: string; image: OutboundImage; text?: string }> = [];
  readonly sentFiles: Array<{ chatId: string; file: OutboundFile; text?: string }> = [];

  async sendImage(chatId: string, image: OutboundImage, text?: string) {
    this.sentImages.push({ chatId, image, ...(text !== undefined ? { text } : {}) });
    return toChatMessage(
      { id: 'img-1', chatId, createdDateTime: '2026-08-19T10:00:00Z', body: { content: '' } },
      chatId,
    );
  }

  async sendFile(chatId: string, file: OutboundFile, text?: string) {
    this.sentFiles.push({ chatId, file, ...(text !== undefined ? { text } : {}) });
    return toChatMessage(
      { id: 'file-1', chatId, createdDateTime: '2026-08-19T10:00:00Z', body: { content: '' } },
      chatId,
    );
  }

  reactions: Array<{ chatId: string; messageId: string; emoji: string }> = [];

  async setReaction(chatId: string, messageId: string, reactionType: string) {
    this.reactions.push({ chatId, messageId, emoji: reactionType });
  }

  async getAttachment(): Promise<AttachmentPayload> {
    return {
      bytes: new Uint8Array([1, 2, 3]),
      contentType: 'application/pdf',
      name: '../../escape.pdf',
    };
  }
}

let chats: FakeTeamsChats;
let downloadDir: string;

async function connect() {
  chats = new FakeTeamsChats();
  downloadDir = mkdtempSync(join(tmpdir(), 'teams-mcp-test-'));
  const server = buildServer({
    chats,
    allowlist: new ChatAllowlist([
      { id: PILOT, label: 'Pilot', canPost: true },
      { id: WATCHED, label: 'Leadership', canPost: false },
    ]),
    assistantDisplayName: 'Assistant (AI)',
    downloadDir,
  });
  const client = new Client({ name: 'test', version: '1.0.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
  return client;
}

async function call(client: Client, name: string, args: Record<string, unknown> = {}) {
  const result = (await client.callTool({ name, arguments: args })) as {
    isError?: boolean;
    content: Array<{ type: string; text: string }>;
  };
  const text = result.content.map((part) => part.text).join('');
  return { isError: result.isError === true, text, json: () => JSON.parse(text) as never };
}

let client: Client;
beforeEach(async () => {
  client = await connect();
});

describe('tool surface', () => {
  it('offers the ten chat tools plus the poll helper', async () => {
    const { tools } = await client.listTools();

    expect(tools.map((tool) => tool.name).sort()).toEqual([
      'delete_chat_message',
      'edit_chat_message',
      'get_chat_attachment',
      'list_chats',
      'poll_chats',
      'react_to_chat_message',
      'read_chat_messages',
      'reply_chat_message',
      'send_chat_file',
      'send_chat_image',
      'send_chat_message',
    ]);
  });
});

describe('react_to_chat_message', () => {
  it('puts the emoji on the message in an allowlisted chat', async () => {
    const result = await call(client, 'react_to_chat_message', {
      chatId: PILOT,
      messageId: 'msg-9',
      emoji: '👍',
    });

    expect(result.isError).toBeFalsy();
    expect(chats.reactions).toEqual([{ chatId: PILOT, messageId: 'msg-9', emoji: '👍' }]);
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'react_to_chat_message', {
      chatId: '19:not-ours@thread.v2',
      messageId: 'msg-9',
      emoji: '👍',
    });

    expect(result.isError).toBe(true);
    expect(chats.reactions).toEqual([]);
  });
});

describe('reply_chat_message', () => {
  it('posts a quoted reply into an allowlisted chat that permits posting', async () => {
    const result = await call(client, 'reply_chat_message', {
      chatId: PILOT,
      replyToMessageId: 'm1',
      text: 'good question',
    });

    expect(result.isError).toBe(false);
    expect(chats.replies).toEqual([
      { chatId: PILOT, replyToMessageId: 'm1', text: 'good question' },
    ]);
  });

  it('refuses a read-only chat', async () => {
    const result = await call(client, 'reply_chat_message', {
      chatId: WATCHED,
      replyToMessageId: 'w1',
      text: 'x',
    });

    expect(result.isError).toBe(true);
    expect(chats.replies).toEqual([]);
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'reply_chat_message', {
      chatId: OUTSIDE,
      replyToMessageId: 'm1',
      text: 'x',
    });

    expect(result.isError).toBe(true);
    expect(chats.replies).toEqual([]);
  });
});

describe('edit_chat_message', () => {
  it('edits a message in an allowlisted chat that permits posting', async () => {
    const result = await call(client, 'edit_chat_message', {
      chatId: PILOT,
      messageId: 'm1',
      newText: 'corrected',
    });

    expect(result.isError).toBe(false);
    expect(chats.edits).toEqual([{ chatId: PILOT, messageId: 'm1', newText: 'corrected' }]);
  });

  it('refuses a read-only chat', async () => {
    const result = await call(client, 'edit_chat_message', {
      chatId: WATCHED,
      messageId: 'w1',
      newText: 'x',
    });

    expect(result.isError).toBe(true);
    expect(chats.edits).toEqual([]);
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'edit_chat_message', {
      chatId: OUTSIDE,
      messageId: 'm1',
      newText: 'x',
    });

    expect(result.isError).toBe(true);
    expect(chats.edits).toEqual([]);
  });
});

describe('delete_chat_message', () => {
  it('soft-deletes a message in an allowlisted chat that permits posting', async () => {
    const result = await call(client, 'delete_chat_message', { chatId: PILOT, messageId: 'm1' });

    expect(result.isError).toBe(false);
    expect(chats.deletes).toEqual([{ chatId: PILOT, messageId: 'm1' }]);
  });

  it('refuses a read-only chat', async () => {
    const result = await call(client, 'delete_chat_message', { chatId: WATCHED, messageId: 'w1' });

    expect(result.isError).toBe(true);
    expect(chats.deletes).toEqual([]);
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'delete_chat_message', { chatId: OUTSIDE, messageId: 'm1' });

    expect(result.isError).toBe(true);
    expect(chats.deletes).toEqual([]);
  });
});

describe('list_chats', () => {
  it('shows the allowlisted chats and never the ones outside it', async () => {
    const result = await call(client, 'list_chats');
    const payload = result.json() as { chats: Array<{ id: string; visibleToAccount: boolean }> };

    expect(payload.chats.map((chat) => chat.id)).toEqual([PILOT, WATCHED]);
    expect(result.text).not.toContain(OUTSIDE);
    expect(result.text).not.toContain('HR private');
  });

  it('marks an allowlisted chat the account cannot actually see', async () => {
    const payload = (await call(client, 'list_chats')).json() as {
      chats: Array<{ id: string; visibleToAccount: boolean }>;
    };

    expect(payload.chats.find((chat) => chat.id === PILOT)?.visibleToAccount).toBe(true);
    expect(payload.chats.find((chat) => chat.id === WATCHED)?.visibleToAccount).toBe(false);
  });
});

describe('read_chat_messages', () => {
  it('reads an allowlisted chat and hands back a watermark', async () => {
    const payload = (await call(client, 'read_chat_messages', { chatId: PILOT })).json() as {
      messages: Array<{ id: string; text: string }>;
      watermark: string;
    };

    expect(payload.messages.map((message) => message.id)).toEqual(['m1', 'm2']);
    expect(payload.watermark).toBe('2026-08-19T09:00:00Z');
  });

  it('returns only what arrived after the watermark', async () => {
    const payload = (
      await call(client, 'read_chat_messages', { chatId: PILOT, since: '2026-08-19T08:00:00Z' })
    ).json() as { messages: Array<{ id: string }> };

    expect(payload.messages.map((message) => message.id)).toEqual(['m2']);
  });

  it('reads a chat marked read-only', async () => {
    const result = await call(client, 'read_chat_messages', { chatId: WATCHED });

    expect(result.isError).toBe(false);
  });

  it('refuses a chat outside the allowlist without calling Graph', async () => {
    const result = await call(client, 'read_chat_messages', { chatId: OUTSIDE });

    expect(result.isError).toBe(true);
    expect(result.text).toContain('not on the allowlist');
    expect(chats.reads).toEqual([]);
  });
});

describe('send_chat_message', () => {
  it('posts to an allowlisted chat that permits posting', async () => {
    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: 'Hi' });

    expect(result.isError).toBe(false);
    expect(chats.sent).toEqual([{ chatId: PILOT, text: 'Hi' }]);
  });

  it('refuses to post to a chat outside the allowlist and sends nothing', async () => {
    const result = await call(client, 'send_chat_message', { chatId: OUTSIDE, text: 'Hi' });

    expect(result.isError).toBe(true);
    expect(chats.sent).toEqual([]);
  });

  it('refuses to post to a read-only chat even though reading it is fine', async () => {
    const result = await call(client, 'send_chat_message', { chatId: WATCHED, text: 'Hi' });

    expect(result.isError).toBe(true);
    expect(result.text).toContain('for post');
    expect(chats.sent).toEqual([]);
  });
});

describe('send_chat_image', () => {
  const PNG_BASE64 = Buffer.from([137, 80, 78, 71]).toString('base64');

  it('posts base64 image bytes to an allowlisted chat that permits posting', async () => {
    const result = await call(client, 'send_chat_image', {
      chatId: PILOT,
      base64: PNG_BASE64,
      mime: 'image/png',
      text: 'here is the image',
    });

    expect(result.isError).toBe(false);
    expect(chats.sentImages).toHaveLength(1);
    expect(chats.sentImages[0]).toMatchObject({
      chatId: PILOT,
      text: 'here is the image',
      image: { contentType: 'image/png' },
    });
    expect([...chats.sentImages[0]!.image.bytes]).toEqual([137, 80, 78, 71]);
  });

  it('reads a local file and infers the mime from its extension', async () => {
    const path = join(downloadDir, 'diagram.png');
    writeFileSync(path, Buffer.from([137, 80, 78, 71]));

    const result = await call(client, 'send_chat_image', { chatId: PILOT, path });

    expect(result.isError).toBe(false);
    expect(chats.sentImages[0]?.image.contentType).toBe('image/png');
  });

  it('refuses base64 without a mime instead of guessing', async () => {
    const result = await call(client, 'send_chat_image', { chatId: PILOT, base64: PNG_BASE64 });

    expect(result.isError).toBe(true);
    expect(chats.sentImages).toEqual([]);
  });

  it('refuses both path and base64 at once', async () => {
    const result = await call(client, 'send_chat_image', {
      chatId: PILOT,
      path: '/x.png',
      base64: PNG_BASE64,
      mime: 'image/png',
    });

    expect(result.isError).toBe(true);
    expect(result.text).toContain('exactly one');
  });

  it('refuses a read-only chat before touching any file or Graph', async () => {
    const result = await call(client, 'send_chat_image', {
      chatId: WATCHED,
      base64: PNG_BASE64,
      mime: 'image/png',
    });

    expect(result.isError).toBe(true);
    expect(chats.sentImages).toEqual([]);
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'send_chat_image', {
      chatId: OUTSIDE,
      base64: PNG_BASE64,
      mime: 'image/png',
    });

    expect(result.isError).toBe(true);
    expect(chats.sentImages).toEqual([]);
  });
});

describe('send_chat_file', () => {
  it('reads the local file and shares it under its basename', async () => {
    const path = join(downloadDir, 'note.txt');
    writeFileSync(path, 'hi');

    const result = await call(client, 'send_chat_file', { chatId: PILOT, path, text: 'attached' });

    expect(result.isError).toBe(false);
    expect(chats.sentFiles).toHaveLength(1);
    expect(chats.sentFiles[0]).toMatchObject({
      chatId: PILOT,
      text: 'attached',
      file: { name: 'note.txt' },
    });
    expect(Buffer.from(chats.sentFiles[0]!.file.bytes).toString()).toBe('hi');
  });

  it('refuses a read-only chat', async () => {
    const result = await call(client, 'send_chat_file', { chatId: WATCHED, path: '/tmp/x.txt' });

    expect(result.isError).toBe(true);
    expect(chats.sentFiles).toEqual([]);
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'send_chat_file', { chatId: OUTSIDE, path: '/tmp/x.txt' });

    expect(result.isError).toBe(true);
    expect(chats.sentFiles).toEqual([]);
  });
});

describe('get_chat_attachment', () => {
  it('writes the attachment into the download directory and returns the path', async () => {
    const payload = (
      await call(client, 'get_chat_attachment', { chatId: PILOT, messageId: 'm1' })
    ).json() as { path: string; bytes: number };

    expect(payload.path.startsWith(downloadDir)).toBe(true);
    expect(payload.bytes).toBe(3);
    expect([...readFileSync(payload.path)]).toEqual([1, 2, 3]);
  });

  it('cannot be walked out of the download directory by a hostile file name', async () => {
    const payload = (
      await call(client, 'get_chat_attachment', { chatId: PILOT, messageId: 'm1' })
    ).json() as { path: string };

    expect(payload.path).toBe(join(downloadDir, 'm1-escape.pdf'));
  });

  it('refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'get_chat_attachment', {
      chatId: OUTSIDE,
      messageId: 'm1',
    });

    expect(result.isError).toBe(true);
  });
});

describe('poll_chats', () => {
  it('covers every allowlisted chat and returns per-chat watermarks', async () => {
    const payload = (await call(client, 'poll_chats')).json() as {
      messagesByChat: Record<string, unknown[]>;
      watermarks: Record<string, string>;
    };

    expect(Object.keys(payload.messagesByChat).sort()).toEqual([PILOT, WATCHED].sort());
    expect(payload.watermarks[PILOT]).toBe('2026-08-19T09:00:00Z');
    expect(payload.watermarks[WATCHED]).toBe('2026-08-19T07:00:00Z');
  });

  it('omits chats with nothing new but keeps their watermark', async () => {
    const first = (await call(client, 'poll_chats')).json() as { watermarks: Record<string, string> };
    const second = (await call(client, 'poll_chats', { watermarks: first.watermarks })).json() as {
      messagesByChat: Record<string, unknown[]>;
      watermarks: Record<string, string>;
    };

    expect(second.messagesByChat).toEqual({});
    expect(second.watermarks).toEqual(first.watermarks);
  });

  it('never polls a chat outside the allowlist', async () => {
    await call(client, 'poll_chats');

    expect(chats.reads).not.toContain(OUTSIDE);
  });

  it('reports one failing chat without losing the others', async () => {
    chats.failingChats.add(WATCHED);

    const payload = (await call(client, 'poll_chats')).json() as {
      messagesByChat: Record<string, unknown[]>;
      failures: Array<{ chatId: string }>;
    };

    expect(Object.keys(payload.messagesByChat)).toEqual([PILOT]);
    expect(payload.failures.map((failure) => failure.chatId)).toEqual([WATCHED]);
  });
});
