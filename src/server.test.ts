import { beforeEach, describe, expect, it } from 'vitest';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { ChatAllowlist } from './allowlist.js';
import { GraphError } from './graph/graph-client.js';
import { buildServer } from './server.js';
import { applyWatermark, toChatMessage } from './messages.js';
import type {
  AttachmentPayload,
  ChatSummary,
  MentionTarget,
  OutboundFile,
  OutboundImage,
  PinnedMessage,
  TeamsChatsPort,
} from './graph/teams-chats.js';
import { resolveMentionTargets } from './graph/mentions.js';

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

  members: Record<string, Array<{ id?: string; displayName: string }>> = {
    [PILOT]: [
      { id: 'aad-alice', displayName: 'Alice Anderson' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ],
    [OUTSIDE]: [{ id: 'aad-hr', displayName: 'HR' }],
  };

  async listChats(): Promise<ChatSummary[]> {
    return [
      { id: PILOT, topic: 'Pilot', chatType: 'group', members: this.members[PILOT] ?? [] },
      { id: OUTSIDE, topic: 'HR private', chatType: 'group', members: this.members[OUTSIDE] ?? [] },
    ];
  }

  readonly resolveMentionsCalls: Array<{ chatId: string; names: readonly string[] }> = [];

  async resolveMentions(chatId: string, names: readonly string[]): Promise<MentionTarget[]> {
    this.resolveMentionsCalls.push({ chatId, names });
    return resolveMentionTargets(names, this.members[chatId] ?? []);
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

  readonly sentMentions: Array<readonly MentionTarget[]> = [];

  async sendMessage(chatId: string, text: string, mentions: readonly MentionTarget[] = []) {
    this.sent.push({ chatId, text });
    this.sentMentions.push(mentions);
    return toChatMessage(
      { id: 'sent-1', chatId, createdDateTime: '2026-08-19T10:00:00Z', body: { content: text } },
      chatId,
    );
  }

  readonly sentHtml: Array<{ chatId: string; html: string }> = [];

  async sendHtmlMessage(chatId: string, html: string, mentions: readonly MentionTarget[] = []) {
    this.sentHtml.push({ chatId, html });
    this.sentMentions.push(mentions);
    // contentType 'html' here so toChatMessage's htmlToText round-trip mirrors the real Graph
    // readback — the same shape a genuine sendHtmlMessage readback would return.
    return toChatMessage(
      {
        id: 'sent-html-1',
        chatId,
        createdDateTime: '2026-08-19T10:00:00Z',
        body: { contentType: 'html', content: html },
      },
      chatId,
    );
  }

  readonly replies: Array<{ chatId: string; replyToMessageId: string; text: string }> = [];
  readonly edits: Array<{ chatId: string; messageId: string; newText: string }> = [];
  readonly htmlEdits: Array<{ chatId: string; messageId: string; html: string }> = [];
  readonly deletes: Array<{ chatId: string; messageId: string }> = [];

  async replyToMessage(
    chatId: string,
    replyToMessageId: string,
    text: string,
    mentions: readonly MentionTarget[] = [],
  ) {
    this.replies.push({ chatId, replyToMessageId, text });
    this.sentMentions.push(mentions);
    return toChatMessage(
      { id: 'reply-1', chatId, createdDateTime: '2026-08-19T10:00:00Z', body: { content: text } },
      chatId,
    );
  }

  async editMessage(chatId: string, messageId: string, newText: string, mentions: readonly MentionTarget[] = []) {
    this.edits.push({ chatId, messageId, newText });
    this.sentMentions.push(mentions);
  }

  async editHtmlMessage(chatId: string, messageId: string, html: string, mentions: readonly MentionTarget[] = []) {
    this.htmlEdits.push({ chatId, messageId, html });
    this.sentMentions.push(mentions);
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

  // One pin per chat, replaced on every pin — mirrors the real single-pin-slot Graph behaviour.
  pinned: Record<string, PinnedMessage | undefined> = {};
  // Toggle to simulate Graph accepting the POST but the re-list NOT showing it pinned — the
  // pathological case the server must catch rather than reporting pinned:true regardless.
  pinConfirms = true;

  async pinMessage(chatId: string, messageId: string): Promise<PinnedMessage[]> {
    if (this.pinConfirms) {
      this.pinned[chatId] = { id: `pin-${messageId}`, messageId, preview: `preview of ${messageId}` };
    }
    return this.listPinnedMessages(chatId);
  }

  async unpinMessage(chatId: string, messageId: string): Promise<void> {
    if (this.pinned[chatId]?.messageId !== messageId) {
      throw new Error(`Message ${messageId} is not currently pinned in chat ${chatId} — nothing to unpin.`);
    }
    delete this.pinned[chatId];
  }

  async listPinnedMessages(chatId: string): Promise<PinnedMessage[]> {
    const entry = this.pinned[chatId];
    return entry ? [entry] : [];
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
  it('offers the chat tools, the pin tools, and the poll helper', async () => {
    const { tools } = await client.listTools();

    expect(tools.map((tool) => tool.name).sort()).toEqual([
      'delete_chat_message',
      'edit_chat_message',
      'get_chat_attachment',
      'list_chats',
      'list_pinned_messages',
      'pin_chat_message',
      'poll_chats',
      'react_to_chat_message',
      'read_chat_messages',
      'reply_chat_message',
      'send_chat_file',
      'send_chat_image',
      'send_chat_message',
      'unpin_chat_message',
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

  it('mentions are resolved against the chat and forwarded to replyToMessage', async () => {
    const result = await call(client, 'reply_chat_message', {
      chatId: PILOT,
      replyToMessageId: 'm1',
      text: 'Bob, thoughts?',
      mentions: ['Bob'],
    });

    expect(result.isError).toBe(false);
    expect(chats.sentMentions).toEqual([[{ name: 'Bob', id: 'aad-bob', displayName: 'Bob Brown' }]]);
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

  it('format "html" patches the content verbatim through editHtmlMessage', async () => {
    const html = '<table border="1"><tr><td>ok</td></tr></table>';
    const result = await call(client, 'edit_chat_message', {
      chatId: PILOT,
      messageId: 'm1',
      newText: html,
      format: 'html',
    });

    expect(result.isError).toBe(false);
    expect(chats.htmlEdits).toEqual([{ chatId: PILOT, messageId: 'm1', html }]);
    expect(chats.edits).toEqual([]);
  });

  it('mentions are resolved against the chat and forwarded to editMessage', async () => {
    const result = await call(client, 'edit_chat_message', {
      chatId: PILOT,
      messageId: 'm1',
      newText: 'Alice, please re-check',
      mentions: ['Alice'],
    });

    expect(result.isError).toBe(false);
    expect(chats.sentMentions).toEqual([
      [{ name: 'Alice', id: 'aad-alice', displayName: 'Alice Anderson' }],
    ]);
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

describe('pin_chat_message / unpin_chat_message / list_pinned_messages', () => {
  it('pins a message and returns the resulting pinned-list state', async () => {
    const result = await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm1' });
    const payload = result.json() as { pinnedMessages: Array<{ messageId: string }> };

    expect(result.isError).toBe(false);
    expect(payload.pinnedMessages).toEqual([{ id: 'pin-m1', messageId: 'm1', preview: 'preview of m1' }]);
  });

  it('pinning a second message REPLACES the first — the resulting list shows only the new one', async () => {
    await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm1' });
    const second = await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm2' });
    const payload = second.json() as { pinnedMessages: Array<{ messageId: string }> };

    expect(payload.pinnedMessages.map((entry) => entry.messageId)).toEqual(['m2']);
  });

  it('pin refuses a read-only chat', async () => {
    const result = await call(client, 'pin_chat_message', { chatId: WATCHED, messageId: 'w1' });

    expect(result.isError).toBe(true);
  });

  it('pin refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'pin_chat_message', { chatId: OUTSIDE, messageId: 'm1' });

    expect(result.isError).toBe(true);
  });

  it('refuses to claim pinned:true when the post-pin re-list does not actually show the message pinned (review round 2, MINOR 4)', async () => {
    // Graph reporting POST success is not proof the pin landed — only the re-list is. If the
    // target messageId is missing from that list, the outcome must be reported as a loud
    // failure, never as pinned:true on faith in the write's own response.
    chats.pinConfirms = false;

    const result = await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm1' });

    expect(result.isError).toBe(true);
    expect(result.text).toMatch(/not (confirmed|show|pinned)/i);
  });

  it('lists what is currently pinned', async () => {
    await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm1' });
    const result = await call(client, 'list_pinned_messages', { chatId: PILOT });
    const payload = result.json() as { pinnedMessages: Array<{ messageId: string }> };

    expect(result.isError).toBe(false);
    expect(payload.pinnedMessages).toEqual([{ id: 'pin-m1', messageId: 'm1', preview: 'preview of m1' }]);
  });

  it('lists an empty pinnedMessages array when nothing is pinned', async () => {
    const result = await call(client, 'list_pinned_messages', { chatId: PILOT });
    const payload = result.json() as { pinnedMessages: unknown[] };

    expect(result.isError).toBe(false);
    expect(payload.pinnedMessages).toEqual([]);
  });

  it('list works on a read-only chat (readable, not postable)', async () => {
    const result = await call(client, 'list_pinned_messages', { chatId: WATCHED });

    expect(result.isError).toBe(false);
  });

  it('list refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'list_pinned_messages', { chatId: OUTSIDE });

    expect(result.isError).toBe(true);
  });

  it('unpins the currently-pinned message', async () => {
    await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm1' });
    const result = await call(client, 'unpin_chat_message', { chatId: PILOT, messageId: 'm1' });

    expect(result.isError).toBe(false);
    expect(await chats.listPinnedMessages(PILOT)).toEqual([]);
  });

  it('unpin refuses a message that is not the one currently pinned — never silently no-ops', async () => {
    await call(client, 'pin_chat_message', { chatId: PILOT, messageId: 'm1' });
    const result = await call(client, 'unpin_chat_message', { chatId: PILOT, messageId: 'm2' });

    expect(result.isError).toBe(true);
    expect(result.text).toMatch(/not currently pinned/);
  });

  it('unpin refuses a read-only chat', async () => {
    const result = await call(client, 'unpin_chat_message', { chatId: WATCHED, messageId: 'w1' });

    expect(result.isError).toBe(true);
  });

  it('unpin refuses a chat outside the allowlist', async () => {
    const result = await call(client, 'unpin_chat_message', { chatId: OUTSIDE, messageId: 'm1' });

    expect(result.isError).toBe(true);
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

  it('members stays a plain array of display names — AAD ids (needed only for @mentions) are not part of this tool\'s output', async () => {
    const payload = (await call(client, 'list_chats')).json() as {
      chats: Array<{ id: string; members?: string[] }>;
    };

    expect(payload.chats.find((chat) => chat.id === PILOT)?.members).toEqual([
      'Alice Anderson',
      'Bob Brown',
    ]);
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

  it('format "html" posts the content verbatim through sendHtmlMessage, untouched by textToHtml', async () => {
    const html = '<b>bold</b> &amp; <span style="color:#c00">red</span>';
    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: html, format: 'html' });

    expect(result.isError).toBe(false);
    expect(chats.sentHtml).toEqual([{ chatId: PILOT, html }]);
    expect(chats.sent).toEqual([]); // the plain-text path was never touched
  });

  it('format omitted still goes through the escaping sendMessage path', async () => {
    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: '<b>not bold</b>' });

    expect(result.isError).toBe(false);
    expect(chats.sent).toEqual([{ chatId: PILOT, text: '<b>not bold</b>' }]);
    expect(chats.sentHtml).toEqual([]);
  });

  it('mentions are resolved against the chat and forwarded to sendMessage', async () => {
    const result = await call(client, 'send_chat_message', {
      chatId: PILOT,
      text: 'Alice can you take this?',
      mentions: ['Alice'],
    });

    expect(result.isError).toBe(false);
    expect(chats.resolveMentionsCalls).toEqual([{ chatId: PILOT, names: ['Alice'] }]);
    expect(chats.sentMentions).toEqual([
      [{ name: 'Alice', id: 'aad-alice', displayName: 'Alice Anderson' }],
    ]);
  });

  it('no mentions given: resolveMentions is never called and an empty array is forwarded', async () => {
    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: 'Hi' });

    expect(result.isError).toBe(false);
    expect(chats.resolveMentionsCalls).toEqual([]);
    expect(chats.sentMentions).toEqual([[]]);
  });

  it('an unresolvable mention name refuses the whole call — never a silent drop', async () => {
    const result = await call(client, 'send_chat_message', {
      chatId: PILOT,
      text: 'Nobody knows this',
      mentions: ['Nobody'],
    });

    expect(result.isError).toBe(true);
    expect(result.text).toMatch(/No chat member matches/);
    expect(chats.sent).toEqual([]); // refused before ever reaching the send
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

describe('guard() — every MCP tool failure renders Retry-After the same way the CLI does (0.4.1 review round 2)', () => {
  // Round 1 fixed the CLI's error text but never touched guard() (server.ts), which only ever
  // read `error.name`/`error.message` — an agent driving the MCP server got a throttle refusal
  // with NO wait time at all, a regression vs main (main's embedded-in-message wording at least
  // named a number, however duplicated). retryAfterSuffix (graph-client.ts) is now the one place
  // both guard() and the CLI's formatCliError render this from, so this test exercises the MCP
  // tool path specifically, not the CLI.
  it('a Graph 429 with a named Retry-After: the MCP tool result names it, agent-readable', async () => {
    chats.sendMessage = async () => {
      throw new GraphError('Too many requests', 429, 'TooManyRequests', 23);
    };

    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: 'Hi' });

    expect(result.isError).toBe(true);
    expect(result.text).toMatch(/throttled, retry after 23s/);
  });

  it('a Graph 429 with NO named Retry-After: no invented number in the MCP tool result', async () => {
    chats.sendMessage = async () => {
      throw new GraphError('Too many requests', 429, 'TooManyRequests');
    };

    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: 'Hi' });

    expect(result.isError).toBe(true);
    expect(result.text).not.toMatch(/throttled, retry after/);
    expect(result.text).toMatch(/Too many requests/);
  });

  it('a non-429 failure carries no throttle phrasing at all', async () => {
    chats.sendMessage = async () => {
      throw new Error('network unreachable');
    };

    const result = await call(client, 'send_chat_message', { chatId: PILOT, text: 'Hi' });

    expect(result.isError).toBe(true);
    expect(result.text).not.toMatch(/throttled, retry after/);
    expect(result.text).toMatch(/network unreachable/);
  });
});
