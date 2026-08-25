import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, extname, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { z } from 'zod';
import type { ChatAllowlist } from './allowlist.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';

export interface ServerDeps {
  chats: TeamsChatsPort;
  allowlist: ChatAllowlist;
  assistantDisplayName: string;
  downloadDir?: string;
}

interface ToolResult {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
  // The SDK's CallToolResult carries an index signature for protocol extensions.
  [key: string]: unknown;
}

function ok(payload: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
}

/**
 * Every failure reaches the model as text rather than a transport error, including an allowlist
 * refusal — the model needs to read why it was refused and stop, not retry blindly.
 */
function guard(handler: () => Promise<ToolResult>): Promise<ToolResult> {
  return handler().catch((error: unknown) => ({
    content: [
      {
        type: 'text' as const,
        text: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
      },
    ],
    isError: true,
  }));
}

export function buildServer(deps: ServerDeps): McpServer {
  const { chats, allowlist } = deps;
  const downloadDir = deps.downloadDir ?? join(tmpdir(), 'teams-assistant-mcp');

  const server = new McpServer(
    { name: 'teams-assistant-mcp', version: '0.1.0' },
    {
      instructions:
        `Reads and posts in a fixed set of Microsoft Teams group chats as the account ` +
        `"${deps.assistantDisplayName}". Only the chats returned by list_chats can be used; any ` +
        `other chat id is refused. Anything you post is visible to real colleagues under an ` +
        `account whose name says it is an AI assistant, so write as that assistant would.`,
    },
  );

  server.registerTool(
    'list_chats',
    {
      title: 'List allowed Teams chats',
      description:
        'Lists the Teams chats this server may touch. The result is the allowlist intersected ' +
        'with what the signed-in account can actually see, so a chat missing here means either ' +
        'it is not allowlisted or the account is not a member. Never lists other chats.',
      inputSchema: {},
      annotations: { readOnlyHint: true },
    },
    () =>
      guard(async () => {
        const visible = await chats.listChats();
        const byId = new Map(visible.map((chat) => [chat.id, chat]));
        const allowed = allowlist.entries().map((entry) => {
          const chat = byId.get(entry.id);
          return {
            id: entry.id,
            label: entry.label,
            canPost: entry.canPost,
            visibleToAccount: chat !== undefined,
            ...(chat
              ? {
                  topic: chat.topic,
                  chatType: chat.chatType,
                  members: chat.members,
                  ...(chat.lastUpdatedDateTime
                    ? { lastUpdatedDateTime: chat.lastUpdatedDateTime }
                    : {}),
                }
              : {}),
          };
        });
        return ok({ chats: allowed });
      }),
  );

  server.registerTool(
    'read_chat_messages',
    {
      title: 'Read Teams chat messages',
      description:
        'Reads messages from one allowlisted chat, oldest first. Pass the watermark from the ' +
        'previous call as `since` to get only what arrived after it; the watermark is exclusive. ' +
        'When nothing is new, `messages` is empty and no watermark is returned, so keep the one ' +
        'you already had.',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be on the allowlist'),
        since: z
          .string()
          .optional()
          .describe('ISO-8601 watermark from a previous read; exclusive'),
        limit: z.number().int().min(1).max(200).optional().describe('Max messages, default 50'),
      },
      annotations: { readOnlyHint: true },
    },
    ({ chatId, since, limit }) =>
      guard(async () => {
        allowlist.assertReadable(chatId);
        const result = await chats.readMessages(chatId, since, limit ?? 50);
        return ok({ chatId, ...result });
      }),
  );

  server.registerTool(
    'send_chat_message',
    {
      title: 'Post a message to a Teams chat',
      description:
        'Posts a message to an allowlisted chat that has canPost enabled. The text is always ' +
        'posted as HTML rendered from it: blank lines become paragraph breaks, single newlines ' +
        'become line breaks, http(s) URLs become clickable links, everything else is escaped ' +
        'verbatim — no markdown. Real people read this immediately and it cannot be unsent ' +
        'through this server.',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        text: z
          .string()
          .min(1)
          .describe('Plain text; newlines and URLs render properly, markdown does not'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false },
    },
    ({ chatId, text }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        const sent = await chats.sendMessage(chatId, text);
        return ok({ posted: true, chatId, messageId: sent.id, createdDateTime: sent.createdDateTime });
      }),
  );

  server.registerTool(
    'reply_chat_message',
    {
      title: 'Reply to a specific message in a Teams chat',
      description:
        'Posts a reply that quotes an existing message, the way the Teams UI does it. Chats have ' +
        'no real reply threads (that is channels-only): the reply is an ordinary new message ' +
        'carrying a quote card of the original, so it appears at the bottom of the chat like any ' +
        'other message. Use it when a bare message would lose which question it answers. Same ' +
        'rules and rendering as send_chat_message: allowlisted chat with canPost, visible ' +
        'immediately, text always posted as HTML (newlines and links render, markdown does not).',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        replyToMessageId: z.string().describe('Id of the message being answered'),
        text: z
          .string()
          .min(1)
          .describe('Plain text of the reply; newlines and URLs render properly, markdown does not'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false },
    },
    ({ chatId, replyToMessageId, text }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        const sent = await chats.replyToMessage(chatId, replyToMessageId, text);
        return ok({
          posted: true,
          chatId,
          messageId: sent.id,
          repliedTo: replyToMessageId,
          createdDateTime: sent.createdDateTime,
        });
      }),
  );

  server.registerTool(
    'edit_chat_message',
    {
      title: 'Edit a message this account sent',
      description:
        'Replaces the text of a message in an allowlisted chat. Graph only allows editing ' +
        'messages the signed-in account sent itself; trying anyone else\'s message returns ' +
        'Graph\'s own refusal. Teams shows the message with an "Edited" marker. The new text is ' +
        'always posted as HTML, same rendering as send_chat_message (newlines and links render, ' +
        'markdown does not).',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        messageId: z.string().describe('Id of a message this account sent'),
        newText: z
          .string()
          .min(1)
          .describe('Replacement plain text; newlines and URLs render properly, markdown does not'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
    },
    ({ chatId, messageId, newText }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        await chats.editMessage(chatId, messageId, newText);
        return ok({ edited: true, chatId, messageId });
      }),
  );

  server.registerTool(
    'react_to_chat_message',
    {
      title: 'React to a message with an emoji',
      description:
        'Puts an emoji reaction (👍, ❤️, …) on a message in an allowlisted chat, as the ' +
        'signed-in account. The receipt gesture for "seen and being handled" — cheaper than a ' +
        'reply and idempotent: reacting twice with the same emoji is one reaction.',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        messageId: z.string().describe('Id of the message to react to'),
        emoji: z.string().min(1).describe('The reaction emoji, e.g. 👍'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
    },
    ({ chatId, messageId, emoji }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        await chats.setReaction(chatId, messageId, emoji);
        return ok({ reacted: true, chatId, messageId, emoji });
      }),
  );

  server.registerTool(
    'delete_chat_message',
    {
      title: 'Delete a message this account sent',
      description:
        'Soft-deletes a message in an allowlisted chat - the reversible kind, leaving the ' +
        '"This message was deleted" stub that Teams can restore. Graph only allows deleting ' +
        'messages the signed-in account sent itself. There is deliberately no hard delete here.',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        messageId: z.string().describe('Id of a message this account sent'),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true },
    },
    ({ chatId, messageId }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        await chats.deleteMessage(chatId, messageId);
        return ok({ deleted: true, chatId, messageId });
      }),
  );

  const IMAGE_MIME_BY_EXT: Record<string, 'image/png' | 'image/jpeg'> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
  };

  server.registerTool(
    'send_chat_image',
    {
      title: 'Post an image to a Teams chat',
      description:
        'Posts a PNG or JPEG that renders inline in an allowlisted chat with canPost enabled. ' +
        'Give either a local file path or base64 bytes; base64 needs an explicit mime. Real ' +
        'people see this immediately and it cannot be unsent through this server.',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        path: z.string().optional().describe('Local path to a .png/.jpg/.jpeg file'),
        base64: z.string().optional().describe('Raw image bytes, base64-encoded'),
        mime: z
          .enum(['image/png', 'image/jpeg'])
          .optional()
          .describe('Image type; required with base64, inferred from the extension for a path'),
        text: z
          .string()
          .optional()
          .describe('Caption shown above the image; rendered like send_chat_message text'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false },
    },
    ({ chatId, path, base64, mime, text }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        if ((path === undefined) === (base64 === undefined)) {
          throw new Error('Give exactly one of path or base64.');
        }
        let bytes: Uint8Array;
        let contentType = mime;
        if (path !== undefined) {
          bytes = new Uint8Array(await readFile(path));
          contentType ??= IMAGE_MIME_BY_EXT[extname(path).toLowerCase()];
          if (!contentType) {
            throw new Error(`Cannot tell the image type of ${basename(path)}; pass mime.`);
          }
        } else {
          bytes = new Uint8Array(Buffer.from(base64 as string, 'base64'));
          if (!contentType) {
            throw new Error('base64 input needs an explicit mime (image/png or image/jpeg).');
          }
        }
        if (bytes.byteLength === 0) {
          throw new Error('The image is empty.');
        }
        const sent = await chats.sendImage(chatId, { bytes, contentType }, text);
        return ok({
          posted: true,
          chatId,
          messageId: sent.id,
          createdDateTime: sent.createdDateTime,
          bytes: bytes.byteLength,
          contentType,
        });
      }),
  );

  server.registerTool(
    'send_chat_file',
    {
      title: 'Share a file into a Teams chat',
      description:
        'Uploads a local file to the assistant account\'s OneDrive and shares it into an ' +
        'allowlisted chat with canPost enabled, as a normal Teams file attachment. Real people ' +
        'see this immediately and it cannot be unsent through this server.',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be allowlisted with canPost: true'),
        path: z.string().describe('Local path of the file to share'),
        text: z
          .string()
          .optional()
          .describe('Caption shown above the file card; rendered like send_chat_message text'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false },
    },
    ({ chatId, path, text }) =>
      guard(async () => {
        allowlist.assertPostable(chatId);
        const bytes = new Uint8Array(await readFile(path));
        const name = basename(path);
        const sent = await chats.sendFile(chatId, { bytes, name }, text);
        return ok({
          posted: true,
          chatId,
          messageId: sent.id,
          createdDateTime: sent.createdDateTime,
          name,
          bytes: bytes.byteLength,
        });
      }),
  );

  server.registerTool(
    'get_chat_attachment',
    {
      title: 'Download a Teams chat attachment',
      description:
        'Downloads an attachment from a message in an allowlisted chat and writes it to a local ' +
        'file, returning the path. Omit attachmentId to take the first attachment on the message. ' +
        'Images pasted into a message show up as attachments named inline-image-1, inline-image-2, …',
      inputSchema: {
        chatId: z.string().describe('Graph chat id, must be on the allowlist'),
        messageId: z.string().describe('Message id from read_chat_messages'),
        attachmentId: z.string().optional().describe('Attachment id; defaults to the first one'),
      },
      annotations: { readOnlyHint: true },
    },
    ({ chatId, messageId, attachmentId }) =>
      guard(async () => {
        allowlist.assertReadable(chatId);
        const attachment = await chats.getAttachment(chatId, messageId, attachmentId);
        await mkdir(downloadDir, { recursive: true });
        // basename strips any path the sender put in the file name; the download must not be
        // able to escape downloadDir.
        const path = resolve(join(downloadDir, `${messageId}-${basename(attachment.name)}`));
        await writeFile(path, attachment.bytes);
        return ok({
          path,
          name: attachment.name,
          contentType: attachment.contentType,
          bytes: attachment.bytes.byteLength,
        });
      }),
  );

  server.registerTool(
    'poll_chats',
    {
      title: 'Poll every allowed chat for new messages',
      description:
        'Reads all allowlisted chats in one call. Pass the watermarks object returned last time ' +
        'to get only new messages; the returned watermarks object replaces it. Chats with ' +
        'nothing new are omitted from messagesByChat but keep their watermark.',
      inputSchema: {
        watermarks: z
          .record(z.string())
          .optional()
          .describe('chatId -> ISO-8601 watermark from the previous poll'),
        limit: z.number().int().min(1).max(200).optional().describe('Max messages per chat'),
      },
      annotations: { readOnlyHint: true },
    },
    ({ watermarks, limit }) =>
      guard(async () => {
        const previous = watermarks ?? {};
        const messagesByChat: Record<string, unknown[]> = {};
        const next: Record<string, string> = { ...previous };
        const failures: Array<{ chatId: string; error: string }> = [];

        for (const entry of allowlist.entries()) {
          try {
            const result = await chats.readMessages(entry.id, previous[entry.id], limit ?? 50);
            if (result.messages.length > 0) {
              messagesByChat[entry.id] = result.messages;
            }
            if (result.watermark) {
              next[entry.id] = result.watermark;
            }
          } catch (error) {
            // One unreachable chat must not blank out the poll for the others.
            failures.push({
              chatId: entry.id,
              error: error instanceof Error ? error.message : String(error),
            });
          }
        }

        return ok({
          messagesByChat,
          watermarks: next,
          ...(failures.length > 0 ? { failures } : {}),
        });
      }),
  );

  return server;
}
