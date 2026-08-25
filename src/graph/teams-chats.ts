import { GraphClient, GraphError } from './graph-client.js';
import { textToHtml } from '../formatting.js';
import { type ChatMessage, type ReadResult, applyWatermark, toChatMessage } from '../messages.js';

export interface ChatSummary {
  id: string;
  topic: string;
  chatType: string;
  lastUpdatedDateTime?: string;
  members: string[];
}

export interface AttachmentPayload {
  bytes: Uint8Array;
  contentType: string;
  name: string;
}

export interface OutboundImage {
  bytes: Uint8Array;
  /** Only PNG and JPEG are worth supporting; Teams renders both inline. */
  contentType: 'image/png' | 'image/jpeg';
}

export interface OutboundFile {
  bytes: Uint8Array;
  name: string;
  contentType?: string;
}

/**
 * Secondary port for Teams chats. Everything above it (the MCP tools) speaks ChatSummary and
 * ChatMessage, never Graph JSON, so a different backing API or a test double swaps in here.
 */
export interface TeamsChatsPort {
  listChats(): Promise<ChatSummary[]>;
  readMessages(chatId: string, since?: string, limit?: number): Promise<ReadResult>;
  sendMessage(chatId: string, text: string): Promise<ChatMessage>;
  /**
   * Opt-in raw path: html is posted as the Graph body content VERBATIM — no textToHtml, no
   * escaping. The caller owns entity-escaping `<`, `>`, `&` inside their own content; this is
   * what lets Teams' HTML subset (tables, headings, colour — see the teams-styling skill's
   * verified vocabulary) actually render, since textToHtml only ever produces plain
   * paragraphs/breaks/links.
   */
  sendHtmlMessage(chatId: string, html: string): Promise<ChatMessage>;
  sendImage(chatId: string, image: OutboundImage, text?: string): Promise<ChatMessage>;
  sendFile(chatId: string, file: OutboundFile, text?: string): Promise<ChatMessage>;
  replyToMessage(chatId: string, replyToMessageId: string, text: string): Promise<ChatMessage>;
  editMessage(chatId: string, messageId: string, newText: string): Promise<void>;
  /** Same verbatim contract as sendHtmlMessage, applied to an edit. */
  editHtmlMessage(chatId: string, messageId: string, html: string): Promise<void>;
  deleteMessage(chatId: string, messageId: string): Promise<void>;
  setReaction(chatId: string, messageId: string, reactionType: string): Promise<void>;
  getAttachment(chatId: string, messageId: string, attachmentId?: string): Promise<AttachmentPayload>;
}

/**
 * Graph's sharing-URL encoding: base64 of the absolute URL, made URL-safe, prefixed "u!".
 * A file shared into a chat lives in SharePoint/OneDrive and its contentUrl wants browser
 * cookies, not a Graph bearer token — fetching it directly answers 401. The /shares facade
 * resolves the same URL into a driveItem that the Graph token *can* download.
 */
function shareIdFor(url: string): string {
  const base64 = Buffer.from(url, 'utf8').toString('base64');
  return `u!${base64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_')}`;
}

interface GraphDriveItem {
  id?: string;
  name?: string;
  eTag?: string;
  webUrl?: string;
}

export interface GraphTeamsChatsOptions {
  /** OneDrive folder (under the account's drive root) where outbound files are uploaded. */
  uploadDir?: string;
}

export const DEFAULT_UPLOAD_DIR = 'ai-test';

interface GraphChat {
  id?: string;
  topic?: string | null;
  chatType?: string;
  lastUpdatedDateTime?: string;
  members?: Array<{ displayName?: string | null; email?: string | null }> | null;
}

export class GraphTeamsChats implements TeamsChatsPort {
  private readonly uploadDir: string;

  constructor(
    private readonly graph: GraphClient,
    options: GraphTeamsChatsOptions = {},
  ) {
    this.uploadDir = options.uploadDir?.trim() || DEFAULT_UPLOAD_DIR;
  }

  async listChats(): Promise<ChatSummary[]> {
    const chats = await this.graph.getAll<GraphChat>('/me/chats?$expand=members&$top=50');
    return chats.flatMap((chat) => {
      if (!chat.id) {
        return [];
      }
      const members = (chat.members ?? []).flatMap((member) =>
        member?.displayName ? [member.displayName] : [],
      );
      return [
        {
          id: chat.id,
          // Group chats often have no topic; the member list is the only way to recognise them.
          topic: chat.topic ?? members.join(', ') ?? '',
          chatType: chat.chatType ?? 'unknown',
          ...(chat.lastUpdatedDateTime ? { lastUpdatedDateTime: chat.lastUpdatedDateTime } : {}),
          members,
        },
      ];
    });
  }

  async readMessages(chatId: string, since?: string, limit = 50): Promise<ReadResult> {
    const raw = await this.graph.getAll<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages?$top=${Math.min(limit, 50)}`,
      limit,
    );
    return applyWatermark(
      raw.map((message) => toChatMessage(message, chatId)),
      since,
    );
  }

  async sendMessage(chatId: string, text: string): Promise<ChatMessage> {
    // Always HTML: a 'text' body renders in Teams as one unbroken blob — no line breaks, no
    // clickable links. textToHtml escapes everything, so plain text stays plain.
    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      { body: { contentType: 'html', content: textToHtml(text) } },
    );
    return toChatMessage(created, chatId);
  }

  async sendHtmlMessage(chatId: string, html: string): Promise<ChatMessage> {
    // No textToHtml here — html IS the body content, posted exactly as given. See
    // TeamsChatsPort.sendHtmlMessage's doc comment for why this exists alongside sendMessage.
    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      { body: { contentType: 'html', content: html } },
    );
    return toChatMessage(created, chatId);
  }

  async sendImage(chatId: string, image: OutboundImage, text?: string): Promise<ChatMessage> {
    // The temporary-id pattern: the img src refers to hostedContents entry "1" by its
    // temporaryId, and Graph rewrites the src to the real hosted content URL on creation.
    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: {
          contentType: 'html',
          content:
            (text ? textToHtml(text) : '') +
            '<img src="../hostedContents/1/$value" alt="image">',
        },
        hostedContents: [
          {
            '@microsoft.graph.temporaryId': '1',
            contentBytes: Buffer.from(image.bytes).toString('base64'),
            contentType: image.contentType,
          },
        ],
      },
    );
    return toChatMessage(created, chatId);
  }

  async sendFile(chatId: string, file: OutboundFile, text?: string): Promise<ChatMessage> {
    // A chat cannot host a real file; it has to live in the sender's OneDrive first, and the
    // message then carries a reference attachment pointing at it.
    const folder = this.uploadDir.split('/').map(encodeURIComponent).join('/');
    const item = await this.graph.putBinary<GraphDriveItem>(
      `/me/drive/root:/${folder}/${encodeURIComponent(file.name)}:/content`,
      file.bytes,
      file.contentType ?? 'application/octet-stream',
    );

    // The attachment id must be the GUID inside the driveItem's eTag ("{GUID},n") — the
    // driveItem id itself is a different identifier and Teams rejects it.
    const guid = /\{([0-9a-fA-F-]+)\}/.exec(item.eTag ?? '')?.[1];
    if (!guid || !item.webUrl) {
      throw new Error(
        `OneDrive upload of ${file.name} returned no usable eTag/webUrl; ` +
          'cannot build the reference attachment.',
      );
    }

    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: {
          contentType: 'html',
          content:
            (text ? textToHtml(text) : '') +
            `<attachment id="${guid}"></attachment>`,
        },
        attachments: [
          {
            id: guid,
            contentType: 'reference',
            contentUrl: item.webUrl,
            name: item.name ?? file.name,
          },
        ],
      },
    );
    return toChatMessage(created, chatId);
  }

  /**
   * One message by id. The single-message endpoint is throttled on its own budget — on
   * 2026-08-25 it answered 429 for hours while the list endpoint for the same chat stayed
   * healthy, which silently broke every quoted reply and every attachment download. So a 429
   * there falls back to scanning the recent list for the id: a page costs one healthy request
   * and the message we want is almost always among the last fifty.
   */
  private async fetchMessage(chatId: string, messageId: string): Promise<ChatMessage> {
    try {
      return toChatMessage(
        await this.graph.get<unknown>(
          `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
          { readRetries: 0 }, // the list fallback below is cheaper than sleeping and re-hitting a throttled family
        ),
        chatId,
      );
    } catch (caught) {
      if (!(caught instanceof GraphError) || caught.status !== 429) {
        throw caught;
      }
      const recent = await this.graph.get<{ value?: unknown[] }>(
        `/chats/${encodeURIComponent(chatId)}/messages?$top=50`,
      );
      const found = (recent.value ?? [])
        .map((raw) => toChatMessage(raw, chatId))
        .find((message) => message.id === messageId);
      if (!found) {
        throw new GraphError(
          `Message ${messageId} could not be fetched (that endpoint is throttled` +
            `${caught.retryAfterSeconds ? `, retry in ${caught.retryAfterSeconds}s` : ''}) and is not ` +
            'among the chat\'s last 50 messages — nothing was posted.',
          429,
          'MessageFetchThrottled',
          caught.retryAfterSeconds,
        );
      }
      return found;
    }
  }

  async replyToMessage(
    chatId: string,
    replyToMessageId: string,
    text: string,
  ): Promise<ChatMessage> {
    // Chats have no reply threads (that is channels-only). What the Teams UI calls a reply in a
    // chat is a new message carrying a messageReference attachment - a quote card built from the
    // original message - so that is what gets posted here.
    const original = await this.fetchMessage(chatId, replyToMessageId);

    const reference = JSON.stringify({
      messageId: original.id,
      messagePreview: original.text.slice(0, 150),
      messageSender: {
        application: null,
        device: null,
        user: {
          userIdentityType: 'aadUser',
          id: original.fromId ?? null,
          displayName: original.from,
        },
      },
    });

    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: {
          contentType: 'html',
          content: `<attachment id="${original.id}"></attachment>${textToHtml(text)}`,
        },
        attachments: [
          { id: original.id, contentType: 'messageReference', content: reference },
        ],
      },
    );
    return toChatMessage(created, chatId);
  }

  async editMessage(chatId: string, messageId: string, newText: string): Promise<void> {
    // Graph only lets the delegated user edit messages they sent themselves; anything else
    // comes back as a Graph error, which the caller sees verbatim.
    await this.graph.patch(
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
      { body: { contentType: 'html', content: textToHtml(newText) } },
    );
  }

  async editHtmlMessage(chatId: string, messageId: string, html: string): Promise<void> {
    // Same verbatim contract as sendHtmlMessage — no textToHtml, no escaping.
    await this.graph.patch(
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
      { body: { contentType: 'html', content: html } },
    );
  }

  async deleteMessage(chatId: string, messageId: string): Promise<void> {
    // The reversible soft delete — the message becomes "This message was deleted" in Teams and
    // can be restored. Hard delete is deliberately not offered here.
    await this.graph.postAction(
      `/me/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/softDelete`,
    );
  }

  async setReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    // The payload is {reactionType} at the top level — NOT wrapped in {body: …} like a message
    // post. Graph answers 204 No Content, hence postNoContent. Both halves of this comment are
    // paid-for knowledge: the 2026-08-24 ad-hoc script got each of them wrong in turn.
    await this.graph.postNoContent(
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/setReaction`,
      { reactionType },
    );
  }

  async getAttachment(
    chatId: string,
    messageId: string,
    attachmentId?: string,
  ): Promise<AttachmentPayload> {
    const parsed = await this.fetchMessage(chatId, messageId);
    // Without an explicit id, "the attachment" means the first one that IS a file — a quoted
    // reply carries a messageReference card first, and a caller asking for "the attachment" on
    // "Logo: <file>" never means the quote (2026-08-25: that pick answered 404 from hostedContents).
    const attachment = attachmentId
      ? parsed.attachments.find((candidate) => candidate.id === attachmentId)
      : parsed.attachments.find((candidate) => candidate.contentType !== 'messageReference');

    if (!attachment) {
      const onlyQuote = parsed.attachments.length > 0;
      throw new Error(
        attachmentId
          ? `Message ${messageId} has no attachment ${attachmentId}.`
          : onlyQuote
            ? `Message ${messageId} has no downloadable attachment — only a quoted-message card.`
            : `Message ${messageId} has no attachments.`,
      );
    }

    // Two shapes exist. A pasted image is hosted content on the message itself; a shared file
    // is a driveItem in SharePoint/OneDrive whose contentUrl must be resolved through the
    // /shares facade — see shareIdFor.
    if (!attachment.contentUrl) {
      const hosted = await this.graph.getBinary(
        `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}` +
          `/hostedContents/${encodeURIComponent(attachment.id)}/$value`,
      );
      return { ...hosted, name: attachment.name ?? `${attachment.id}.bin` };
    }

    const downloaded = await this.graph.getBinary(
      `/shares/${shareIdFor(attachment.contentUrl)}/driveItem/content`,
    );
    // attachment.contentType is usually the literal "reference", not a media type; the download
    // response knows better.
    const declared = attachment.contentType?.includes('/') ? attachment.contentType : undefined;
    return {
      bytes: downloaded.bytes,
      contentType: declared ?? downloaded.contentType,
      name: attachment.name ?? `${attachment.id}.bin`,
    };
  }
}

export type { ChatMessage, ReadResult };
