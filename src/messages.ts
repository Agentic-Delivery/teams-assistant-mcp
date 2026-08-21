export interface ChatAttachmentRef {
  id: string;
  name?: string;
  contentType?: string;
  contentUrl?: string;
  /**
   * Inline payload some attachment kinds carry instead of a URL — for a messageReference (a
   * quoted reply) it is the JSON string with the quoted message's id, preview and sender.
   */
  content?: string;
}

export interface ChatMessage {
  id: string;
  chatId: string;
  /** ISO-8601, straight from Graph. This doubles as the watermark. */
  createdDateTime: string;
  lastModifiedDateTime?: string;
  from: string;
  fromId?: string;
  text: string;
  isDeleted: boolean;
  attachments: ChatAttachmentRef[];
}

export interface ReadResult {
  messages: ChatMessage[];
  /**
   * Pass this back as `since` on the next call to get only what arrived after this batch.
   * Undefined when nothing was returned, so the caller keeps the watermark it already had.
   */
  watermark?: string;
}

// Only the *closing* block tag becomes a newline. Breaking on the opening tag too would turn
// every ordinary two-paragraph message into a double-spaced one.
const BLOCK_END = /(?:<\/(?:p|div|li|tr|h[1-6])\s*>|<br\b[^>]*>)/gi;
const ANY_TAG = /<[^>]*>/g;

const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  '#39': "'",
  nbsp: ' ',
};

/**
 * Teams messages arrive as HTML even when the user typed plain text. The orchestrator reads these
 * as text, so tags become newlines or disappear. This is not a general HTML renderer and does not
 * need to be — mentions, emoji and images degrade to their text content or to nothing.
 */
export function htmlToText(html: string): string {
  return html
    .replace(/<\s*(script|style)\b[^>]*>[\s\S]*?<\/\s*\1\s*>/gi, '')
    // Raw newlines in HTML source are ordinary whitespace, not line breaks — and Teams
    // pretty-prints stored bodies with them (after <br>, between <p> elements). Line breaks in
    // the text come only from the tags handled below.
    .replace(/[\r\n]+/g, ' ')
    .replace(BLOCK_END, '\n')
    .replace(ANY_TAG, '')
    .replace(/&(#?\w+);/g, (match, entity: string) => {
      const key = entity.toLowerCase();
      if (ENTITIES[key] !== undefined) {
        return ENTITIES[key];
      }
      if (/^#\d+$/.test(entity)) {
        return String.fromCodePoint(Number(entity.slice(1)));
      }
      return match;
    })
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

interface GraphMessage {
  id?: string;
  chatId?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string | null;
  messageType?: string;
  from?: {
    user?: { id?: string; displayName?: string } | null;
    application?: { id?: string; displayName?: string } | null;
  } | null;
  body?: { contentType?: string; content?: string } | null;
  attachments?: Array<{
    id?: string;
    name?: string | null;
    contentType?: string | null;
    contentUrl?: string | null;
    content?: string | null;
  }> | null;
}

// A pasted image is not an attachment in Graph's eyes: it only exists as an <img> in the body
// HTML pointing at the message's hostedContents collection. Matching the src is what makes it
// discoverable at all.
const HOSTED_IMG_SRC = /<img\b[^>]*\bsrc=["'][^"']*\/hostedContents\/([^/"']+)\/\$value["']/gi;

/**
 * Pseudo-attachment refs for images pasted into the message body. They carry no contentUrl, so
 * downloading one goes through the hostedContents endpoint, same as any hosted attachment. The
 * real media type is only known once downloaded, hence the wildcard.
 */
export function inlineImageRefs(html: string): ChatAttachmentRef[] {
  return [...html.matchAll(HOSTED_IMG_SRC)].flatMap((match, index) => {
    const encoded = match[1];
    if (!encoded) {
      return [];
    }
    let id = encoded;
    try {
      id = decodeURIComponent(encoded);
    } catch {
      // Not URL-encoded after all; use it as-is.
    }
    return [{ id, contentType: 'image/*', name: `inline-image-${index + 1}` }];
  });
}

export function toChatMessage(raw: unknown, fallbackChatId: string): ChatMessage {
  const message = (raw ?? {}) as GraphMessage;
  const sender = message.from?.user ?? message.from?.application ?? undefined;
  const body = message.body ?? {};
  const content = body.content ?? '';

  return {
    id: message.id ?? '',
    chatId: message.chatId ?? fallbackChatId,
    createdDateTime: message.createdDateTime ?? '',
    ...(message.lastModifiedDateTime ? { lastModifiedDateTime: message.lastModifiedDateTime } : {}),
    // A system message (someone joined, chat renamed) has no user; label it rather than blank.
    from: sender?.displayName ?? (message.messageType === 'systemEventMessage' ? 'system' : 'unknown'),
    ...(sender?.id ? { fromId: sender.id } : {}),
    text: body.contentType === 'html' ? htmlToText(content) : content.trim(),
    isDeleted: Boolean(message.deletedDateTime),
    attachments: [
      ...(message.attachments ?? []).flatMap((attachment) =>
        attachment?.id
          ? [
              {
                id: attachment.id,
                ...(attachment.name ? { name: attachment.name } : {}),
                ...(attachment.contentType ? { contentType: attachment.contentType } : {}),
                ...(attachment.contentUrl ? { contentUrl: attachment.contentUrl } : {}),
                ...(attachment.content ? { content: attachment.content } : {}),
              },
            ]
          : [],
      ),
      ...(body.contentType === 'html' ? inlineImageRefs(content) : []),
    ],
  };
}

function timestamp(value: string): number {
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? Number.NEGATIVE_INFINITY : parsed;
}

/**
 * Graph returns chat messages newest-first and its $filter on createdDateTime is unreliable for
 * chats, so the watermark is applied here instead. `since` is exclusive: a message created at
 * exactly the watermark was already delivered.
 */
export function applyWatermark(messages: ChatMessage[], since?: string): ReadResult {
  const ordered = [...messages].sort(
    (a, b) => timestamp(a.createdDateTime) - timestamp(b.createdDateTime),
  );

  const cutoff = since ? timestamp(since) : Number.NEGATIVE_INFINITY;
  if (since && cutoff === Number.NEGATIVE_INFINITY) {
    throw new Error(`since must be an ISO-8601 timestamp, got "${since}".`);
  }

  const fresh = ordered.filter((message) => timestamp(message.createdDateTime) > cutoff);
  const newest = fresh.at(-1);

  return newest ? { messages: fresh, watermark: newest.createdDateTime } : { messages: fresh };
}
