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

// Teams renders a pasted link as <a href="URL" title="URL">LABEL</a>. LABEL is often a page
// title or link-preview text, not the URL itself — everything ANY_TAG below would keep, silently
// dropping the href. Matched (and rewritten to "LABEL (URL)") before the generic tag strip runs,
// while the href attribute is still present. Requires a quoted href (Teams always quotes it); an
// unquoted href="X" is not matched and degrades to the pre-fix behaviour (label kept, URL lost).
const ANCHOR = /<a\b[^>]*\bhref\s*=\s*(["'])([\s\S]*?)\1[^>]*>([\s\S]*?)<\/a\s*>/gi;

const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  '#39': "'",
  nbsp: ' ',
};

function decodeEntities(text: string): string {
  return text.replace(/&(#?\w+);/g, (match, entity: string) => {
    const key = entity.toLowerCase();
    if (ENTITIES[key] !== undefined) {
      return ENTITIES[key];
    }
    if (/^#\d+$/.test(entity)) {
      return String.fromCodePoint(Number(entity.slice(1)));
    }
    return match;
  });
}

/**
 * href vs label equality check only, never used for the text that ends up in the message: HTML
 * entities and URL percent-encoding are two independent ways the same character can show up (a
 * label shows an entity, an href shows the percent-encoded bytes), so both are undone before
 * comparing. Malformed percent-encoding (not actually a URL, just text containing "%") is left
 * as-is rather than thrown on.
 */
function normalizeForCompare(text: string): string {
  const decoded = decodeEntities(text);
  try {
    return decodeURIComponent(decoded);
  } catch {
    return decoded;
  }
}

/**
 * Teams messages arrive as HTML even when the user typed plain text. The orchestrator reads these
 * as text, so tags become newlines or disappear. This is not a general HTML renderer and does not
 * need to be — mentions, emoji and images degrade to their text content or to nothing.
 */
export function htmlToText(html: string): string {
  const stripped = html
    .replace(/<\s*(script|style)\b[^>]*>[\s\S]*?<\/\s*\1\s*>/gi, '')
    // Raw newlines in HTML source are ordinary whitespace, not line breaks — and Teams
    // pretty-prints stored bodies with them (after <br>, between <p> elements). Line breaks in
    // the text come only from the tags handled below.
    .replace(/[\r\n]+/g, ' ')
    .replace(ANCHOR, (_match, _quote, hrefRaw: string, innerRaw: string) => {
      const href = hrefRaw.trim();
      // A block tag inside the label (rare, but seen with rich-preview cards) would otherwise
      // glue adjacent words together once ANY_TAG below drops it; turn it into a space first,
      // same as BLOCK_END does for the rest of the message.
      const text = innerRaw.replace(BLOCK_END, ' ').replace(ANY_TAG, '').replace(/\s+/g, ' ').trim();
      if (!href) {
        return text;
      }
      // A bare pasted URL comes through as label === href; keep it plain instead of "URL (URL)".
      // Compared after normalising both sides — href and label are encoded independently by
      // Teams (a non-ASCII character in the URL may show as an HTML entity in the label but be
      // percent-encoded in the href), so a raw-string comparison can miss a same-URL match.
      return text && normalizeForCompare(text) !== normalizeForCompare(href) ? `${text} (${href})` : href;
    })
    .replace(BLOCK_END, '\n')
    .replace(ANY_TAG, '');

  return decodeEntities(stripped)
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
