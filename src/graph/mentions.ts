import { escapeHtml, findUrlSpans, textToHtml } from '../formatting.js';

/** One chat member as Graph's `$expand=members`/`/members` shape actually carries it. */
export interface ChatMember {
  /**
   * The AAD user id (Graph's `userId` on a conversation member — NOT the membership `id`,
   * which is a different, composite identifier Graph rejects for a mention). Missing only for a
   * member Graph did not report one for; such a member cannot be @mentioned.
   */
  id?: string;
  displayName: string;
}

/**
 * One resolved @mention: the caller's own search string plus the chat member it matched.
 * `name` is kept (not just id/displayName) because the text-substitution passes need the
 * ORIGINAL string to search for in the caller's text/HTML — the thing typed is not necessarily
 * the member's full displayName ("Shiv" resolves to "Garg, Shivankit", but "Shiv" is what
 * appears in the message).
 */
export interface MentionTarget {
  name: string;
  id: string;
  displayName: string;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * `name` matched only when NOT tight against another letter/number on either side ("Shiv" must
 * not match inside "Shivankit") — Unicode-aware (`\p{L}`/`\p{N}` with the `u` flag) so an accented
 * name like "Spännare" gets correct boundaries too. Built fresh per call: a shared `g`-flagged
 * RegExp instance would carry `lastIndex` state across the separate `.test()` calls this is used
 * for (existence checks over several segments), silently skipping real matches.
 */
function wordBoundaryPattern(name: string, flags: string): RegExp {
  return new RegExp(`(?<![\\p{L}\\p{N}])${escapeRegExp(name)}(?![\\p{L}\\p{N}])`, flags);
}

interface TextSegment {
  /** true when this segment IS a URL — never eligible for mention substitution. */
  isUrl: boolean;
  value: string;
}

/**
 * Splits `text` into alternating URL / non-URL segments using the exact boundaries `textToHtml`
 * will itself wrap in an `<a href>` (findUrlSpans, shared with formatting.ts). A mention name is
 * only ever searched for — and only ever substituted — inside the non-URL segments: a name that
 * merely happens to appear inside a URL's path/query is part of the link's literal text, not a
 * mention, and tokenizing it there used to land <at> markup inside the href attribute value,
 * whose own quote character terminates the attribute early (malformed HTML) — verified 2026-08-26
 * review round with `see https://x.com/shiv/pr Shiv` and mention "Shiv".
 */
function splitOnUrls(text: string): TextSegment[] {
  const segments: TextSegment[] = [];
  let consumed = 0;
  for (const { start, end } of findUrlSpans(text)) {
    if (start > consumed) {
      segments.push({ isUrl: false, value: text.slice(consumed, start) });
    }
    segments.push({ isUrl: true, value: text.slice(start, end) });
    consumed = end;
  }
  if (consumed < text.length) {
    segments.push({ isUrl: false, value: text.slice(consumed) });
  }
  return segments;
}

/**
 * Resolves each of `names` against `members`: case-insensitive substring match against
 * displayName, required to be unambiguous. This is what lets "Shiv" find "Garg, Shivankit"
 * without the caller ever seeing Graph's "Lastname, Firstname" convention.
 *
 * Every failure mode throws a specific, actionable message — no match, an ambiguous match, or a
 * matched member with no recorded AAD id (cannot build a notifying mention for them) — because a
 * mention that silently resolves to nothing is a colleague who never gets notified and nobody
 * finds out. "Never a silent drop" is the whole point of this function existing.
 */
export function resolveMentionTargets(names: readonly string[], members: readonly ChatMember[]): MentionTarget[] {
  return names.map((name) => {
    const query = name.trim().toLowerCase();
    if (query === '') {
      throw new Error('A mention name cannot be empty.');
    }
    const matches = members.filter((member) => member.displayName.toLowerCase().includes(query));
    if (matches.length === 0) {
      throw new Error(
        `No chat member matches mention "${name}". Members: ` +
          `${members.map((member) => member.displayName).join(', ') || '(none visible)'}.`,
      );
    }
    if (matches.length > 1) {
      throw new Error(
        `Mention "${name}" is ambiguous — it matches ${matches.length} chat members: ` +
          `${matches.map((member) => member.displayName).join(', ')}. Use a more specific name.`,
      );
    }
    const member = matches[0] as ChatMember;
    if (!member.id) {
      throw new Error(
        `Chat member "${member.displayName}" has no AAD user id on record, so a Teams mention ` +
          'cannot be built for them (it would post an @-tag that never notifies).',
      );
    }
    return { name, id: member.id, displayName: member.displayName };
  });
}

/** The Graph request-body shape for a chatMessage's `mentions` array. `id` here is the small,
 *  per-message index the <at id="N"> tags reference — unrelated to the AAD user id, which lives
 *  in `mentioned.user.id`. */
export function buildGraphMentionsPayload(
  mentions: readonly MentionTarget[],
): Array<{ id: number; mentionText: string; mentioned: { user: { id: string; displayName: string } } }> {
  return mentions.map((mention, index) => ({
    id: index,
    mentionText: mention.displayName,
    mentioned: { user: { id: mention.id, displayName: mention.displayName } },
  }));
}

// Private-use-area sentinel: not a character any real message text will contain, so it survives
// escapeHtml (which only touches &<>") and the URL matcher untouched, then gets swapped for the
// real <at> tag once textToHtml has finished escaping everything else.
const PLACEHOLDER_GUARD = '\uE000';

/**
 * Renders plain text to the outbound HTML body, same as textToHtml, except every whole-word,
 * case-insensitive occurrence of each mention's `name` — outside any URL — becomes a
 * `<at id="N">displayName</at>` tag that Graph's `mentions` array (buildGraphMentionsPayload)
 * will actually notify.
 *
 * Substitution happens via a placeholder token BEFORE textToHtml runs — inserting the raw <at>
 * markup after escaping would be safe from re-escaping, but doing the name search on already
 * paragraph/br-wrapped HTML risks matching inside markup Teams itself added (an autolinked URL,
 * say). The placeholder is inert to every transform textToHtml performs, so it always survives
 * to the swap at the end. Two eligibility rules narrow WHERE a name may match, both added after a
 * 2026-08-26 review round found them missing:
 *  - never inside a URL span (splitOnUrls) — a name that is really part of a link's path/query is
 *    the link's own text, not a mention, and tokenizing it there corrupts the eventual href;
 *  - only at a word boundary (wordBoundaryPattern) — "Shiv" must not match inside "Shivankit",
 *    which would silently mangle the VISIBLE text, not just misplace the tag.
 *
 * Throws if a mention's name has zero ELIGIBLE occurrences (whole-word, outside any URL): a
 * mention resolved but never placed would otherwise be a silent drop — resolved, but nobody
 * actually gets notified. A sub-word or URL-only match does not count as an occurrence at all,
 * so it hits this same loud failure rather than a confusing silent skip.
 */
export function renderTextWithMentions(text: string, mentions: readonly MentionTarget[]): string {
  if (mentions.length === 0) {
    return textToHtml(text);
  }
  let segments = splitOnUrls(text);
  const tokens = mentions.map((mention, index) => {
    const hasEligibleOccurrence = segments.some(
      (segment) => !segment.isUrl && wordBoundaryPattern(mention.name, 'iu').test(segment.value),
    );
    if (!hasEligibleOccurrence) {
      throw new Error(
        `Mention "${mention.name}" (resolved to ${mention.displayName}) does not occur anywhere ` +
          'in the text as a whole word outside any URL — nothing to attach the mention to, so it ' +
          'would notify no one silently.',
      );
    }
    const token = `${PLACEHOLDER_GUARD}MENTION${index}${PLACEHOLDER_GUARD}`;
    const pattern = wordBoundaryPattern(mention.name, 'giu');
    segments = segments.map((segment) =>
      segment.isUrl ? segment : { isUrl: false, value: segment.value.replace(pattern, token) },
    );
    return token;
  });
  let html = textToHtml(segments.map((segment) => segment.value).join(''));
  mentions.forEach((mention, index) => {
    html = html.split(tokens[index] as string).join(`<at id="${index}">${escapeHtml(mention.displayName)}</at>`);
  });
  return html;
}

const PLACEHOLDER_TOKEN = /@\{([^{}]+)\}/g;

/**
 * The format:'html' contract: the caller writes their own raw HTML and marks each mention spot
 * with a literal `@{Name}` token (Name matched case-insensitively against a resolved mention's
 * `name`), and this replaces every such token with the matching `<at id="N">displayName</at>`
 * tag. Everything else in the HTML is left untouched — no escaping, same verbatim contract as
 * sendHtmlMessage itself.
 *
 * Throws on an unresolved placeholder (a `@{Name}` naming no known mention — a typo would
 * otherwise post literal `@{Name}` text with no notification) and on a resolved mention with no
 * placeholder anywhere in the html (resolved but never placed — the same silent-drop hazard
 * renderTextWithMentions guards against).
 */
export function renderHtmlWithMentions(html: string, mentions: readonly MentionTarget[]): string {
  if (mentions.length === 0) {
    return html;
  }
  const used = new Set<number>();
  const rendered = html.replace(PLACEHOLDER_TOKEN, (_match, rawName: string) => {
    const index = mentions.findIndex((mention) => mention.name.toLowerCase() === rawName.trim().toLowerCase());
    if (index === -1) {
      throw new Error(
        `Placeholder @{${rawName}} does not match any resolved mention ` +
          `(${mentions.map((mention) => mention.name).join(', ')}).`,
      );
    }
    used.add(index);
    const mention = mentions[index] as MentionTarget;
    return `<at id="${index}">${escapeHtml(mention.displayName)}</at>`;
  });
  const unplaced = mentions.filter((_, index) => !used.has(index));
  if (unplaced.length > 0) {
    throw new Error(
      `Mention(s) ${unplaced.map((mention) => mention.name).join(', ')} were resolved but no ` +
        `@{Name}-style placeholder for ${unplaced.length === 1 ? 'it' : 'them'} was found in the ` +
        'html — nothing to attach the mention to, so it would notify no one silently.',
    );
  }
  return rendered;
}
