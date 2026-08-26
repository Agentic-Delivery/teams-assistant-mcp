import { escapeHtml, textToHtml } from '../formatting.js';

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
 * Renders plain text to the outbound HTML body, same as textToHtml, except every case-insensitive
 * occurrence of each mention's `name` becomes a `<at id="N">displayName</at>` tag that Graph's
 * `mentions` array (buildGraphMentionsPayload) will actually notify.
 *
 * Substitution happens via a placeholder token BEFORE textToHtml runs — inserting the raw <at>
 * markup after escaping would be safe from re-escaping, but doing the name search on already
 * paragraph/br-wrapped HTML risks matching inside markup Teams itself added (an autolinked URL,
 * say). The placeholder is inert to every transform textToHtml performs, so it always survives
 * to the swap at the end.
 *
 * Throws if a mention's name has zero occurrences in the text: a mention resolved but never
 * placed would otherwise be a silent drop — resolved, but nobody actually gets notified.
 */
export function renderTextWithMentions(text: string, mentions: readonly MentionTarget[]): string {
  if (mentions.length === 0) {
    return textToHtml(text);
  }
  let working = text;
  const tokens = mentions.map((mention, index) => {
    const pattern = new RegExp(escapeRegExp(mention.name), 'i');
    if (!pattern.test(working)) {
      throw new Error(
        `Mention "${mention.name}" (resolved to ${mention.displayName}) does not occur anywhere ` +
          'in the text — nothing to attach the mention to, so it would notify no one silently.',
      );
    }
    const token = `${PLACEHOLDER_GUARD}MENTION${index}${PLACEHOLDER_GUARD}`;
    working = working.replace(new RegExp(escapeRegExp(mention.name), 'gi'), token);
    return token;
  });
  let html = textToHtml(working);
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
