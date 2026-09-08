import { describe, expect, it } from 'vitest';
import {
  buildGraphMentionsPayload,
  renderHtmlWithMentions,
  renderTextWithMentions,
  resolveMentionTargets,
  type ChatMember,
  type MentionTarget,
} from './mentions.js';

const MEMBERS: ChatMember[] = [
  { id: 'aad-mika', displayName: 'Berggren, Mikael' },
  { id: 'aad-johan', displayName: 'Spännare, Johan' },
  { id: 'aad-petri', displayName: 'Aaltonen, Petri' },
];

describe('resolveMentionTargets — case-insensitive, unambiguous-substring resolution', () => {
  it('resolves a short first name against "Lastname, Firstname" (the doctrine example)', () => {
    const [resolved] = resolveMentionTargets(['Mika'], MEMBERS);

    expect(resolved).toEqual({ name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' });
  });

  it('is case-insensitive', () => {
    const [resolved] = resolveMentionTargets(['mIkA'], MEMBERS);

    expect(resolved?.displayName).toBe('Berggren, Mikael');
  });

  it('resolves several names in the order given', () => {
    const resolved = resolveMentionTargets(['Johan', 'Mika'], MEMBERS);

    expect(resolved.map((m) => m.displayName)).toEqual(['Spännare, Johan', 'Berggren, Mikael']);
  });

  it('refuses a name matching no member — never a silent drop', () => {
    expect(() => resolveMentionTargets(['Nobody'], MEMBERS)).toThrow(/No chat member matches/);
  });

  it('refuses an ambiguous name matching more than one member', () => {
    const bothOs = [{ id: 'a', displayName: 'O Brien' }, { id: 'b', displayName: 'O Connor' }];

    expect(() => resolveMentionTargets(['O'], bothOs)).toThrow(/ambiguous/);
  });

  it('refuses an empty name', () => {
    expect(() => resolveMentionTargets([' '], MEMBERS)).toThrow(/empty/);
  });

  it('refuses a matched member with no recorded AAD id — cannot notify them', () => {
    const noId: ChatMember[] = [{ displayName: 'Ghost Account' }];

    expect(() => resolveMentionTargets(['Ghost'], noId)).toThrow(/no AAD user id/);
  });
});

describe('buildGraphMentionsPayload', () => {
  it('numbers mentions 0..n-1, distinct from the AAD id', () => {
    const targets: MentionTarget[] = [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
      { name: 'Johan', id: 'aad-johan', displayName: 'Spännare, Johan' },
    ];

    expect(buildGraphMentionsPayload(targets)).toEqual([
      { id: 0, mentionText: 'Berggren, Mikael', mentioned: { user: { id: 'aad-mika', displayName: 'Berggren, Mikael' } } },
      { id: 1, mentionText: 'Spännare, Johan', mentioned: { user: { id: 'aad-johan', displayName: 'Spännare, Johan' } } },
    ]);
  });
});

describe('renderTextWithMentions — plain-text path (send_chat_message default format)', () => {
  it('with no mentions, behaves exactly like textToHtml', () => {
    expect(renderTextWithMentions('hello there', [])).toBe('<p>hello there</p>');
  });

  it('replaces the occurrence of the mention name with an <at> tag carrying the resolved displayName', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(renderTextWithMentions('Mika can you review this?', [target])).toBe(
      '<p><at id="0">Berggren, Mikael</at> can you review this?</p>',
    );
  });

  it('matches the mention name case-insensitively', () => {
    const target: MentionTarget = { name: 'mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(renderTextWithMentions('MIKA please look', [target])).toContain('<at id="0">Berggren, Mikael</at>');
  });

  it('replaces every occurrence of the same mention with the same id', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    const rendered = renderTextWithMentions('Mika, Mika are you there?', [target]);

    expect(rendered.match(/<at id="0">/g)).toHaveLength(2);
  });

  it('handles two different mentions in one message', () => {
    const mika: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };
    const johan: MentionTarget = { name: 'Johan', id: 'aad-johan', displayName: 'Spännare, Johan' };

    const rendered = renderTextWithMentions('Mika and Johan, please sync', [mika, johan]);

    expect(rendered).toBe(
      '<p><at id="0">Berggren, Mikael</at> and <at id="1">Spännare, Johan</at>, please sync</p>',
    );
  });

  it('HTML-escapes the resolved displayName inside the <at> tag', () => {
    const target: MentionTarget = { name: 'A&B', id: 'x', displayName: 'A & B <team>' };

    expect(renderTextWithMentions('A&B please check', [target])).toContain(
      '<at id="0">A &amp; B &lt;team&gt;</at>',
    );
  });

  it('refuses a mention whose name never occurs in the text — never a silent drop', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(() => renderTextWithMentions('no names here', [target])).toThrow(/does not occur anywhere/);
  });

  // Review round 2, MAJOR 1 (2026-08-26): a mention name occurring inside a URL used to corrupt
  // the outbound HTML — the placeholder was substituted into raw text BEFORE textToHtml
  // autolinked the URL, so the name-inside-URL got tokenized too, and the final <at>-tag swap
  // landed inside the eventual href attribute value, whose own quote character terminates the
  // attribute early (malformed HTML). URL spans must be excluded from substitution entirely.
  it('does not corrupt a URL that happens to contain the mention name as a path segment', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    const rendered = renderTextWithMentions('see https://x.com/mika/pr Mika', [target]);

    // The URL renders as one intact, unmangled anchor — no <at> markup anywhere inside it.
    expect(rendered).toContain('<a href="https://x.com/mika/pr">https://x.com/mika/pr</a>');
    expect(rendered).not.toMatch(/<a href="[^"]*<at/);
    // The real, standalone mention target outside the URL still gets tagged.
    expect(rendered).toContain('<at id="0">Berggren, Mikael</at>');
    expect(rendered).toBe(
      '<p>see <a href="https://x.com/mika/pr">https://x.com/mika/pr</a> <at id="0">Berggren, Mikael</at></p>',
    );
  });

  it('a mention name occurring ONLY inside a URL is not an eligible occurrence — throws, does not silently skip', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(() => renderTextWithMentions('see https://x.com/mika/pr for details', [target])).toThrow(
      /does not occur anywhere/,
    );
  });

  // Review round 2, MAJOR 2 (2026-08-26): no word-boundary requirement meant "Mika" matched
  // inside "Mikael's", silently corrupting the VISIBLE text (not just misplacing the tag) —
  // the rendered message would read "...<at>Berggren, Mikael</at>el's branch...".
  it('requires a word boundary — does not match "Mika" inside "Mikael\'s"', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    const rendered = renderTextWithMentions("Mikael's branch is ready — Mika please review", [target]);

    expect(rendered).toBe(
      '<p>Mikael\'s branch is ready — <at id="0">Berggren, Mikael</at> please review</p>',
    );
  });

  it('throws the loud "does not occur" error when only a sub-word match exists — never a silent skip', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(() => renderTextWithMentions("Mikael's branch is ready", [target])).toThrow(
      /does not occur anywhere/,
    );
  });

  it('a name with non-ASCII letters still respects word boundaries (Unicode-aware)', () => {
    const target: MentionTarget = { name: 'Johan', id: 'aad-johan', displayName: 'Spännare, Johan' };

    // "Johana" is a different word entirely; only the standalone "Johan" should match.
    expect(() => renderTextWithMentions('Johana is not Johan', [target])).not.toThrow();
    const rendered = renderTextWithMentions('Johana is not Johan', [target]);
    expect(rendered).toBe('<p>Johana is not <at id="0">Spännare, Johan</at></p>');
  });
});

describe('renderHtmlWithMentions — the format:"html" @{Name} placeholder contract', () => {
  it('with no mentions, returns the html untouched', () => {
    expect(renderHtmlWithMentions('<b>hi</b>', [])).toBe('<b>hi</b>');
  });

  it('replaces a @{Name} placeholder with the matching <at> tag', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(renderHtmlWithMentions('<p>@{Mika} please review</p>', [target])).toBe(
      '<p><at id="0">Berggren, Mikael</at> please review</p>',
    );
  });

  it('matches the placeholder name case-insensitively', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(renderHtmlWithMentions('<p>@{mika}</p>', [target])).toContain('<at id="0">');
  });

  it('leaves everything else in the html untouched — no escaping (same contract as sendHtmlMessage)', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(renderHtmlWithMentions('<table><tr><td>@{Mika}</td></tr></table>', [target])).toBe(
      '<table><tr><td><at id="0">Berggren, Mikael</at></td></tr></table>',
    );
  });

  it('refuses an unresolved placeholder — a typo would otherwise post literal @{Name} text with no notification', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(() => renderHtmlWithMentions('<p>@{Mkia}</p>', [target])).toThrow(/does not match any resolved mention/);
  });

  it('refuses a resolved mention with no placeholder anywhere in the html — never a silent drop', () => {
    const target: MentionTarget = { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' };

    expect(() => renderHtmlWithMentions('<p>no placeholder here</p>', [target])).toThrow(/no @\{Name\}-style placeholder/);
  });
});
