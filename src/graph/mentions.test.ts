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
  { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
  { id: 'aad-johan', displayName: 'Spännare, Johan' },
  { id: 'aad-burhan', displayName: 'Öcüt, Burhan' },
];

describe('resolveMentionTargets — case-insensitive, unambiguous-substring resolution', () => {
  it('resolves a short first name against "Lastname, Firstname" (the doctrine example)', () => {
    const [resolved] = resolveMentionTargets(['Shiv'], MEMBERS);

    expect(resolved).toEqual({ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' });
  });

  it('is case-insensitive', () => {
    const [resolved] = resolveMentionTargets(['sHIv'], MEMBERS);

    expect(resolved?.displayName).toBe('Garg, Shivankit');
  });

  it('resolves several names in the order given', () => {
    const resolved = resolveMentionTargets(['Johan', 'Shiv'], MEMBERS);

    expect(resolved.map((m) => m.displayName)).toEqual(['Spännare, Johan', 'Garg, Shivankit']);
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
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { name: 'Johan', id: 'aad-johan', displayName: 'Spännare, Johan' },
    ];

    expect(buildGraphMentionsPayload(targets)).toEqual([
      { id: 0, mentionText: 'Garg, Shivankit', mentioned: { user: { id: 'aad-shiv', displayName: 'Garg, Shivankit' } } },
      { id: 1, mentionText: 'Spännare, Johan', mentioned: { user: { id: 'aad-johan', displayName: 'Spännare, Johan' } } },
    ]);
  });
});

describe('renderTextWithMentions — plain-text path (send_chat_message default format)', () => {
  it('with no mentions, behaves exactly like textToHtml', () => {
    expect(renderTextWithMentions('hello there', [])).toBe('<p>hello there</p>');
  });

  it('replaces the occurrence of the mention name with an <at> tag carrying the resolved displayName', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(renderTextWithMentions('Shiv can you review this?', [target])).toBe(
      '<p><at id="0">Garg, Shivankit</at> can you review this?</p>',
    );
  });

  it('matches the mention name case-insensitively', () => {
    const target: MentionTarget = { name: 'shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(renderTextWithMentions('SHIV please look', [target])).toContain('<at id="0">Garg, Shivankit</at>');
  });

  it('replaces every occurrence of the same mention with the same id', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    const rendered = renderTextWithMentions('Shiv, Shiv are you there?', [target]);

    expect(rendered.match(/<at id="0">/g)).toHaveLength(2);
  });

  it('handles two different mentions in one message', () => {
    const shiv: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };
    const johan: MentionTarget = { name: 'Johan', id: 'aad-johan', displayName: 'Spännare, Johan' };

    const rendered = renderTextWithMentions('Shiv and Johan, please sync', [shiv, johan]);

    expect(rendered).toBe(
      '<p><at id="0">Garg, Shivankit</at> and <at id="1">Spännare, Johan</at>, please sync</p>',
    );
  });

  it('HTML-escapes the resolved displayName inside the <at> tag', () => {
    const target: MentionTarget = { name: 'A&B', id: 'x', displayName: 'A & B <team>' };

    expect(renderTextWithMentions('A&B please check', [target])).toContain(
      '<at id="0">A &amp; B &lt;team&gt;</at>',
    );
  });

  it('refuses a mention whose name never occurs in the text — never a silent drop', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(() => renderTextWithMentions('no names here', [target])).toThrow(/does not occur anywhere/);
  });

  // Review round 2, MAJOR 1 (2026-08-26): a mention name occurring inside a URL used to corrupt
  // the outbound HTML — the placeholder was substituted into raw text BEFORE textToHtml
  // autolinked the URL, so the name-inside-URL got tokenized too, and the final <at>-tag swap
  // landed inside the eventual href attribute value, whose own quote character terminates the
  // attribute early (malformed HTML). URL spans must be excluded from substitution entirely.
  it('does not corrupt a URL that happens to contain the mention name as a path segment', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    const rendered = renderTextWithMentions('see https://x.com/shiv/pr Shiv', [target]);

    // The URL renders as one intact, unmangled anchor — no <at> markup anywhere inside it.
    expect(rendered).toContain('<a href="https://x.com/shiv/pr">https://x.com/shiv/pr</a>');
    expect(rendered).not.toMatch(/<a href="[^"]*<at/);
    // The real, standalone mention target outside the URL still gets tagged.
    expect(rendered).toContain('<at id="0">Garg, Shivankit</at>');
    expect(rendered).toBe(
      '<p>see <a href="https://x.com/shiv/pr">https://x.com/shiv/pr</a> <at id="0">Garg, Shivankit</at></p>',
    );
  });

  it('a mention name occurring ONLY inside a URL is not an eligible occurrence — throws, does not silently skip', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(() => renderTextWithMentions('see https://x.com/shiv/pr for details', [target])).toThrow(
      /does not occur anywhere/,
    );
  });

  // Review round 2, MAJOR 2 (2026-08-26): no word-boundary requirement meant "Shiv" matched
  // inside "Shivankit's", silently corrupting the VISIBLE text (not just misplacing the tag) —
  // the rendered message would read "...<at>Garg, Shivankit</at>ankit's branch...".
  it('requires a word boundary — does not match "Shiv" inside "Shivankit\'s"', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    const rendered = renderTextWithMentions("Shivankit's branch is ready — Shiv please review", [target]);

    expect(rendered).toBe(
      '<p>Shivankit\'s branch is ready — <at id="0">Garg, Shivankit</at> please review</p>',
    );
  });

  it('throws the loud "does not occur" error when only a sub-word match exists — never a silent skip', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(() => renderTextWithMentions("Shivankit's branch is ready", [target])).toThrow(
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
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(renderHtmlWithMentions('<p>@{Shiv} please review</p>', [target])).toBe(
      '<p><at id="0">Garg, Shivankit</at> please review</p>',
    );
  });

  it('matches the placeholder name case-insensitively', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(renderHtmlWithMentions('<p>@{shiv}</p>', [target])).toContain('<at id="0">');
  });

  it('leaves everything else in the html untouched — no escaping (same contract as sendHtmlMessage)', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(renderHtmlWithMentions('<table><tr><td>@{Shiv}</td></tr></table>', [target])).toBe(
      '<table><tr><td><at id="0">Garg, Shivankit</at></td></tr></table>',
    );
  });

  it('refuses an unresolved placeholder — a typo would otherwise post literal @{Name} text with no notification', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(() => renderHtmlWithMentions('<p>@{Shvi}</p>', [target])).toThrow(/does not match any resolved mention/);
  });

  it('refuses a resolved mention with no placeholder anywhere in the html — never a silent drop', () => {
    const target: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    expect(() => renderHtmlWithMentions('<p>no placeholder here</p>', [target])).toThrow(/no @\{Name\}-style placeholder/);
  });
});
