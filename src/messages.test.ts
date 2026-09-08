import { describe, expect, it } from 'vitest';
import { applyWatermark, htmlToText, toChatMessage } from './messages.js';

describe('message text', () => {
  it('turns the HTML Teams actually sends into readable text', () => {
    const html = '<div><p>Hello &amp; welcome</p><p>Second line</p></div>';

    expect(htmlToText(html)).toBe('Hello & welcome\nSecond line');
  });

  it('keeps the visible text of a mention and drops the markup', () => {
    const html = '<div><at id="0">Alice</at> shall we meet at the caf&#233;?</div>';

    expect(htmlToText(html)).toBe('Alice shall we meet at the café?');
  });

  it('collapses the run of empty blocks Teams leaves behind', () => {
    expect(htmlToText('<p>one</p><p></p><p></p><p>two</p>')).toBe('one\n\ntwo');
  });

  it('treats the raw newlines Teams pretty-prints into stored bodies as whitespace', () => {
    // Observed live 2026-08-20: Teams rewrites a posted body with \n after <br> and between
    // the <p> elements. Those are HTML source whitespace, not line breaks.
    const html = '<p>line one<br>\nline two</p>\n<p>&nbsp;</p>\n<p>paragraph two</p>';

    expect(htmlToText(html)).toBe('line one\nline two\n\nparagraph two');
  });

  it('keeps a pasted link\'s URL when the visible label differs from it', () => {
    // Observed live 2026-09-08: a chat message with a pasted link arrived with attachments: 0
    // and text that kept only the anchor's label, silently dropping the href. This reproduces
    // the shape (a page-title label pointing at a slugified URL), not the live content.
    const html =
      '<div>Please review this page and its sub pages. Summarise the Product Overview ' +
      '<a href="https://contoso.sharepoint.com/wiki/Product-Overview" ' +
      'title="https://contoso.sharepoint.com/wiki/Product-Overview">Product Overview page</a></div>';

    expect(htmlToText(html)).toBe(
      'Please review this page and its sub pages. Summarise the Product Overview ' +
        'Product Overview page (https://contoso.sharepoint.com/wiki/Product-Overview)',
    );
  });

  it('keeps a bare pasted URL as-is instead of doubling it up', () => {
    // Guards against a doubling regression: main (pre-fix) also passes this one, since it never
    // wrapped a bare URL at all — this proves the fix's own dedupe branch, not just "no crash".
    const html = '<p>see <a href="https://example.com/plan">https://example.com/plan</a></p>';

    expect(htmlToText(html)).toBe('see https://example.com/plan');
  });

  it('decodes an ampersand in the href the same way the rest of the text is decoded', () => {
    const html = '<p><a href="https://example.com/x?a=1&amp;b=2">the plan</a></p>';

    expect(htmlToText(html)).toBe('the plan (https://example.com/x?a=1&b=2)');
  });

  it('strips nested markup from a link label without losing the href', () => {
    const html = '<p><a href="https://example.com/plan">the <b>plan</b></a></p>';

    expect(htmlToText(html)).toBe('the plan (https://example.com/plan)');
  });

  it('turns a block tag inside a link label into a space instead of gluing words together', () => {
    const html = '<p><a href="https://example.com/plan">line1<br>line2</a></p>';

    expect(htmlToText(html)).toBe('line1 line2 (https://example.com/plan)');
  });

  it('dedupes a bare pasted URL whose non-ASCII character is percent-encoded in the href but shown decoded in the label', () => {
    // A synthetic non-ASCII word (not customer content): "smorgasbord" written as "smörgåsbord".
    const html = '<p><a href="https://example.com/sm%C3%B6rg%C3%A5sbord">https://example.com/sm&#246;rg&#229;sbord</a></p>';

    expect(htmlToText(html)).toBe('https://example.com/sm%C3%B6rg%C3%A5sbord');
  });
});

describe('graph message mapping', () => {
  it('maps a normal message to sender, text and timestamp', () => {
    const message = toChatMessage(
      {
        id: '1700000000000',
        chatId: '19:pilot@thread.v2',
        createdDateTime: '2026-08-19T08:00:00Z',
        from: { user: { id: 'oid-1', displayName: 'Alice Anderson' } },
        body: { contentType: 'html', content: '<p>Looks good</p>' },
      },
      '19:fallback@thread.v2',
    );

    expect(message).toMatchObject({
      id: '1700000000000',
      chatId: '19:pilot@thread.v2',
      from: 'Alice Anderson',
      fromId: 'oid-1',
      text: 'Looks good',
      isDeleted: false,
      attachments: [],
    });
  });

  it('labels a system event instead of leaving the sender blank', () => {
    const message = toChatMessage(
      {
        id: '2',
        createdDateTime: '2026-08-19T08:00:00Z',
        messageType: 'systemEventMessage',
        from: null,
        body: { contentType: 'html', content: '<systemEventMessage/>' },
      },
      '19:pilot@thread.v2',
    );

    expect(message.from).toBe('system');
    expect(message.chatId).toBe('19:pilot@thread.v2');
  });

  it('carries attachments through with the fields needed to fetch them', () => {
    const message = toChatMessage(
      {
        id: '3',
        createdDateTime: '2026-08-19T08:00:00Z',
        body: { contentType: 'text', content: 'see attachment' },
        attachments: [
          {
            id: 'att-1',
            name: 'report.xlsx',
            contentType: 'reference',
            contentUrl: 'https://contoso.sharepoint.com/x/report.xlsx',
          },
          { name: 'no id, dropped' },
        ],
      },
      '19:pilot@thread.v2',
    );

    expect(message.attachments).toEqual([
      {
        id: 'att-1',
        name: 'report.xlsx',
        contentType: 'reference',
        contentUrl: 'https://contoso.sharepoint.com/x/report.xlsx',
      },
    ]);
  });

  it('surfaces a pasted inline image as a downloadable pseudo-attachment', () => {
    const message = toChatMessage(
      {
        id: '5',
        createdDateTime: '2026-08-19T08:00:00Z',
        body: {
          contentType: 'html',
          content:
            '<p>see here</p><img src="https://graph.microsoft.com/v1.0/chats/19%3Apilot%40thread.v2/' +
            'messages/5/hostedContents/aWQ9eF8wLXNlLWQx/$value" itemid="0-se-d1" width="250">',
        },
      },
      '19:pilot@thread.v2',
    );

    expect(message.attachments).toEqual([
      { id: 'aWQ9eF8wLXNlLWQx', contentType: 'image/*', name: 'inline-image-1' },
    ]);
    expect(message.text).toBe('see here');
  });

  it('numbers several pasted images and keeps the real attachments first', () => {
    const message = toChatMessage(
      {
        id: '6',
        createdDateTime: '2026-08-19T08:00:00Z',
        body: {
          contentType: 'html',
          content:
            '<img src="…/messages/6/hostedContents/first/$value">' +
            '<img src="…/messages/6/hostedContents/second/$value">',
        },
        attachments: [{ id: 'att-1', name: 'plan.docx', contentType: 'reference', contentUrl: 'https://x/y' }],
      },
      '19:pilot@thread.v2',
    );

    expect(message.attachments.map((a) => a.id)).toEqual(['att-1', 'first', 'second']);
    expect(message.attachments[2]?.name).toBe('inline-image-2');
  });

  it('does not invent attachments from a plain-text body that mentions hostedContents', () => {
    const message = toChatMessage(
      {
        id: '7',
        createdDateTime: '2026-08-19T08:00:00Z',
        body: {
          contentType: 'text',
          content: 'see <img src="https://g/hostedContents/x/$value"> in the docs',
        },
      },
      '19:pilot@thread.v2',
    );

    expect(message.attachments).toEqual([]);
  });

  it('flags a deleted message rather than hiding it', () => {
    const message = toChatMessage(
      { id: '4', createdDateTime: '2026-08-19T08:00:00Z', deletedDateTime: '2026-08-19T09:00:00Z' },
      '19:pilot@thread.v2',
    );

    expect(message.isDeleted).toBe(true);
  });
});

function at(iso: string, id = iso) {
  return toChatMessage({ id, createdDateTime: iso, body: { contentType: 'text', content: id } }, 'c');
}

describe('watermark diffing', () => {
  const batch = [
    at('2026-08-19T10:00:00Z', 'third'),
    at('2026-08-19T08:00:00Z', 'first'),
    at('2026-08-19T09:00:00Z', 'second'),
  ];

  it('returns everything oldest-first on a first read, with the newest as watermark', () => {
    const result = applyWatermark(batch);

    expect(result.messages.map((m) => m.id)).toEqual(['first', 'second', 'third']);
    expect(result.watermark).toBe('2026-08-19T10:00:00Z');
  });

  it('treats the watermark as exclusive so the same message is never delivered twice', () => {
    const result = applyWatermark(batch, '2026-08-19T09:00:00Z');

    expect(result.messages.map((m) => m.id)).toEqual(['third']);
    expect(result.watermark).toBe('2026-08-19T10:00:00Z');
  });

  it('returns no watermark when nothing is new, so the caller keeps the one it has', () => {
    const result = applyWatermark(batch, '2026-08-19T10:00:00Z');

    expect(result.messages).toEqual([]);
    expect(result.watermark).toBeUndefined();
  });

  it('compares instants, not strings, across timezone offsets', () => {
    // 12:00+02:00 is the same instant as 10:00Z — a string compare would wrongly call it newer.
    const result = applyWatermark(batch, '2026-08-19T12:00:00+02:00');

    expect(result.messages).toEqual([]);
  });

  it('rejects a watermark that is not a timestamp instead of silently returning everything', () => {
    expect(() => applyWatermark(batch, 'yesterday')).toThrow(/ISO-8601/);
  });
});
