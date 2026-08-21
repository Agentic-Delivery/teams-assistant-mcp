import { describe, expect, it } from 'vitest';
import { textToHtml } from './formatting.js';
import { htmlToText } from './messages.js';

describe('outbound text rendering', () => {
  it('wraps a single-line message in a paragraph', () => {
    expect(textToHtml('Hello everyone')).toBe('<p>Hello everyone</p>');
  });

  it('turns single newlines into <br> and blank lines into paragraph breaks', () => {
    expect(textToHtml('line one\nline two\n\nnew paragraph')).toBe(
      '<p>line one<br>line two</p><p>&nbsp;</p><p>new paragraph</p>',
    );
  });

  it('escapes markup so a hostile message cannot inject HTML', () => {
    const html = textToHtml('<script>alert(1)</script> and <img src=x onerror=alert(2)>');

    expect(html).not.toContain('<script');
    expect(html).not.toContain('<img');
    expect(html).toBe(
      '<p>&lt;script&gt;alert(1)&lt;/script&gt; and &lt;img src=x onerror=alert(2)&gt;</p>',
    );
  });

  it('turns an http(s) URL into a clickable anchor', () => {
    expect(textToHtml('see https://example.com/plan?a=1&b=2 for details')).toBe(
      '<p>see <a href="https://example.com/plan?a=1&amp;b=2">https://example.com/plan?a=1&amp;b=2</a> for details</p>',
    );
  });

  it('leaves sentence punctuation after a URL outside the link', () => {
    expect(textToHtml('details here: https://example.com/plan.')).toBe(
      '<p>details here: <a href="https://example.com/plan">https://example.com/plan</a>.</p>',
    );
  });

  it('keeps a closing paren inside the link only when the URL opened it', () => {
    expect(textToHtml('(see https://example.com/x)')).toBe(
      '<p>(see <a href="https://example.com/x">https://example.com/x</a>)</p>',
    );
    expect(textToHtml('https://en.wikipedia.org/wiki/API_(disambiguation)')).toBe(
      '<p><a href="https://en.wikipedia.org/wiki/API_(disambiguation)">' +
        'https://en.wikipedia.org/wiki/API_(disambiguation)</a></p>',
    );
  });

  it('does not linkify non-http schemes', () => {
    expect(textToHtml('run ftp://old.server/file')).toBe('<p>run ftp://old.server/file</p>');
  });

  it('round-trips through htmlToText, blank lines included', () => {
    const original =
      'Hello all!\n' +
      'The pilot round starts tomorrow.\n' +
      '\n' +
      'Details and signup: https://example.sharepoint.com/sites/pilot\n' +
      '\n' +
      'Regards, the assistant & the team';

    expect(htmlToText(textToHtml(original))).toBe(original);
  });
});
