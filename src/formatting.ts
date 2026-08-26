/**
 * Outbound counterpart of htmlToText in messages.ts. Teams renders a chatMessage body sent as
 * contentType 'text' as one unbroken blob — no line breaks, no clickable links — so everything
 * this server posts goes out as HTML built here.
 *
 * Deliberately tiny: paragraphs, line breaks and autolinked http(s) URLs. No markdown — bold,
 * lists and the rest would promise formatting the model never asked for and Teams' HTML subset
 * only half supports.
 *
 * The output round-trips through htmlToText: reading back a message this module produced yields
 * the original text again, provided the input is normalised the way htmlToText normalises
 * (single \n line endings, no runs of three or more newlines, no trailing spaces on a line).
 */

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const URL = /https?:\/\/[^\s<]+/g;

/**
 * A URL pasted mid-sentence usually drags its punctuation along ("see https://example.com/plan.").
 * Trailing sentence punctuation is never part of the link; a closing paren only is when the URL
 * itself opened one (Wikipedia-style "..._(disambiguation)").
 */
function splitTrailingPunctuation(candidate: string): [url: string, rest: string] {
  let url = candidate;
  for (;;) {
    const last = url.at(-1);
    if (last !== undefined && '.,;:!?\'"'.includes(last)) {
      url = url.slice(0, -1);
      continue;
    }
    if (
      last === ')' &&
      (url.match(/\(/g)?.length ?? 0) < (url.match(/\)/g)?.length ?? 0)
    ) {
      url = url.slice(0, -1);
      continue;
    }
    return [url, candidate.slice(url.length)];
  }
}

export interface UrlSpan {
  /** Index of the URL's first character in the input string. */
  start: number;
  /** Index one past the URL's last character — trailing sentence punctuation is NOT included. */
  end: number;
  url: string;
}

/**
 * Every URL span in `text`, trimmed of trailing sentence punctuation exactly the way renderLine
 * trims it before wrapping in an anchor. Exported so a caller that needs to know WHERE a URL
 * will land in the rendered HTML (mentions.ts excludes URL spans from name substitution, so a
 * mention name that is really just part of a link's path never gets tokenized) uses the same
 * boundary renderLine itself uses, rather than a second, driftable definition of "URL". The regex
 * itself only ever matches within one whitespace-delimited run, so scanning the full (possibly
 * multi-line) text at once is equivalent to scanning line by line.
 */
export function findUrlSpans(text: string): UrlSpan[] {
  return [...text.matchAll(URL)].map((match) => {
    const [url] = splitTrailingPunctuation(match[0]);
    const start = match.index;
    return { start, end: start + url.length, url };
  });
}

/** One line of text: everything escaped, URLs additionally wrapped in an anchor. */
function renderLine(line: string): string {
  let html = '';
  let consumed = 0;
  for (const { start, end, url } of findUrlSpans(line)) {
    html += escapeHtml(line.slice(consumed, start)) + `<a href="${escapeHtml(url)}">${escapeHtml(url)}</a>`;
    consumed = end;
  }
  return html + escapeHtml(line.slice(consumed));
}

/**
 * Renders plain text as the HTML body of an outbound Teams chatMessage. Blank-line-separated
 * paragraphs become <p> elements, single newlines become <br>, http(s) URLs become anchors.
 *
 * Paragraphs are separated by an explicit <p>&nbsp;</p> spacer because that is how Teams itself
 * encodes a blank line — the client styles <p> without margins, so adjacent <p> elements alone
 * render as consecutive lines, and the spacer is also what makes htmlToText give the blank
 * line back as \n\n.
 */
export function textToHtml(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .split(/\n{2,}/)
    .map((paragraph) => `<p>${paragraph.split('\n').map(renderLine).join('<br>')}</p>`)
    .join('<p>&nbsp;</p>');
}
