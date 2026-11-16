/**
 * Line breaking for wrapped cell text; mirrors `xlsx_render::wrap` rule for rule.
 * `\n` (and `\r\n`) always break, then words wrap greedily at spaces to `width`,
 * and a word wider than `width` breaks between grapheme clusters. Spaces where
 * a line breaks belong to no line; the text's last line keeps its trailing spaces,
 * and a final line break opens no empty line (Excel centres "a\nb\n" as two lines).
 */

import type { TextWrap } from '../display-list/types';

type Segmenter = { segment(value: string): Iterable<{ segment: string; index: number }> };

let segmenter: Segmenter | null | undefined;

function graphemeEnds(text: string): number[] {
  if (segmenter === undefined) {
    const Ctor = (Intl as unknown as {
      Segmenter?: new (locales?: string | string[], options?: { granularity: 'grapheme' }) => Segmenter;
    }).Segmenter;
    try {
      segmenter = typeof Ctor === 'function' ? new Ctor(undefined, { granularity: 'grapheme' }) : null;
    } catch {
      segmenter = null;
    }
  }
  if (segmenter) return [...segmenter.segment(text)].map((part) => part.index + part.segment.length);
  const ends: number[] = [];
  let at = 0;
  for (const char of text) ends.push((at += char.length));
  return ends;
}

export function wrapLines(text: string, width: number, measure: (run: string) => number): string[] {
  const limit = Number.isNaN(width) ? Number.POSITIVE_INFINITY : width;
  const lines: string[] = [];
  const body = text.endsWith('\n') ? text.slice(0, text.endsWith('\r\n') ? -2 : -1) : text;
  const paragraphs = body.split('\n');
  paragraphs.forEach((raw, index) => {
    wrapParagraph(raw.endsWith('\r') ? raw.slice(0, -1) : raw, limit, measure, lines);
    if (index < paragraphs.length - 1) lines[lines.length - 1] = lines[lines.length - 1].replace(/ +$/, '');
  });
  return lines;
}

function wrapParagraph(text: string, width: number, measure: (run: string) => number, lines: string[]): void {
  if (measure(text) <= width) {
    lines.push(text);
    return;
  }
  let start = 0;
  let end: number | null = null;
  for (const [wordStart, wordEnd] of words(text)) {
    if (end !== null) {
      if (measure(text.slice(start, wordEnd)) <= width) {
        end = wordEnd;
        continue;
      }
      lines.push(text.slice(start, end));
      start = wordStart;
    }
    end = wordEnd;
    if (measure(text.slice(start, wordEnd)) <= width) continue;
    const base = start;
    const ends = graphemeEnds(text.slice(base, wordEnd)).map((offset) => base + offset);
    let first = 0;
    for (let fits = lastFitting(text, start, ends, first, width, measure); fits !== null; ) {
      lines.push(text.slice(start, ends[fits]));
      start = ends[fits];
      first = fits + 1;
      fits = lastFitting(text, start, ends, first, width, measure);
    }
  }
  lines.push(text.slice(start));
}

/**
 * The last of the graphemes ending at `ends[first..]` that still fits on a line
 * from `start`, or `null` when they all do; a line keeps at least one. A run only
 * widens as it grows, so doubling then bisecting finds the break in
 * logarithmically many measures.
 */
function lastFitting(
  text: string,
  start: number,
  ends: number[],
  first: number,
  width: number,
  measure: (run: string) => number,
): number | null {
  let fits = first;
  let over: number;
  for (let step = 1; ; step *= 2) {
    const probe = Math.min(fits + step, ends.length - 1);
    if (probe === fits) return null;
    if (measure(text.slice(start, ends[probe])) > width) {
      over = probe;
      break;
    }
    fits = probe;
  }
  while (over - fits > 1) {
    const mid = fits + Math.floor((over - fits) / 2);
    if (measure(text.slice(start, ends[mid])) > width) over = mid;
    else fits = mid;
  }
  return fits;
}

function words(text: string): [number, number][] {
  const spans: [number, number][] = [];
  let at = 0;
  for (const word of text.split(' ')) {
    if (word.length > 0) spans.push([at, at + word.length]);
    at += word.length + 1;
  }
  return spans;
}

/** baseline of line `index` of `count`, from the single-line baseline `y`. */
export function wrappedLineY(y: number, index: number, count: number, wrap: TextWrap): number {
  const last = Math.max(count - 1, 0);
  const step = wrap.valign === 'top' ? index : wrap.valign === 'center' ? index - last / 2 : index - last;
  return y + wrap.lineHeight * step;
}
