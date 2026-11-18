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
  return wrapSpans(text, width, (start, end) => measure(text.slice(start, end))).map(([start, end]) => text.slice(start, end));
}

/** The `[start, end)` spans of the lines `wrapLines` breaks `text` into, with `measure` given the span to measure. */
export function wrapSpans(text: string, width: number, measure: (start: number, end: number) => number): [number, number][] {
  const limit = Number.isNaN(width) ? Number.POSITIVE_INFINITY : width;
  const spans: [number, number][] = [];
  const body = text.endsWith('\n') ? text.slice(0, text.endsWith('\r\n') ? -2 : -1) : text;
  const paragraphs = body.split('\n');
  let at = 0;
  paragraphs.forEach((raw, index) => {
    const start = at;
    at += raw.length + 1;
    wrapParagraph(text, start, start + (raw.endsWith('\r') ? raw.length - 1 : raw.length), limit, measure, spans);
    const last = spans[spans.length - 1]!;
    if (index < paragraphs.length - 1) last[1] = last[0] + text.slice(last[0], last[1]).replace(/ +$/, '').length;
  });
  return spans;
}

function wrapParagraph(
  text: string,
  from: number,
  to: number,
  width: number,
  measure: (start: number, end: number) => number,
  spans: [number, number][],
): void {
  if (measure(from, to) <= width) {
    spans.push([from, to]);
    return;
  }
  let start = from;
  let end: number | null = null;
  for (const [wordStart, wordEnd] of words(text.slice(from, to), from)) {
    if (end !== null) {
      if (measure(start, wordEnd) <= width) {
        end = wordEnd;
        continue;
      }
      spans.push([start, end]);
      start = wordStart;
    }
    end = wordEnd;
    if (measure(start, wordEnd) <= width) continue;
    const base = start;
    const ends = graphemeEnds(text.slice(base, wordEnd)).map((offset) => base + offset);
    let first = 0;
    for (let fits = lastFitting(start, ends, first, width, measure); fits !== null; ) {
      spans.push([start, ends[fits]!]);
      start = ends[fits]!;
      first = fits + 1;
      fits = lastFitting(start, ends, first, width, measure);
    }
  }
  spans.push([start, to]);
}

/**
 * The last of the graphemes ending at `ends[first..]` that still fits on a line
 * from `start`, or `null` when they all do; a line keeps at least one. A run only
 * widens as it grows, so doubling then bisecting finds the break in
 * logarithmically many measures.
 */
function lastFitting(
  start: number,
  ends: number[],
  first: number,
  width: number,
  measure: (start: number, end: number) => number,
): number | null {
  let fits = first;
  let over: number;
  for (let step = 1; ; step *= 2) {
    const probe = Math.min(fits + step, ends.length - 1);
    if (probe === fits) return null;
    if (measure(start, ends[probe]!) > width) {
      over = probe;
      break;
    }
    fits = probe;
  }
  while (over - fits > 1) {
    const mid = fits + Math.floor((over - fits) / 2);
    if (measure(start, ends[mid]!) > width) over = mid;
    else fits = mid;
  }
  return fits;
}

function words(text: string, offset: number): [number, number][] {
  const spans: [number, number][] = [];
  let at = offset;
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
