import type { PositionedTextRun } from '../types';

/** A piece to draw: `x` is where the run puts it; `width` and the split come from its single-precision values. */
interface TextChunk { text: string; x: number; start: number; width?: number; fitText?: string }

export function paintTextChunks(
  ctx: CanvasRenderingContext2D,
  run: PositionedTextRun,
  runBaseline: number,
  spacing: number
): void {
  const { chunks, fitted } = positionedTextChunks(run);
  for (const chunk of chunks) {
    const scale = chunkScale(chunk, spacing, (text) => ctx.measureText(text)?.width, fitted ? PIECE_FIT_MAX : WORD_FIT_MAX);
    if (scale === 1) {
      ctx.fillText(chunk.text, chunk.x, runBaseline);
    } else {
      ctx.save();
      ctx.translate(chunk.x, runBaseline);
      ctx.scale(scale, 1);
      ctx.fillText(chunk.text, 0, 0);
      ctx.restore();
    }
  }
}

/** The widest stretch of a word fitted to its caret stops. */
const WORD_FIT_MAX = 1.33;
/** The widest stretch of a piece laid out on another family's stored widths; the native engine uses the same. */
const PIECE_FIT_MAX = 1.45;

function chunkScale(
  chunk: TextChunk, spacing: number, measure: (text: string) => number | undefined, limit: number,
): number {
  const cjk = Array.from(chunk.text).some((character) => isCjkCharacter(character.codePointAt(0)));
  const measured = cjk || chunk.width === undefined || spacing !== 0 ? 0 : measure(chunk.fitText ?? chunk.text);
  let scale = measured && measured > 0 ? (chunk.width ?? 0) / measured : 1;
  if (Math.abs(scale - 1) < 0.001 || scale < 0.75 || scale > limit) scale = 1;
  return scale;
}

/**
 * The pieces a run laid out on another family's stored widths is fitted in, each with its start offset and
 * scale, or null for a run that is not fitted. The native engine makes the same decisions.
 * @internal
 */
export function fittedPieceDecisions(
  run: PositionedTextRun, measure: (text: string) => number,
): Array<{ start: number; scale: number }> | null {
  const { chunks, fitted } = positionedTextChunks(run);
  if (!fitted) return null;
  return chunks.map((chunk) => ({ start: chunk.start, scale: chunkScale(chunk, run.letterSpacingPx ?? 0, measure, PIECE_FIT_MAX) }));
}

/** The run in the single precision the engine laid it out in, so every decision reads the values it has. */
function singlePrecision(run: PositionedTextRun): PositionedTextRun {
  const f = Math.fround;
  return {
    ...run, x: f(run.x), width: f(run.width),
    glyphs: run.glyphs.map((glyph) => ({ ...glyph, x: f(glyph.x), advance: f(glyph.advance), xOffset: f(glyph.xOffset) })),
  };
}

function positionedTextChunks(laidOut: PositionedTextRun): { chunks: TextChunk[]; fitted: boolean } {
  const run = singlePrecision(laidOut);
  if (run.glyphs.length === 0) return { chunks: [{ text: run.text, x: laidOut.x, start: 0 }], fitted: false };
  // A right-to-left run's clusters step leftward: the canvas orders each tab-free
  // piece itself, drawn from that piece's leftmost glyph. A cluster can open
  // with a mark or dots glyph at its right edge, where the cluster before it
  // ends, so it is the clusters' leftmost glyphs that tell the direction, from
  // the first one that a zero-width cluster does not leave in place.
  const lefts = new Map<number, number>();
  for (const glyph of run.glyphs) {
    lefts.set(glyph.cluster, Math.min(lefts.get(glyph.cluster) ?? Infinity, glyph.x));
  }
  const [firstLeft, ...laterLefts] = lefts.values();
  const nextLeft = laterLefts.find((left) => Math.abs(left - firstLeft) > 0.01);
  if (nextLeft !== undefined && nextLeft < firstLeft) {
    const texts = run.text.split('\t');
    const ends: number[] = [];
    let end = -1;
    for (const text of texts) {
      end += text.length + 1;
      ends.push(end);
    }
    const lefts = texts.map(() => Infinity);
    for (const [index, glyph] of run.glyphs.entries()) {
      const offset = glyph.cluster - run.start;
      let low = 0;
      let high = ends.length - 1;
      while (low < high) {
        const middle = (low + high) >> 1;
        if (ends[middle] < offset) low = middle + 1;
        else high = middle;
      }
      if (offset >= 0 && offset < ends[low]) lefts[low] = Math.min(lefts[low], laidOut.glyphs[index].x);
    }
    const starts = texts.map((_, index) => (index ? ends[index - 1] + 1 : 0));
    return {
      chunks: texts.flatMap((text, index) =>
        text.length > 0 && lefts[index] < Infinity ? [{ text, x: lefts[index], start: starts[index] }] : [],
      ),
      fitted: false,
    };
  }
  const chunks: TextChunk[] = [];
  const continuation = (offset: number): boolean =>
    /^[\p{M}\u200c\u200d]/u.test(run.text.slice(offset)) || run.text[offset - 1] === '\u200d';
  const boundary = (offset: number, departs: boolean): boolean =>
    /\s/u.test(run.text[offset - 1]) ||
    isCjkCharacter(run.text.codePointAt(offset)) ||
    isCjkCharacter(codePointBefore(run.text, offset)) ||
    (departs && /\s/u.test(run.text[offset]));
  // The pen starts at the run's own x, so a leading tab, which paints no glyph,
  // splits the run where the text after it lands.
  const expected = (index: number): number => {
    const previous = run.glyphs[index - 1];
    return previous ? Math.fround(previous.x + previous.advance) : run.x;
  };
  const departs = (index: number): boolean => Math.abs(run.glyphs[index].x - expected(index)) > 0.0001;
  // A fitted piece starts only at a cluster's first glyph that departs by more than 0.0001 px or two
  // single-precision steps at its x, so rounding inside a cluster or at large x does not split one.
  const pieceDeparts = (index: number): boolean => {
    if (index > 0 && run.glyphs[index - 1].cluster === run.glyphs[index].cluster) return false;
    const expectedX = expected(index);
    const tolerance = Math.max(Math.fround(0.0001), Math.fround(Math.abs(expectedX) * 2 ** -21));
    return Math.abs(Math.fround(run.glyphs[index].x - expectedX)) > tolerance;
  };

  // A face standing in for a family we only have widths for steps each glyph
  // to that family's x, so such a run is drawn piece by piece where it departs,
  // each piece fitted to the span up to the next one.
  const namedWidths = !run.letterSpacingPx && run.glyphs.some((glyph, index) => {
    const offset = glyph.cluster - run.start;
    return offset > 0 && offset < run.text.length && departs(index) &&
      !continuation(offset) && !boundary(offset, true);
  });
  // A departure after a mark, which layout sets back over its base, is not a
  // width table's, so only pieces of runs that depart after a spacing glyph are fitted.
  // Pieces are text ranges, so only a run whose clusters follow the text in order is fitted.
  const inOrder = run.glyphs.every((glyph, index) => index === 0 || glyph.cluster >= run.glyphs[index - 1].cluster);
  const fitted = namedWidths && inOrder && run.glyphs.some((glyph, index) => {
    const offset = glyph.cluster - run.start;
    return index > 0 && run.glyphs[index - 1].advance > 0 && offset > 0 && offset < run.text.length &&
      pieceDeparts(index) && !continuation(offset) && !boundary(offset, true);
  });
  // Right-to-left text laid out in a left-to-right line keeps the browser's own
  // bidi ordering, so only the pen discontinuities split it.
  const anchored = !namedWidths && !RTL_TEXT.test(run.text) && inOrder;
  let textStart = 0;
  const startsAtGlyph = anchored && run.glyphs[0].cluster === run.start;
  let x = startsAtGlyph ? run.glyphs[0].x : run.x;
  let paintX = startsAtGlyph ? laidOut.glyphs[0].x : laidOut.x;
  let glyphStart = 0;
  const pushChunk = (textEnd: number, glyphEnd: number, endX: number): void => {
    const text = run.text.slice(textStart, textEnd);
    const chunk: TextChunk = { text, x: paintX, start: textStart };
    if (anchored && !text.includes('\t')) {
      const fitText = text.trimEnd();
      let span = endX - x;
      let advance = 0;
      let lastX = x;
      for (let index = glyphStart; index < glyphEnd; index++) {
        const glyph = run.glyphs[index];
        if (glyph.cluster - run.start >= textStart + fitText.length) {
          // Fit visible text before the whitespace that absorbs justification.
          span = glyph.x - x;
          chunk.fitText = fitText;
          break;
        }
        advance += glyph.advance;
        lastX = Math.max(lastX, glyph.x);
      }
      chunk.width = Number.isFinite(span) && span > 0 && span >= lastX - x ? span : advance;
    } else if (fitted && !text.includes('\t') && endX > x) {
      chunk.width = endX - x;
    }
    chunks.push(chunk);
  };
  for (const [index, glyph] of run.glyphs.entries()) {
    const offset = glyph.cluster - run.start;
    if (
      offset > textStart &&
      offset < run.text.length &&
      (anchored ? !continuation(offset) && boundary(offset, departs(index)) : fitted ? pieceDeparts(index) : departs(index))
    ) {
      pushChunk(offset, index, glyph.x);
      textStart = offset;
      glyphStart = index;
      x = glyph.x;
      paintX = laidOut.glyphs[index].x;
    }
  }
  pushChunk(run.text.length, run.glyphs.length, run.x + run.width);
  return { chunks, fitted };
}

const RTL_TEXT = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufefc]/u;

function codePointBefore(text: string, offset: number): number | undefined {
  return text.codePointAt(offset - (/[\ud800-\udbff][\udc00-\udfff]$/.test(text.slice(0, offset)) ? 2 : 1));
}

function isCjkCharacter(codePoint: number | undefined): boolean {
  return /[\u3000-\u3098\u309b-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef\u{20000}-\u{2ffff}]/u.test(
    String.fromCodePoint(codePoint ?? 0),
  );
}
