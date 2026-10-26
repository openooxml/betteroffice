import type { PositionedTextRun } from '../types';

export function paintTextChunks(
  ctx: CanvasRenderingContext2D,
  run: PositionedTextRun,
  runBaseline: number,
  spacing: number
): void {
  for (const chunk of positionedTextChunks(run)) {
    const measured = chunk.width === undefined || spacing !== 0
      ? 0 : ctx.measureText(chunk.fitText ?? chunk.text)?.width;
    let scale = measured && measured > 0 ? (chunk.width ?? 0) / measured : 1;
    if (Math.abs(scale - 1) < 0.001 || scale < 0.75 || scale > 1.33) scale = 1;
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

function positionedTextChunks(
  run: PositionedTextRun
): Array<{ text: string; x: number; width?: number; fitText?: string }> {
  if (run.glyphs.length === 0) return [{ text: run.text, x: run.x }];
  // A right-to-left run's glyphs step leftward: the canvas orders each tab-free
  // piece itself, drawn from that piece's leftmost glyph.
  const nextCluster = run.glyphs.find((glyph) => glyph.cluster !== run.glyphs[0].cluster);
  if (nextCluster && nextCluster.x < run.glyphs[0].x) {
    const texts = run.text.split('\t');
    const ends: number[] = [];
    let end = -1;
    for (const text of texts) {
      end += text.length + 1;
      ends.push(end);
    }
    const lefts = texts.map(() => Infinity);
    for (const glyph of run.glyphs) {
      const offset = glyph.cluster - run.start;
      let low = 0;
      let high = ends.length - 1;
      while (low < high) {
        const middle = (low + high) >> 1;
        if (ends[middle] < offset) low = middle + 1;
        else high = middle;
      }
      if (offset >= 0 && offset < ends[low]) lefts[low] = Math.min(lefts[low], glyph.x);
    }
    return texts.flatMap((text, index) =>
      text.length > 0 && lefts[index] < Infinity ? [{ text, x: lefts[index] }] : [],
    );
  }
  const chunks: Array<{ text: string; x: number; width?: number; fitText?: string }> = [];
  const continuation = (offset: number): boolean =>
    /^[\p{M}\u200c\u200d]/u.test(run.text.slice(offset)) || run.text[offset - 1] === '\u200d';
  const boundary = (offset: number, departs: boolean): boolean =>
    /\s/u.test(run.text[offset - 1]) ||
    isCjkCharacter(run.text.codePointAt(offset)) ||
    isCjkCharacter(codePointBefore(run.text, offset)) ||
    (departs && /\s/u.test(run.text[offset]));
  // The pen starts at the run's own x, so a leading tab, which paints no glyph,
  // splits the run where the text after it lands.
  const departs = (index: number): boolean => {
    const previous = run.glyphs[index - 1];
    const expectedX = previous
      ? Math.fround(Math.fround(previous.x) + Math.fround(previous.advance)) : run.x;
    return Math.abs(run.glyphs[index].x - expectedX) > 0.0001;
  };
  // A face standing in for a family we only have widths for steps each glyph
  // to that family's x, so such a run is drawn piece by piece where it departs.
  const namedWidths = !run.letterSpacingPx && run.glyphs.some((glyph, index) => {
    const offset = glyph.cluster - run.start;
    return offset > 0 && offset < run.text.length && departs(index) &&
      !continuation(offset) && !boundary(offset, true);
  });
  // Right-to-left text laid out in a left-to-right line keeps the browser's own
  // bidi ordering, so only the pen discontinuities split it.
  const anchored = !namedWidths && !RTL_TEXT.test(run.text) && run.glyphs.every(
    (glyph, index) => index === 0 || glyph.cluster >= run.glyphs[index - 1].cluster,
  );
  let textStart = 0;
  let x = anchored && run.glyphs[0].cluster === run.start ? run.glyphs[0].x : run.x;
  let glyphStart = 0;
  const pushChunk = (textEnd: number, glyphEnd: number, endX: number): void => {
    const text = run.text.slice(textStart, textEnd);
    const chunk: (typeof chunks)[number] = { text, x };
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
    }
    chunks.push(chunk);
  };
  for (const [index, glyph] of run.glyphs.entries()) {
    const offset = glyph.cluster - run.start;
    if (
      offset > textStart &&
      offset < run.text.length &&
      (anchored ? !continuation(offset) && boundary(offset, departs(index)) : departs(index))
    ) {
      pushChunk(offset, index, glyph.x);
      textStart = offset;
      glyphStart = index;
      x = glyph.x;
    }
  }
  pushChunk(run.text.length, run.glyphs.length, run.x + run.width);
  return chunks;
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

