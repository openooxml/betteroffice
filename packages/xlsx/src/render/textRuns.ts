import type { TextPiece } from '../../../../shared/office-session/xlsxTextRuns';
import { wrapSpans } from './wrap';

export { hasRightToLeft, runPieces, type PaintedTextRun, type TextPiece } from '../../../../shared/office-session/xlsxTextRuns';

/**
 * The pieces of each line a wrapped command paints, split where its lines break; `measure` gives the width of part of
 * a piece's text in its font. Follows `wrap_run_lines` in the Rust display list.
 */
export function wrapRunLines(
  text: string,
  pieces: TextPiece[],
  width: number,
  measure: (piece: TextPiece, part: string) => number,
): TextPiece[][] {
  const within = (start: number, end: number): TextPiece[] => pieces.flatMap((piece) => {
    const from = Math.max(piece.start, start);
    const to = Math.min(piece.end, end);
    return from < to ? [{ ...piece, text: text.slice(from, to), start: from, end: to }] : [];
  });
  const spans = wrapSpans(text, width, (start, end) => within(start, end).reduce((sum, piece) => sum + measure(piece, piece.text), 0));
  return spans.map(([start, end]) => within(start, end));
}
