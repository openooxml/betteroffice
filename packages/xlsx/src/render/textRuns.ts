import type { TextCmd, TextRun } from '../display-list/types';

/** A text run as a painter may carry it, with the CSS family list a PNG export resolved for it. */
export type PaintedTextRun = TextRun & { cssFontFamily?: string };

/** A span of a text command's text with the command it paints as: the command's own font with the run's properties over it. */
export interface TextPiece {
  text: string;
  cmd: TextCmd & { cssFontFamily?: string };
  baseline?: TextRun['baseline'];
}

const RIGHT_TO_LEFT = /[֐-ࣿיִ-﷿ﹰ-ﻼ‏‫‮⁧\u{10800}-\u{10fff}\u{1e800}-\u{1efff}]/u;

/** Whether `text` holds a right-to-left letter or a control that opens right-to-left text. */
export function hasRightToLeft(text: string): boolean {
  return RIGHT_TO_LEFT.test(text);
}

/** `offset` clamped into `text`, moved back to the start of a surrogate pair it would split. */
function boundary(text: string, offset: number): number {
  const at = Math.max(0, Math.min(text.length, Math.floor(offset)));
  const low = text.charCodeAt(at);
  const high = text.charCodeAt(at - 1);
  return at > 0 && at < text.length && low >= 0xdc00 && low <= 0xdfff && high >= 0xd800 && high <= 0xdbff ? at - 1 : at;
}

function underlineKind(underline: boolean, style: TextRun['underlineStyle']): TextCmd['underlineStyle'] {
  return underline && style !== undefined && style !== 'single' ? style : undefined;
}

function styled(cmd: TextCmd & { cssFontFamily?: string }, run?: PaintedTextRun): TextPiece['cmd'] {
  const underline = run?.underline ?? cmd.underline === true;
  const cssFontFamily = run?.cssFontFamily ?? (run?.fontFamily === undefined ? cmd.cssFontFamily : undefined);
  return {
    ...cmd,
    runs: undefined,
    fontFamily: run?.fontFamily ?? cmd.fontFamily,
    cssFontFamily,
    fontSize: run?.fontSize ?? cmd.fontSize,
    bold: run?.bold ?? cmd.bold === true,
    italic: run?.italic ?? cmd.italic === true,
    underline,
    underlineStyle: underlineKind(underline, run?.underlineStyle ?? cmd.underlineStyle),
    strike: run?.strike ?? cmd.strike === true,
    color: run?.color ?? cmd.color,
  };
}

function sameStyle(a: TextPiece, b: TextPiece): boolean {
  return a.cmd.fontFamily === b.cmd.fontFamily && a.cmd.cssFontFamily === b.cmd.cssFontFamily
    && a.cmd.fontSize === b.cmd.fontSize && a.cmd.bold === b.cmd.bold && a.cmd.italic === b.cmd.italic
    && a.cmd.underline === b.cmd.underline && a.cmd.underlineStyle === b.cmd.underlineStyle
    && a.cmd.strike === b.cmd.strike && a.cmd.color === b.cmd.color && a.baseline === b.baseline;
}

/**
 * The pieces a painter draws for a text command left to right on one baseline, or `undefined` when it paints the
 * command as one plain run: no runs, runs that all carry the command's own font, and the cases the pieces do not
 * cover yet (wrapped, rotated, stacked, number-formatted or highlighted text, and text with a right-to-left
 * character). Neighbouring pieces in the same font join into one. Follows `run_pieces` in the Rust display list.
 */
export function runPieces(cmd: TextCmd & { cssFontFamily?: string }): TextPiece[] | undefined {
  const runs = cmd.runs as PaintedTextRun[] | undefined;
  if (!runs?.length || cmd.wrap || cmd.stacked || (cmd.rotationDeg !== undefined && cmd.rotationDeg !== null)
    || cmd.numberFormat || cmd.highlight || cmd.dashedUnderline || hasRightToLeft(cmd.text)) return undefined;
  const base: TextPiece = { text: '', cmd: styled(cmd) };
  const pieces: TextPiece[] = [];
  const push = (from: number, to: number, style: TextPiece) => {
    if (to <= from) return;
    const last = pieces[pieces.length - 1];
    if (last && sameStyle(last, style)) last.text += cmd.text.slice(from, to);
    else pieces.push({ ...style, text: cmd.text.slice(from, to) });
  };
  let cursor = 0;
  for (const run of runs) {
    const start = Math.max(boundary(cmd.text, run.start), cursor);
    const end = boundary(cmd.text, run.end);
    if (end <= start) continue;
    push(cursor, start, base);
    push(start, end, { text: '', cmd: styled(cmd, run), baseline: run.baseline });
    cursor = end;
  }
  push(cursor, cmd.text.length, base);
  if (pieces.length === 0 || (pieces.length === 1 && sameStyle(pieces[0]!, base))) return undefined;
  return pieces;
}
