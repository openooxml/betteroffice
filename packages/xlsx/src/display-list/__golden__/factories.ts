/**
 * Synthetic factories for the golden corpus.
 *
 * Build display lists and viewport states without touching the DOM, mirroring
 * the shapes the Rust renderer and the viewport math produce so the pinned
 * goldens exercise the real code paths (visible-range math, clip/align on text).
 */

import type { DisplayList, DrawCmd, Rect, TextAlign } from '../types';
import type { TrackOffsets, ViewportState } from '../../viewport/index';
import { uniformOffsets } from '../../viewport/index';

/** a solid fill command. */
export function fillRect(x: number, y: number, w: number, h: number, color: string): DrawCmd {
  return { op: 'fillRect', x, y, w, h, color };
}

/** a stroked line command, optionally dashed/dotted/double. */
export function line(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  width: number,
  color: string,
  style?: 'dashed' | 'dotted' | 'double'
): DrawCmd {
  return { op: 'line', x1, y1, x2, y2, width, color, style };
}

/** a clipped, aligned text command with optional style facets. */
export function text(
  x: number,
  y: number,
  value: string,
  opts: {
    fontSize?: number;
    color?: string;
    clip?: Rect;
    align?: TextAlign;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    fontFamily?: string;
  } = {}
): DrawCmd {
  return {
    op: 'text',
    x,
    y,
    text: value,
    fontSize: opts.fontSize ?? 12,
    color: opts.color ?? '#000000',
    clip: opts.clip,
    align: opts.align,
    bold: opts.bold,
    italic: opts.italic,
    underline: opts.underline,
    strike: opts.strike,
    fontFamily: opts.fontFamily,
  };
}

/**
 * Horizontal + vertical gridlines spanning a uniform grid, as the renderer
 * emits them for a plain sheet region.
 */
export function gridLines(
  cols: number,
  rows: number,
  colWidth: number,
  rowHeight: number,
  color = '#d0d0d0'
): DrawCmd[] {
  const width = cols * colWidth;
  const height = rows * rowHeight;
  const cmds: DrawCmd[] = [];
  for (let c = 0; c <= cols; c++) cmds.push(line(c * colWidth, 0, c * colWidth, height, 1, color));
  for (let r = 0; r <= rows; r++) cmds.push(line(0, r * rowHeight, width, r * rowHeight, 1, color));
  return cmds;
}

/**
 * A minimal grid display list: gridlines plus one clipped, left-aligned label
 * per top-row cell.
 */
export function makeGridDisplayList(
  cols: number,
  rows: number,
  colWidth: number,
  rowHeight: number
): DisplayList {
  const commands = gridLines(cols, rows, colWidth, rowHeight);
  for (let c = 0; c < cols; c++) {
    const clip: Rect = { x: c * colWidth, y: 0, w: colWidth, h: rowHeight };
    commands.push(text(c * colWidth + 2, rowHeight - 4, `C${c}`, { clip, align: 'left' }));
  }
  return { width: cols * colWidth, height: rows * rowHeight, commands };
}

/** a viewport state with sensible defaults, overridable per field. */
export function makeViewport(overrides: Partial<ViewportState> = {}): ViewportState {
  return {
    scrollX: 0,
    scrollY: 0,
    width: 400,
    height: 300,
    dpr: 1,
    frozenRows: 0,
    frozenCols: 0,
    ...overrides,
  };
}

/** uniform column/row offsets for a `cols` × `rows` grid. */
export function makeUniformGridOffsets(
  cols: number,
  rows: number,
  colWidth: number,
  rowHeight: number
): { colOffsets: TrackOffsets; rowOffsets: TrackOffsets } {
  return {
    colOffsets: uniformOffsets(cols, colWidth),
    rowOffsets: uniformOffsets(rows, rowHeight),
  };
}
