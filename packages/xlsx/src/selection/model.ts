/**
 * Selection model: pure transforms over {@link Selection} data. Every function
 * returns a fresh value and never mutates its input, so state lives with the
 * chrome and this layer stays a deterministic reducer core.
 */

import type { CellAddr, CellRange, Direction, MoveOpts, Selection, SelectionLimits } from './types';

function clamp(value: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, value));
}

/**
 * A collapsed selection at a single cell.
 */
export function selectionAt(addr: CellAddr): Selection {
  return { anchor: { ...addr }, focus: { ...addr } };
}

/**
 * The normalized inclusive rectangle spanned by a selection's anchor and focus.
 */
export function normalizeRange(sel: Selection): CellRange {
  return {
    top: Math.min(sel.anchor.row, sel.focus.row),
    bottom: Math.max(sel.anchor.row, sel.focus.row),
    left: Math.min(sel.anchor.col, sel.focus.col),
    right: Math.max(sel.anchor.col, sel.focus.col),
  };
}

/**
 * Whether a cell falls inside a selection's normalized range.
 */
export function rangeContains(sel: Selection, addr: CellAddr): boolean {
  const r = normalizeRange(sel);
  return addr.row >= r.top && addr.row <= r.bottom && addr.col >= r.left && addr.col <= r.right;
}

/**
 * Number of cells covered by a selection.
 */
export function cellCount(sel: Selection): number {
  const r = normalizeRange(sel);
  return (r.bottom - r.top + 1) * (r.right - r.left + 1);
}

/**
 * Extend the selection to a new focus, holding the anchor fixed (shift-click /
 * shift-arrow semantics). The address is clamped into the grid when limits pass.
 */
export function extendTo(sel: Selection, addr: CellAddr, limits?: SelectionLimits): Selection {
  const focus = limits
    ? { row: clamp(addr.row, 0, limits.rows - 1), col: clamp(addr.col, 0, limits.cols - 1) }
    : { ...addr };
  return { anchor: { ...sel.anchor }, focus };
}

function delta(direction: Direction, step: number): CellAddr {
  switch (direction) {
    case 'up':
      return { row: -step, col: 0 };
    case 'down':
      return { row: step, col: 0 };
    case 'left':
      return { row: 0, col: -step };
    case 'right':
      return { row: 0, col: step };
  }
}

/**
 * Move the focus one (or `step`) cells in a direction. When `extend` is set the
 * anchor holds and the range grows; otherwise the whole selection collapses onto
 * the new focus. The result is clamped to the grid bounds in `opts.limits`.
 */
export function moveFocus(sel: Selection, direction: Direction, opts: MoveOpts): Selection {
  const { limits } = opts;
  const step = opts.step ?? 1;
  const d = delta(direction, step);
  const next: CellAddr = {
    row: clamp(sel.focus.row + d.row, 0, limits.rows - 1),
    col: clamp(sel.focus.col + d.col, 0, limits.cols - 1),
  };
  return opts.extend ? extendTo(sel, next) : selectionAt(next);
}
