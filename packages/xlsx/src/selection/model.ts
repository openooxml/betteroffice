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
 * The explicit bounds or normalized rectangle spanned by the endpoints.
 */
export function normalizeRange(sel: Selection): CellRange {
  if (sel.range) return { ...sel.range };
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

/**
 * Move the focus one (or `step`) cells in a direction. When `extend` is set the
 * anchor holds and the range grows; otherwise the whole selection collapses onto
 * the new focus. The result is clamped to the grid bounds in `opts.limits`.
 */
export function moveFocus(sel: Selection, direction: Direction, opts: MoveOpts): Selection {
  const { limits, hidden } = opts;
  if (!hidden) {
    const step = opts.step ?? 1;
    const next = {
      row: clamp(sel.focus.row + (direction === 'down' ? step : direction === 'up' ? -step : 0), 0, limits.rows - 1),
      col: clamp(sel.focus.col + (direction === 'right' ? step : direction === 'left' ? -step : 0), 0, limits.cols - 1),
    };
    return opts.extend ? extendTo(sel, next) : selectionAt(next);
  }
  const vertical = direction === 'up' || direction === 'down';
  const back = direction === 'up' || direction === 'left';
  const isHidden = vertical ? hidden.isRowHidden : hidden.isColHidden;
  const limit = vertical ? limits.rows : limits.cols;
  let position = vertical ? sel.focus.row : sel.focus.col;
  let remaining = opts.step ?? 1;
  while (remaining > 0) {
    let candidate = position + (back ? -1 : 1);
    while (candidate >= 0 && candidate < limit && isHidden(candidate)) candidate += back ? -1 : 1;
    if (candidate < 0 || candidate >= limit) break;
    position = candidate;
    remaining -= 1;
  }
  const next = vertical ? { row: position, col: sel.focus.col } : { row: sel.focus.row, col: position };
  if (next.row === sel.focus.row && next.col === sel.focus.col) return sel;
  return opts.extend ? extendTo(sel, next) : selectionAt(next);
}
