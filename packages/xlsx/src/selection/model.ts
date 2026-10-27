/**
 * Selection model: pure transforms over {@link Selection} data. Every function
 * returns a fresh value and never mutates its input, so state lives with the
 * chrome and this layer stays a deterministic reducer core.
 */

import type { CellAddr, CellRange, Direction, MoveOpts, Selection, SelectionLimits, SelectionMerges } from './types';

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
 * The anchor of the merged area covering `addr`, or `addr` itself.
 */
export function mergeAnchor(addr: CellAddr, merges?: SelectionMerges): CellAddr {
  const merged = merges?.mergeAt(addr.row, addr.col);
  return merged ? { row: merged.top, col: merged.left } : { ...addr };
}

/**
 * Move the focus one (or `step`) cells in a direction. When `extend` is set the
 * anchor holds and the range grows; otherwise the whole selection collapses onto
 * the new focus. The result is clamped to the grid bounds in `opts.limits`. A
 * merged area counts as one cell: the move starts from its far edge, and a
 * collapsed selection that lands inside one lands on its anchor.
 */
export function moveFocus(sel: Selection, direction: Direction, opts: MoveOpts): Selection {
  const { limits, hidden, merges } = opts;
  const vertical = direction === 'up' || direction === 'down';
  const back = direction === 'up' || direction === 'left';
  const merged = merges?.mergeAt(sel.focus.row, sel.focus.col);
  const edge = merged ? vertical ? back ? merged.top : merged.bottom : back ? merged.left : merged.right
    : vertical ? sel.focus.row : sel.focus.col;
  const limit = vertical ? limits.rows : limits.cols;
  let position = edge;
  if (!hidden) position = clamp(edge + (back ? -1 : 1) * (opts.step ?? 1), 0, limit - 1);
  else {
    const isHidden = vertical ? hidden.isRowHidden : hidden.isColHidden;
    let remaining = opts.step ?? 1;
    while (remaining > 0) {
      let candidate = position + (back ? -1 : 1);
      while (candidate >= 0 && candidate < limit && isHidden(candidate)) candidate += back ? -1 : 1;
      if (candidate < 0 || candidate >= limit) break;
      position = candidate;
      remaining -= 1;
    }
  }
  const moved = position !== edge;
  const target = vertical ? { row: position, col: sel.focus.col } : { row: sel.focus.row, col: position };
  const landed = moved ? target : sel.focus;
  const next = opts.extend ? landed : mergeAnchor(landed, merges);
  if (hidden && next.row === sel.focus.row && next.col === sel.focus.col) return sel;
  return opts.extend ? extendTo(sel, next) : selectionAt(next);
}
