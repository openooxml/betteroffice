/**
 * Pure hit-testing over a frame's {@link GridMeta}: map a viewport-local point
 * to a cell, and map cells/ranges back to their pixel rects for overlay
 * placement. Binary search over the boundary offsets; DOM-free so click
 * handling and the selection overlay share one geometry source with the painter.
 */

import type { GridMeta, Rect } from '../display-list/types';
import type { CellAddr, CellRange } from '../selection/index';

// largest index i in [0, offsets.length-2] with offsets[i] <= target, or -1 when
// target is left of the first edge. offsets is ascending; the last entry is the
// one-past-end trailing edge, so a hit requires target < offsets[last].
function trackIndexAt(offsets: number[], target: number): number {
  const tracks = offsets.length - 1;
  if (tracks <= 0) return -1;
  if (target < offsets[0] || target >= offsets[tracks]) return -1;
  let left = 0;
  let right = tracks - 1;
  while (left < right) {
    const midpoint = (left + right + 1) >> 1;
    if (offsets[midpoint] <= target) left = midpoint;
    else right = midpoint - 1;
  }
  return left;
}

/**
 * The cell under a viewport-local point, or `null` when the point is outside the
 * visible grid (or the frame carries no grid metadata).
 */
export function cellAtPoint(grid: GridMeta | undefined, x: number, y: number): CellAddr | null {
  if (!grid) return null;
  const col = trackIndexAt(grid.colOffsets, x);
  const row = trackIndexAt(grid.rowOffsets, y);
  if (col < 0 || row < 0) return null;
  return { row: grid.startRow + row, col: grid.startCol + col };
}

// local track index for an absolute row/col, or -1 when it is off the visible
// window described by grid.
function localIndex(start: number, offsets: number[], absolute: number): number {
  const local = absolute - start;
  return local >= 0 && local < offsets.length - 1 ? local : -1;
}

/**
 * The pixel rect of a single cell in viewport-local coordinates, or `null` when
 * the cell is not within the visible window.
 */
export function cellRect(grid: GridMeta | undefined, row: number, col: number): Rect | null {
  if (!grid) return null;
  const lc = localIndex(grid.startCol, grid.colOffsets, col);
  const lr = localIndex(grid.startRow, grid.rowOffsets, row);
  if (lc < 0 || lr < 0) return null;
  return {
    x: grid.colOffsets[lc],
    y: grid.rowOffsets[lr],
    w: grid.colOffsets[lc + 1] - grid.colOffsets[lc],
    h: grid.rowOffsets[lr + 1] - grid.rowOffsets[lr],
  };
}

// clamp an inclusive [value..] index into the visible track window [0, tracks).
function clampLocal(start: number, tracks: number, absolute: number): number {
  return Math.max(0, Math.min(tracks - 1, absolute - start));
}

/**
 * The pixel rect covering the visible part of a cell range, clipped to the
 * window. `null` when the range lies entirely outside the visible grid — so the
 * selection overlay is hidden rather than mispainted when scrolled away.
 */
export function rangeRect(grid: GridMeta | undefined, range: CellRange): Rect | null {
  if (!grid) return null;
  const cols = grid.colOffsets.length - 1;
  const rows = grid.rowOffsets.length - 1;
  if (cols <= 0 || rows <= 0) return null;

  const lastCol = grid.startCol + cols - 1;
  const lastRow = grid.startRow + rows - 1;
  // reject ranges that do not intersect the visible window on either axis.
  if (range.right < grid.startCol || range.left > lastCol) return null;
  if (range.bottom < grid.startRow || range.top > lastRow) return null;

  const l = clampLocal(grid.startCol, cols, range.left);
  const r = clampLocal(grid.startCol, cols, range.right);
  const t = clampLocal(grid.startRow, rows, range.top);
  const b = clampLocal(grid.startRow, rows, range.bottom);
  return {
    x: grid.colOffsets[l],
    y: grid.rowOffsets[t],
    w: grid.colOffsets[r + 1] - grid.colOffsets[l],
    h: grid.rowOffsets[b + 1] - grid.rowOffsets[t],
  };
}
