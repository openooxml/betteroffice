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
  return {
    row: grid.rowIndices?.[row] ?? grid.startRow + row,
    col: grid.colIndices?.[col] ?? grid.startCol + col,
  };
}

// local track index for an absolute row/col, or -1 when it is off the visible
// window described by grid.
function localIndex(
  start: number,
  indices: number[] | undefined,
  offsets: number[],
  absolute: number
): number {
  if (indices) {
    let left = 0;
    let right = indices.length - 1;
    while (left <= right) {
      const midpoint = (left + right) >> 1;
      const candidate = indices[midpoint];
      if (candidate === absolute) return midpoint;
      if (candidate < absolute) left = midpoint + 1;
      else right = midpoint - 1;
    }
    return -1;
  }
  const local = absolute - start;
  return local >= 0 && local < offsets.length - 1 ? local : -1;
}

/**
 * The pixel rect of a single cell in viewport-local coordinates, or `null` when
 * the cell is not within the visible window.
 */
export function cellRect(grid: GridMeta | undefined, row: number, col: number): Rect | null {
  if (!grid) return null;
  const lc = localIndex(grid.startCol, grid.colIndices, grid.colOffsets, col);
  const lr = localIndex(grid.startRow, grid.rowIndices, grid.rowOffsets, row);
  if (lc < 0 || lr < 0) return null;
  return {
    x: grid.colOffsets[lc],
    y: grid.rowOffsets[lr],
    w: grid.colOffsets[lc + 1] - grid.colOffsets[lc],
    h: grid.rowOffsets[lr + 1] - grid.rowOffsets[lr],
  };
}

function trackAddress(start: number, indices: number[] | undefined, local: number): number {
  return indices?.[local] ?? start + local;
}

function intersectingTracks(
  start: number,
  indices: number[] | undefined,
  tracks: number,
  first: number,
  last: number
): [number, number] | null {
  let leading = -1;
  let trailing = -1;
  for (let local = 0; local < tracks; local++) {
    const address = trackAddress(start, indices, local);
    if (address < first) continue;
    if (address > last) break;
    if (leading < 0) leading = local;
    trailing = local;
  }
  return leading < 0 ? null : [leading, trailing];
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

  const visibleCols = intersectingTracks(
    grid.startCol,
    grid.colIndices,
    cols,
    range.left,
    range.right
  );
  const visibleRows = intersectingTracks(
    grid.startRow,
    grid.rowIndices,
    rows,
    range.top,
    range.bottom
  );
  if (!visibleCols || !visibleRows) return null;
  const [l, r] = visibleCols;
  const [t, b] = visibleRows;
  return {
    x: grid.colOffsets[l],
    y: grid.rowOffsets[t],
    w: grid.colOffsets[r + 1] - grid.colOffsets[l],
    h: grid.rowOffsets[b + 1] - grid.rowOffsets[t],
  };
}
