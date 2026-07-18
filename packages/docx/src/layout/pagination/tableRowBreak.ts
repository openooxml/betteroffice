/**
 * Table row-break geometry.
 *
 * Word lets a table row break across a page boundary ("allow row to break
 * across pages", on by default). When a row is taller than the space left on
 * the page, the portion that fits stays and the rest continues on the next
 * page — broken between whole text lines, never through a glyph.
 *
 * This module computes, per row, the set of safe break offsets (the y of every
 * line bottom across the row's content, including vertically-merged cells that
 * span into the row) so the paginator can snap a break to the deepest whole
 * line that still fits.
 */

import type { TableBlock, TableExtent } from './types';
import { resolveCellGrid } from './tableWidthUtils';
import { layoutCellContent } from './cellBlockLayout';

/**
 * per-table break geometry consumed by snapRowBreak.
 */
export interface TableRowBreakInfo {
  /** Cumulative y of the top of each row; rowTops[rows.length] is the table height. */
  rowTops: number[];
  /**
   * Per-row sorted, de-duplicated line-bottom offsets (relative to the row top)
   * at which a break is clean. Always includes the row's full height as the
   * final boundary.
   */
  breakOffsets: number[][];
}

/**
 * resolves the cell grid once and collects, per row, every whole-line bottom
 * a break is allowed to snap to.
 */
export function buildTableRowBreakInfo(block: TableBlock, measure: TableExtent): TableRowBreakInfo {
  const rowCount = measure.rows.length;
  // True (unrounded) cumulative row offsets — the paginator splits against
  // exact measured heights. The painter has a sibling `buildRowYPositions`
  // that rounds to whole pixels for crisp borders; keep the two SEPARATE
  // (don't "dedupe") or you break either break-offset alignment or crispness.
  const rowTops: number[] = [];
  let acc = 0;
  for (let r = 0; r < rowCount; r++) {
    rowTops.push(acc);
    acc += measure.rows[r]?.height ?? 0;
  }
  rowTops.push(acc);

  // Use the shared grid resolution so "which cells cover row r" matches the
  // measurer and painter. A cell starting in row `sr` with rowSpan covers
  // rows [sr, sr + rowSpan); a merged cell spills its line bottoms into the
  // rows below its restart row.
  const resolved = resolveCellGrid(block);
  const breakOffsets: number[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const rowHeight = measure.rows[r]?.height ?? 0;
    const offsets = new Set<number>();
    offsets.add(rowHeight); // a row boundary is always a clean break
    const unbreakableRanges: Array<{ top: number; bottom: number }> = [];

    for (const g of resolved) {
      if (g.rowIndex > r || g.rowIndex + g.rowSpan - 1 < r) continue;
      const sourceCell = block.rows[g.rowIndex]?.cells?.[g.cellIndex];
      const measuredCell = measure.rows[g.rowIndex]?.cells?.[g.cellIndex];
      if (!sourceCell || !measuredCell) continue;
      // OOXML/TableNormal default top padding is 0 (matches measureTable).
      const padTop = sourceCell.padding?.top ?? 0;
      const { flatBottoms, unbreakableRanges: cellRanges } = layoutCellContent(
        sourceCell.blocks,
        measuredCell.blocks,
        padTop
      );
      // Map cell-content y (relative to the cell/region top at rowTops[startRow])
      // into this row's coordinate space (relative to rowTops[r]).
      const shift = rowTops[r] - rowTops[g.rowIndex];
      for (const b of flatBottoms) {
        const off = b - shift;
        if (off > 0 && off < rowHeight) offsets.add(off);
      }
      for (const range of cellRanges) {
        unbreakableRanges.push({ top: range.top - shift, bottom: range.bottom - shift });
      }
    }
    breakOffsets.push(
      [...offsets]
        .filter(
          (offset) =>
            offset === rowHeight ||
            !unbreakableRanges.some((range) => offset > range.top && offset < range.bottom)
        )
        .sort((a, b) => a - b)
    );
  }

  return { rowTops, breakOffsets };
}

/**
 * Given a row and how much of it has already been placed (`fromOffset`),
 * return how many more px can be placed ending on a whole line, without
 * exceeding `maxSlice`. Returns 0 when not even the first line fits.
 */
export function snapRowBreak(
  info: TableRowBreakInfo,
  rowIndex: number,
  fromOffset: number,
  maxSlice: number
): number {
  const offsets = info.breakOffsets[rowIndex];
  if (!offsets || offsets.length === 0) return 0;
  const limit = fromOffset + maxSlice;
  let best = 0;
  for (const off of offsets) {
    if (off <= fromOffset) continue;
    if (off <= limit) best = off - fromOffset;
    else break;
  }
  return best;
}
