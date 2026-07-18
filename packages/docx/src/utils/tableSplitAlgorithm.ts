/**
 * Shared table split algorithm — model-agnostic core logic.
 *
 * This module contains the pure layout computation for splitting a table cell:
 * column width redistribution, neighbor span adjustment, and new-cell placement.
 * Both the ProseMirror path (tableSplit.ts) and the Document-model path
 * (TableToolbar.tsx) delegate to these functions.
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** A cell's position and span within the logical grid. */
export interface CellAnchor<T> {
  /** Opaque payload — the caller's cell type (PMNode, TableCell, etc.) */
  data: T;
  row: number;
  col: number;
  rowspan: number;
  colspan: number;
}

/** Parameters describing the split target. */
export interface SplitTarget {
  row: number;
  col: number;
  rowspan: number;
  colspan: number;
}

/** Result of `computeSplitLayout`. */
export interface SplitLayoutResult<T> {
  /** All anchors after the split (neighbors adjusted + new split cells). */
  anchors: CellAnchor<T>[];
  deltaRows: number;
  deltaCols: number;
  newRowCount: number;
}

// ---------------------------------------------------------------------------
// Column width math
// ---------------------------------------------------------------------------

export function sumColumnWidths(widths: number[], start: number, span: number): number {
  let total = 0;
  for (let i = start; i < start + span && i < widths.length; i++) {
    total += widths[i];
  }
  return total;
}

/**
 * Redistribute column widths when splitting a cell's column span.
 *
 * @param existing   Current column widths array for the whole table.
 * @param startCol   First column of the cell being split.
 * @param currentSpan Current column span of the cell.
 * @param targetSpan  Desired column span after split.
 * @returns New column widths array with the split applied.
 */
export function redistributeColumnWidths(
  existing: number[],
  startCol: number,
  currentSpan: number,
  targetSpan: number
): number[] {
  const sliceWidth = sumColumnWidths(existing, startCol, currentSpan);
  const segmentWidth = Math.floor(sliceWidth / Math.max(targetSpan, 1));
  const remainder = sliceWidth - segmentWidth * targetSpan;
  const replacement = Array.from(
    { length: targetSpan },
    (_, i) => segmentWidth + (i < remainder ? 1 : 0)
  );

  return [
    ...existing.slice(0, startCol),
    ...replacement,
    ...existing.slice(startCol + currentSpan),
  ];
}

// ---------------------------------------------------------------------------
// Core split layout computation
// ---------------------------------------------------------------------------

/**
 * Compute the new anchor layout after splitting a target cell.
 *
 * This is the core algorithm shared between ProseMirror and Document-model
 * paths. It adjusts neighbor spans, shifts positions for inserted rows/cols,
 * and creates placeholder anchors for the new split cells.
 *
 * @param anchors     All cell anchors in the current table.
 * @param target      The anchor being split.
 * @param rows        Number of rows the target should become.
 * @param cols        Number of columns the target should become.
 * @param totalRows   Current total row count.
 * @param createSplitCellData  Factory to create `data` for each new split cell.
 *   Called with `(isOriginal, rowOffset, colOffset)` — `isOriginal` is true for
 *   the top-left cell that retains the original content.
 */
export function computeSplitLayout<T>(
  anchors: CellAnchor<T>[],
  target: CellAnchor<T>,
  rows: number,
  cols: number,
  totalRows: number,
  createSplitCellData: (isOriginal: boolean, rowOffset: number, colOffset: number) => T
): SplitLayoutResult<T> {
  const deltaRows = rows - target.rowspan;
  const deltaCols = cols - target.colspan;
  const newRowCount = totalRows + deltaRows;

  const targetRowEnd = target.row + target.rowspan;
  const targetColEnd = target.col + target.colspan;

  const nextAnchors: CellAnchor<T>[] = [];

  // Adjust existing anchors (skip the target itself)
  for (const anchor of anchors) {
    if (anchor === target) continue;

    const rowEnd = anchor.row + anchor.rowspan;
    const colEnd = anchor.col + anchor.colspan;
    const rowIntersectsBand = anchor.row < targetRowEnd && rowEnd > target.row;
    const colIntersectsBand = anchor.col < targetColEnd && colEnd > target.col;

    nextAnchors.push({
      data: anchor.data,
      row: anchor.row >= targetRowEnd ? anchor.row + deltaRows : anchor.row,
      col: anchor.col >= targetColEnd ? anchor.col + deltaCols : anchor.col,
      rowspan:
        anchor.rowspan + (deltaRows > 0 && rowIntersectsBand && !colIntersectsBand ? deltaRows : 0),
      colspan:
        anchor.colspan + (deltaCols > 0 && colIntersectsBand && !rowIntersectsBand ? deltaCols : 0),
    });
  }

  // Add new split cells
  for (let rowOffset = 0; rowOffset < rows; rowOffset++) {
    for (let colOffset = 0; colOffset < cols; colOffset++) {
      const isOriginal = rowOffset === 0 && colOffset === 0;
      nextAnchors.push({
        data: createSplitCellData(isOriginal, rowOffset, colOffset),
        row: target.row + rowOffset,
        col: target.col + colOffset,
        rowspan: 1,
        colspan: 1,
      });
    }
  }

  return { anchors: nextAnchors, deltaRows, deltaCols, newRowCount };
}

/**
 * Build lookup maps from an anchor list — by start position and by covered slot.
 */
export function buildAnchorMaps<T>(anchors: CellAnchor<T>[]): {
  byStart: Map<string, CellAnchor<T>>;
  byCoveredSlot: Map<string, CellAnchor<T>>;
} {
  const byStart = new Map<string, CellAnchor<T>>();
  const byCoveredSlot = new Map<string, CellAnchor<T>>();

  for (const anchor of anchors) {
    byStart.set(`${anchor.row}-${anchor.col}`, anchor);
    for (let row = anchor.row; row < anchor.row + anchor.rowspan; row++) {
      for (let col = anchor.col; col < anchor.col + anchor.colspan; col++) {
        byCoveredSlot.set(`${row}-${col}`, anchor);
      }
    }
  }

  return { byStart, byCoveredSlot };
}

/**
 * Compute the initial dialog values for a split-cell dialog.
 */
export function computeSplitDialogDefaults(rowspan: number, colspan: number) {
  const isMerged = rowspan > 1 || colspan > 1;
  return {
    minRows: rowspan,
    minCols: colspan,
    initialRows: rowspan,
    initialCols: isMerged ? colspan : colspan + 1,
  };
}
