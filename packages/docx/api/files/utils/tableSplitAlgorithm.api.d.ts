// @betteroffice/docx declaration utils/tableSplitAlgorithm.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Shared table split algorithm — model-agnostic core logic.
 *
 * This module contains the pure layout computation for splitting a table cell:
 * column width redistribution, neighbor span adjustment, and new-cell placement.
 * The Document-model path (TableToolbar.tsx) delegates to these functions.
 */
/** A cell's position and span within the logical grid. */
export interface CellAnchor<T> {
    /** Opaque payload — the caller's cell type (TableCell, etc.) */
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
export declare function sumColumnWidths(widths: number[], start: number, span: number): number;
/**
 * Redistribute column widths when splitting a cell's column span.
 *
 * @param existing   Current column widths array for the whole table.
 * @param startCol   First column of the cell being split.
 * @param currentSpan Current column span of the cell.
 * @param targetSpan  Desired column span after split.
 * @returns New column widths array with the split applied.
 */
export declare function redistributeColumnWidths(existing: number[], startCol: number, currentSpan: number, targetSpan: number): number[];
/**
 * Compute the new anchor layout after splitting a target cell.
 *
 * This is the core algorithm. It adjusts neighbor spans, shifts positions for
 * inserted rows/cols, and creates placeholder anchors for the new split cells.
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
export declare function computeSplitLayout<T>(anchors: CellAnchor<T>[], target: CellAnchor<T>, rows: number, cols: number, totalRows: number, createSplitCellData: (isOriginal: boolean, rowOffset: number, colOffset: number) => T): SplitLayoutResult<T>;
/**
 * Build lookup maps from an anchor list — by start position and by covered slot.
 */
export declare function buildAnchorMaps<T>(anchors: CellAnchor<T>[]): {
    byStart: Map<string, CellAnchor<T>>;
    byCoveredSlot: Map<string, CellAnchor<T>>;
};
/**
 * Compute the initial dialog values for a split-cell dialog.
 */
export declare function computeSplitDialogDefaults(rowspan: number, colspan: number): {
    minRows: number;
    minCols: number;
    initialRows: number;
    initialCols: number;
};
