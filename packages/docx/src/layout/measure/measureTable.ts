/**
 * Shared table measurement helper.
 *
 * Both React's PagedEditor and Vue's useDocxEditor measure tables the
 * same way: resolve column widths through `tableWidthUtils`, track
 * column occupancy across vertically merged cells, then call back into
 * the adapter's `measureBlock` for each cell's contained blocks.
 *
 * This module lives in core (not in either adapter) so a fix on one
 * side automatically applies to both. The recursive cell-content
 * measurement is delegated to the caller via `measureBlock` because
 * each adapter's block coverage differs (React supports floating
 * zones, footnotes, textBoxes; Vue is a subset).
 */

import type {
  LayoutBlock,
  BlockExtent,
  TableBlock,
  TableCell,
  TableExtent,
} from '../pagination/types';
import {
  resolveCellGrid,
  resolveTableColumnWidths,
  resolveTableWidthPx,
} from '../pagination/tableWidthUtils';

/** Word's TableNormal default — 108 twips ≈ 7px. */
const DEFAULT_CELL_PADDING_X = 7;
/** OOXML/TableNormal default for top/bottom cell padding. */
const DEFAULT_CELL_PADDING_Y = 0;

/**
 * Visual height of a single block inside a table cell.
 *
 * A one-line paragraph that contains only image runs is laid out at the
 * image's intrinsic height (plus the paragraph's explicit spacing), not
 * a full text line — matching Word's per-cell layout. Everything else
 * uses the measured `totalHeight` / `height`.
 */
export function measureTableCellBlockVisualHeight(
  block: LayoutBlock,
  blockMeasure: BlockExtent
): number {
  if (block.kind !== 'paragraph' || blockMeasure.kind !== 'paragraph') {
    if ('totalHeight' in blockMeasure) return blockMeasure.totalHeight;
    if ('height' in blockMeasure) return blockMeasure.height;
    return 0;
  }

  const nonEmptyRuns = block.runs.filter((run) => run.kind !== 'text' || run.text.length > 0);
  const imageOnlySingleLine =
    blockMeasure.lines.length === 1 &&
    nonEmptyRuns.length > 0 &&
    nonEmptyRuns.every((run) => run.kind === 'image');

  if (!imageOnlySingleLine) {
    return blockMeasure.totalHeight;
  }

  const maxImageHeight = nonEmptyRuns.reduce(
    (h, run) => (run.kind === 'image' ? Math.max(h, run.height) : h),
    0
  );
  const spacingBefore = block.attrs?.spacing?.before ?? 0;
  const spacingAfter = block.attrs?.spacing?.after ?? 0;
  return spacingBefore + maxImageHeight + spacingAfter;
}

/** Combined top + bottom border width of a cell, in pixels. */
function getTableCellVerticalBorderHeight(cell: TableCell | undefined): number {
  const top = cell?.borders?.top?.width ?? 0;
  const bottom = cell?.borders?.bottom?.width ?? 0;
  return top + bottom;
}

/**
 * measure a `TableBlock` against a content-width budget: resolve the column
 * grid, measure every cell's blocks, then size rows — with vertically merged
 * cells distributing their surplus height in a second pass.
 *
 * `measureBlock` is the per-cell-content callback; the adapter supplies its
 * own so block coverage stays per-renderer.
 */
export function measureTableBlock(
  tableBlock: TableBlock,
  contentWidth: number,
  measureBlock: (block: LayoutBlock, contentWidth: number) => BlockExtent
): TableExtent {
  const explicitWidthPx = resolveTableWidthPx(tableBlock.width, tableBlock.widthType, contentWidth);
  const targetWidth = explicitWidthPx ?? contentWidth;
  const columnWidths = resolveTableColumnWidths(tableBlock, contentWidth);

  // Resolve each cell's grid column once (shared with the painter and the
  // row-break paginator) so column-width assignment honors vertically-merged
  // cells from earlier rows without re-walking the occupancy grid here.
  const columnIndexByCell = new Map<string, number>();
  for (const g of resolveCellGrid(tableBlock)) {
    columnIndexByCell.set(`${g.rowIndex}-${g.cellIndex}`, g.columnIndex);
  }

  const rows = tableBlock.rows.map((row, rowIdx) => {
    return {
      cells: row.cells.map((cell, cellIdx) => {
        const colSpan = cell.colSpan ?? 1;
        const columnIndex = columnIndexByCell.get(`${rowIdx}-${cellIdx}`) ?? 0;
        let cellWidth = 0;
        for (let c = 0; c < colSpan && columnIndex + c < columnWidths.length; c++) {
          cellWidth += columnWidths[columnIndex + c] ?? 0;
        }
        if (cellWidth === 0) {
          cellWidth =
            (cell.width && cell.width > 0
              ? cell.width
              : resolveTableWidthPx(cell.widthValue, cell.widthType, targetWidth)) ?? 100;
        }

        const padLeft = cell.padding?.left ?? DEFAULT_CELL_PADDING_X;
        const padRight = cell.padding?.right ?? DEFAULT_CELL_PADDING_X;
        const cellContentWidth = Math.max(1, cellWidth - padLeft - padRight);

        return {
          blocks: cell.blocks.map((b) => measureBlock(b, cellContentWidth)),
          width: cellWidth,
          height: 0,
          colSpan: cell.colSpan,
          rowSpan: cell.rowSpan,
        };
      }),
      height: 0,
    };
  });

  // First pass: measure every cell's content height, and size each row from
  // the cells that START in it. A vertically-merged cell (rowSpan > 1) does
  // NOT inflate its restart row here — Word keeps the restart row at the
  // height its own single-row cells need and pushes the merged cell's surplus
  // content down into the span (handled in the second pass below). Without
  // this, the restart row balloons to hold the whole merged column and the
  // continuation rows never get the height needed to flow it across a page.
  const heightRuleExact: boolean[] = [];
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const sourceRowCells = tableBlock.rows[rowIdx]?.cells;
    let maxHeight = 0;
    let maxVerticalBorderHeight = 0;
    for (let cellIdx = 0; cellIdx < row.cells.length; cellIdx++) {
      const cell = row.cells[cellIdx];
      const sourceCell = sourceRowCells?.[cellIdx];
      // Stack the cell's blocks exactly as the painter (renderCellContent) does:
      // adjacent paragraphs' after/before spacing collapses to the larger of the
      // two (CSS margin-collapse — the same rule the body paginator uses). The
      // measured paragraph height already bundles before+after, so strip them
      // and re-add a single collapsed gap. Summing them additively here would
      // over-measure the cell (and the row), so the painter's clip/break offsets
      // would no longer line up with the rendered lines (text cut mid-line at a
      // page break).
      let contentHeight = 0;
      let prevAfter = 0;
      for (let blockIdx = 0; blockIdx < cell.blocks.length; blockIdx++) {
        const sourceBlock = sourceCell?.blocks[blockIdx];
        const blockMeasure = cell.blocks[blockIdx];
        if (!sourceBlock || !blockMeasure) continue;
        const visual = measureTableCellBlockVisualHeight(sourceBlock, blockMeasure);
        const spacing = sourceBlock.kind === 'paragraph' ? sourceBlock.attrs?.spacing : undefined;
        const before = spacing?.before ?? 0;
        const after = spacing?.after ?? 0;
        contentHeight += Math.max(prevAfter, before) + (visual - before - after);
        prevAfter = after;
      }
      // The painter renders the last block's trailing space-after as the cell
      // content's paddingBottom (renderCellContent), so include it in the height.
      contentHeight += prevAfter;
      cell.height = contentHeight;
      const padTop = sourceCell?.padding?.top ?? DEFAULT_CELL_PADDING_Y;
      const padBottom = sourceCell?.padding?.bottom ?? DEFAULT_CELL_PADDING_Y;
      cell.height += padTop + padBottom;
      // Only single-row cells set the natural row height. Merged cells are
      // distributed in the second pass.
      if ((cell.rowSpan ?? 1) <= 1) {
        maxHeight = Math.max(maxHeight, cell.height);
      }
      maxVerticalBorderHeight = Math.max(
        maxVerticalBorderHeight,
        getTableCellVerticalBorderHeight(sourceCell)
      );
    }

    const sourceRow = tableBlock.rows[rowIdx];
    const explicitHeight = sourceRow?.height;
    const heightRule = sourceRow?.heightRule;
    heightRuleExact[rowIdx] = heightRule === 'exact' && !!explicitHeight;
    if (explicitHeight && heightRule === 'exact') {
      row.height = explicitHeight;
    } else if (explicitHeight) {
      // ECMA-376 §17.4.81: when hRule is absent or "auto", val is the minimum row height.
      row.height = Math.max(maxHeight + maxVerticalBorderHeight, explicitHeight);
    } else {
      row.height = maxHeight + maxVerticalBorderHeight;
    }
  }

  // Second pass: distribute vertically-merged cells. A cell with rowSpan = N
  // must fit inside the combined height of the N rows it spans. When its
  // content is taller than that combined height, Word grows the region by
  // pushing the surplus into the LAST spanned row (the merged content then
  // flows past the other rows' own content). This keeps rows aligned with
  // Word and lets the layout engine break the tall last row across a page.
  //
  // Deficits are measured against a snapshot of the natural row heights so that
  // overlapping merges (staggered multi-column vMerges sharing rows) each size
  // against the un-grown rows rather than seeing another merge's added surplus.
  const naturalRowHeights = rows.map((r) => r.height);
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const sourceRowCells = tableBlock.rows[rowIdx]?.cells;
    const measuredCells = rows[rowIdx]?.cells;
    if (!sourceRowCells || !measuredCells) continue;
    for (let cellIdx = 0; cellIdx < sourceRowCells.length; cellIdx++) {
      const rowSpan = sourceRowCells[cellIdx]?.rowSpan ?? 1;
      if (rowSpan <= 1) continue;
      const lastRowIdx = Math.min(rowIdx + rowSpan - 1, rows.length - 1);
      const cellNeeded =
        (measuredCells[cellIdx]?.height ?? 0) +
        getTableCellVerticalBorderHeight(sourceRowCells[cellIdx]);
      let spanned = 0;
      for (let r = rowIdx; r <= lastRowIdx; r++) spanned += naturalRowHeights[r] ?? 0;
      const deficit = cellNeeded - spanned;
      if (deficit <= 0) continue;
      // Prefer the last non-`exact` row in the span; exact rows are fixed.
      let target = lastRowIdx;
      while (target > rowIdx && heightRuleExact[target]) target--;
      if (!heightRuleExact[target]) {
        rows[target].height += deficit;
      }
    }
  }

  const totalHeight = rows.reduce((h, r) => h + r.height, 0);
  const totalWidth = columnWidths.reduce((w, cw) => w + cw, 0) || explicitWidthPx || contentWidth;

  return { kind: 'table', rows, columnWidths, totalWidth, totalHeight } as TableExtent;
}
