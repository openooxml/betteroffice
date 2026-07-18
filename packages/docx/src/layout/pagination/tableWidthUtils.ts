/**
 * @internal Helpers for resolving DOCX table-width metadata into pixel widths.
 */

import type { TableBlock } from './index';
import { twipsToPixels } from '../../utils/units';

/**
 * Resolve a DOCX width pair to pixels. `pct` values are 50ths of a percent
 * (ECMA-376 §17.18.111 — 5000 means 100%). `dxa` / `auto` / unset are twips.
 *
 * @internal
 */
export function resolveTableWidthPx(
  value: number | undefined,
  widthType: string | undefined,
  parentWidth: number
): number | undefined {
  if (!value || value <= 0) return undefined;
  if (widthType === 'pct') {
    return (parentWidth * value) / 5000;
  }
  if (!widthType || widthType === 'dxa' || widthType === 'auto') {
    return twipsToPixels(value);
  }
  return undefined;
}

/** A cell with its resolved grid position (column index honoring spans). */
export interface ResolvedGridCell {
  rowIndex: number;
  cellIndex: number;
  columnIndex: number;
  colSpan: number;
  rowSpan: number;
}

/**
 * Resolve every cell's grid column index, accounting for `colSpan` and the
 * columns occupied by vertically-merged (`rowSpan`) cells from earlier rows.
 *
 * Single source of truth for table grid geometry — the measurer, the painter,
 * and the row-break paginator all consume this so they agree on which column a
 * cell lives in. Width-free on purpose: callers multiply `columnIndex` by their
 * own (possibly scaled) column widths to get an x offset.
 *
 * @internal
 */
export function resolveCellGrid(tableBlock: TableBlock): ResolvedGridCell[] {
  const occupied = new Map<number, Set<number>>();
  const out: ResolvedGridCell[] = [];
  for (let rowIndex = 0; rowIndex < tableBlock.rows.length; rowIndex++) {
    const cells = tableBlock.rows[rowIndex]?.cells ?? [];
    const occ = occupied.get(rowIndex) ?? new Set<number>();
    let columnIndex = Math.max(
      0,
      Math.min(16384, Math.trunc(tableBlock.rows[rowIndex]?.gridBefore ?? 0))
    );
    while (occ.has(columnIndex)) columnIndex++;
    for (let cellIndex = 0; cellIndex < cells.length; cellIndex++) {
      const cell = cells[cellIndex];
      if (!cell) continue;
      if (cell.gridStart != null) {
        columnIndex = Math.max(columnIndex, Math.min(16384, Math.trunc(cell.gridStart)));
      }
      const colSpan = Math.max(1, Math.min(16384, Math.trunc(cell.colSpan ?? 1)));
      const rowSpan = Math.max(1, Math.min(32768, Math.trunc(cell.rowSpan ?? 1)));
      out.push({ rowIndex, cellIndex, columnIndex, colSpan, rowSpan });
      if (rowSpan > 1) {
        for (let r = rowIndex + 1; r < rowIndex + rowSpan; r++) {
          if (!occupied.has(r)) occupied.set(r, new Set());
          const s = occupied.get(r)!;
          for (let c = 0; c < colSpan; c++) s.add(columnIndex + c);
        }
      }
      columnIndex += colSpan;
      while (occ.has(columnIndex)) columnIndex++;
    }
  }
  return out;
}

/** Total grid columns, derived from the widest row's accumulated colSpans. */
export function countTableColumns(tableBlock: TableBlock): number {
  const resolved = resolveCellGrid(tableBlock);
  let count = 1;
  for (const cell of resolved) count = Math.max(count, cell.columnIndex + cell.colSpan);
  for (let rowIndex = 0; rowIndex < tableBlock.rows.length; rowIndex++) {
    const row = tableBlock.rows[rowIndex];
    const rowCells = resolved.filter((cell) => cell.rowIndex === rowIndex);
    const rowEnd = rowCells.reduce(
      (end, cell) => Math.max(end, cell.columnIndex + cell.colSpan),
      Math.max(0, Math.trunc(row.gridBefore ?? 0))
    );
    count = Math.max(count, rowEnd + Math.max(0, Math.trunc(row.gridAfter ?? 0)));
  }
  return Math.min(16384, count);
}

function preferredWidthPx(
  preferred: { value?: number; type?: 'auto' | 'pct' | 'dxa' | 'nil' } | undefined,
  legacyValue: number | undefined,
  legacyType: string | undefined,
  parentWidth: number,
  legacyPx?: number
): number | undefined {
  return (
    resolveTableWidthPx(preferred?.value, preferred?.type, parentWidth) ??
    resolveTableWidthPx(legacyValue, legacyType, parentWidth) ??
    (legacyPx != null && legacyPx > 0 ? legacyPx : undefined)
  );
}

function addSpanConstraint(widths: number[], start: number, span: number, required: number): void {
  if (!(required > 0) || start >= widths.length) return;
  const end = Math.min(widths.length, start + Math.max(1, span));
  const current = widths.slice(start, end).reduce((sum, width) => sum + width, 0);
  const deficit = required - current;
  if (deficit <= 0) return;
  const share = deficit / Math.max(1, end - start);
  for (let column = start; column < end; column++) widths[column] += share;
}

function distributeToTarget(widths: number[], target: number): number[] {
  const current = widths.reduce((sum, width) => sum + width, 0);
  if (!(target > current) || widths.length === 0) return widths;
  const share = (target - current) / widths.length;
  return widths.map((width) => width + share);
}

function resolveFixedColumnWidths(
  tableBlock: TableBlock,
  contentWidth: number,
  colCount: number,
  explicitWidthPx: number | undefined
): number[] {
  const source = tableBlock.gridWidths ?? tableBlock.columnWidths ?? [];
  let widths = normalizeTableColumnWidths(source, colCount, explicitWidthPx ?? contentWidth);
  const firstRow = resolveCellGrid(tableBlock).filter((cell) => cell.rowIndex === 0);
  for (const gridCell of firstRow) {
    const cell = tableBlock.rows[0]?.cells[gridCell.cellIndex];
    if (!cell) continue;
    const preferred = preferredWidthPx(
      cell.preferredWidth,
      cell.widthValue,
      cell.widthType,
      explicitWidthPx ?? contentWidth,
      cell.width
    );
    if (preferred) addSpanConstraint(widths, gridCell.columnIndex, gridCell.colSpan, preferred);
  }
  if (explicitWidthPx) widths = distributeToTarget(widths, explicitWidthPx);
  return widths;
}

function resolveAutofitColumnWidths(
  tableBlock: TableBlock,
  contentWidth: number,
  colCount: number,
  explicitWidthPx: number | undefined
): number[] {
  const base = normalizeTableColumnWidths(
    tableBlock.gridWidths ?? tableBlock.columnWidths ?? [],
    colCount,
    explicitWidthPx ?? contentWidth
  );
  const minimums = Array(colCount).fill(0) as number[];
  const maximums = Array(colCount).fill(0) as number[];
  for (const gridCell of resolveCellGrid(tableBlock)) {
    const cell = tableBlock.rows[gridCell.rowIndex]?.cells[gridCell.cellIndex];
    if (!cell) continue;
    const preferred = preferredWidthPx(
      cell.preferredWidth,
      cell.widthValue,
      cell.widthType,
      explicitWidthPx ?? contentWidth,
      cell.width
    );
    const min = Math.max(
      0,
      cell.minContentWidth ?? 0,
      cell.noWrap ? (cell.maxContentWidth ?? 0) : 0
    );
    const max = Math.max(min, cell.maxContentWidth ?? preferred ?? 0);
    addSpanConstraint(minimums, gridCell.columnIndex, gridCell.colSpan, min);
    addSpanConstraint(maximums, gridCell.columnIndex, gridCell.colSpan, max);
    if (preferred) addSpanConstraint(maximums, gridCell.columnIndex, gridCell.colSpan, preferred);
  }
  for (let column = 0; column < colCount; column++) {
    if (minimums[column] <= 0)
      minimums[column] = Math.min(base[column], maximums[column] || base[column]);
    maximums[column] = Math.max(minimums[column], maximums[column] || base[column]);
  }

  const minTotal = minimums.reduce((sum, width) => sum + width, 0);
  const maxTotal = maximums.reduce((sum, width) => sum + width, 0);
  const target = Math.max(
    minTotal,
    Math.min(contentWidth, explicitWidthPx ?? (maxTotal > 0 ? maxTotal : contentWidth))
  );
  if (target >= maxTotal) return distributeToTarget(maximums, target);
  const flex = maximums.map((max, index) => Math.max(0, max - minimums[index]));
  const flexTotal = flex.reduce((sum, value) => sum + value, 0);
  const extra = Math.max(0, target - minTotal);
  if (flexTotal <= 0) return distributeToTarget(minimums, target);
  return minimums.map((min, index) => min + (extra * flex[index]) / flexTotal);
}

/**
 * Make `columnWidths` exactly `colCount` long with every entry positive.
 * Missing trailing columns inherit the average of existing positives; zero
 * or negative entries split the leftover `targetWidth` evenly. Callers
 * scale down totals that exceed the target — this helper only fills gaps.
 */
export function normalizeTableColumnWidths(
  columnWidths: number[],
  colCount: number,
  targetWidth: number
): number[] {
  if (colCount <= 0) return [];

  const evenWidth = targetWidth > 0 ? targetWidth / colCount : 0;

  if (columnWidths.length === 0) {
    return Array(colCount).fill(evenWidth);
  }

  let normalized = columnWidths.slice(0, colCount);
  const missingColumns = colCount - normalized.length;
  if (missingColumns > 0) {
    const existingPositive = normalized.filter((width) => width > 0);
    const fallbackWidth =
      existingPositive.length > 0
        ? existingPositive.reduce((sum, width) => sum + width, 0) / existingPositive.length
        : evenWidth;
    normalized = normalized.concat(Array(missingColumns).fill(fallbackWidth));
  }

  const positiveTotal = normalized.reduce((sum, width) => sum + (width > 0 ? width : 0), 0);
  const nonPositiveCount = normalized.filter((width) => width <= 0).length;

  if (positiveTotal <= 0) return Array(colCount).fill(evenWidth);
  if (nonPositiveCount === 0) return normalized;

  const remainingWidth = Math.max(0, targetWidth - positiveTotal);
  const fallbackWidth =
    remainingWidth > 0
      ? remainingWidth / nonPositiveCount
      : positiveTotal / Math.max(1, colCount - nonPositiveCount);

  return normalized.map((width) => (width > 0 ? width : fallbackWidth));
}

/**
 * Resolve a table's per-column pixel widths from its grid metadata and width
 * budget — the width half of `measureTableBlock`, with NO cell-content
 * measurement. Factored out so a caller that only needs widths (e.g. deciding
 * whether a floating table is effectively full-width, before the main measure
 * pass) doesn't have to measure every cell.
 *
 * @internal
 */
export function resolveTableColumnWidths(tableBlock: TableBlock, contentWidth: number): number[] {
  let columnWidths = tableBlock.columnWidths ?? [];
  const explicitWidthPx = preferredWidthPx(
    tableBlock.preferredWidth,
    tableBlock.width,
    tableBlock.widthType,
    contentWidth
  );
  const colCount = countTableColumns(tableBlock);
  const targetWidth = explicitWidthPx ?? contentWidth;

  const algorithm = tableBlock.widthAlgorithm ?? tableBlock.layoutMode ?? 'legacy';
  if (tableBlock.rows.length > 0 && algorithm === 'fixed') {
    return resolveFixedColumnWidths(tableBlock, contentWidth, colCount, explicitWidthPx);
  }
  if (tableBlock.rows.length > 0 && algorithm === 'autofit') {
    return resolveAutofitColumnWidths(tableBlock, contentWidth, colCount, explicitWidthPx);
  }

  if (tableBlock.rows.length > 0) {
    columnWidths = normalizeTableColumnWidths(columnWidths, colCount, targetWidth);
  }

  if (columnWidths.length > 0 && explicitWidthPx) {
    const total = columnWidths.reduce((sum, w) => sum + w, 0);
    if (total > 0 && Math.abs(total - explicitWidthPx) > 1) {
      const scale = explicitWidthPx / total;
      columnWidths = columnWidths.map((w) => w * scale);
    }
  }

  return columnWidths;
}

/**
 * Total pixel width of a table — sum of its resolved column widths, falling
 * back to the explicit table width or the content-width budget. No cell-content
 * measurement, so it is safe to call before the main measure pass. Mirrors the
 * `totalWidth` that `measureTableBlock` produces.
 *
 * @internal
 */
export function resolveTableTotalWidthPx(tableBlock: TableBlock, contentWidth: number): number {
  const columnWidths = resolveTableColumnWidths(tableBlock, contentWidth);
  const explicitWidthPx = preferredWidthPx(
    tableBlock.preferredWidth,
    tableBlock.width,
    tableBlock.widthType,
    contentWidth
  );
  return columnWidths.reduce((w, cw) => w + cw, 0) || explicitWidthPx || contentWidth;
}
