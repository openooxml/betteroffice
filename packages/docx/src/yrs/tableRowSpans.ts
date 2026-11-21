import type { Table } from '../types/document';

/** @internal */
export type RowSpanInfo = { rowSpan: number; skip: boolean };

/** @internal Vertical-merge spans keyed by `row-column`; continuation cells under an open span are skipped when seeding. */
export function calculateRowSpans(table: Table): Map<string, RowSpanInfo> {
  const result = new Map<string, RowSpanInfo>();
  const active = new Map<number, number>();
  table.rows.forEach((row, rowIndex) => {
    let column = row.formatting?.gridBefore ?? 0;
    const cells = row.cells.map((cell) => {
      const current = column;
      column += cell.formatting?.gridSpan ?? 1;
      return { column: current, vMerge: cell.formatting?.vMerge, key: `${rowIndex}-${current}` };
    });
    const empty =
      cells.length > 0 &&
      cells.every((cell) => cell.vMerge === 'continue' && active.has(cell.column));
    if (empty) {
      for (const cell of cells) {
        active.delete(cell.column);
        result.set(cell.key, { rowSpan: 1, skip: false });
      }
      return;
    }
    for (const cell of cells) {
      if (cell.vMerge === 'restart') {
        active.set(cell.column, rowIndex);
        result.set(cell.key, { rowSpan: 1, skip: false });
      } else if (cell.vMerge === 'continue') {
        const start = active.get(cell.column);
        if (start === undefined) result.set(cell.key, { rowSpan: 1, skip: false });
        else {
          const owner = result.get(`${start}-${cell.column}`);
          if (owner) owner.rowSpan += 1;
          result.set(cell.key, { rowSpan: 1, skip: true });
        }
      } else {
        active.delete(cell.column);
        result.set(cell.key, { rowSpan: 1, skip: false });
      }
    }
  });
  return result;
}

/** @internal For each row of `table`, the source index of every cell that seeding keeps as a cell story. */
export function seededCellIndices(table: Table): number[][] {
  const spans = calculateRowSpans(table);
  return table.rows.map((row, rowIndex) => {
    let column = row.formatting?.gridBefore ?? 0;
    const kept: number[] = [];
    row.cells.forEach((cell, index) => {
      const start = column;
      column += cell.formatting?.gridSpan ?? 1;
      if (!spans.get(`${rowIndex}-${start}`)?.skip) kept.push(index);
    });
    return kept;
  });
}
