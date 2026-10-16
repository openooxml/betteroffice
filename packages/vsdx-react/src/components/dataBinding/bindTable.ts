import { formatOptions, isShapeDataValueEditable, shapeDataRows, shapeDataValueFormula, visibleShapeDataRows } from '@betteroffice/vsdx';
import type { DiagramHandle, ShapeDataReceipt, ShapeDataWrite, ShapeSnapshot } from '@betteroffice/vsdx';
import { findShapePlacement } from '../ribbon/commands';

/** A table imported from a spreadsheet: one header per column, then the data rows. */
export interface ImportedTable {
  headers: readonly string[];
  rows: readonly (readonly string[])[];
}

/** One column of an imported row mapped onto the shape-data row it writes. */
export interface ColumnBinding {
  column: number;
  rowName: string;
  sectionIndex?: number;
}

export interface BindOutcome {
  receipts: readonly ShapeDataReceipt[];
  refusals: readonly ShapeDataReceipt[];
  applied: boolean;
}

/** Matches headers to visible, named rows whose values pass isShapeDataValueEditable. */
export function matchColumnsToRows(table: ImportedTable, shape: ShapeSnapshot | null): ColumnBinding[] {
  const rows = visibleShapeDataRows(shape);
  const byLabel = new Map<string, Pick<ColumnBinding, 'rowName' | 'sectionIndex'> | null>();
  const claim = (key: string, target: Pick<ColumnBinding, 'rowName' | 'sectionIndex'>) => {
    const held = byLabel.get(key);
    if (held === undefined) byLabel.set(key, target);
    else if (held !== null && rowKey(held.rowName, held.sectionIndex) !== rowKey(target.rowName, target.sectionIndex)) byLabel.set(key, null);
  };
  for (const row of rows) {
    if (row.rowName === null || !isShapeDataValueEditable(row)) continue;
    const target = { rowName: row.rowName, sectionIndex: row.sectionIndex };
    claim(normalize(row.label), target);
    claim(normalize(row.rowName), target);
  }
  const bindings: ColumnBinding[] = [];
  const used = new Set<string>();
  table.headers.forEach((header, column) => {
    const target = byLabel.get(normalize(header));
    if (!target) return;
    const key = rowKey(target.rowName, target.sectionIndex);
    if (used.has(key)) return;
    used.add(key);
    bindings.push({ column, ...target });
  });
  return bindings;
}

function normalize(text: string): string {
  return text.trim().toLowerCase().replace(/\s+/g, ' ');
}

function rowKey(rowName: string | null, sectionIndex?: number): string {
  return `${sectionIndex ?? 0}\u001f${rowName}`;
}

/** Writes one imported row into the shape's data as a single undo entry. */
export function bindRowToShape(
  handle: DiagramHandle,
  pageId: string,
  shapeId: string,
  table: ImportedTable,
  rowIndex: number,
  bindings: readonly ColumnBinding[],
): BindOutcome {
  const source = table.rows[rowIndex];
  if (!source) throw new Error(`Imported table has no row ${rowIndex}.`);
  const page = handle.snapshot().pages.find((entry) => entry.id === pageId);
  const shape = page ? findShapePlacement(page.shapes, shapeId)?.shape : undefined;
  if (!shape) throw new Error(`vsdx shape ${shapeId} is no longer part of page ${pageId}`);
  const rows = new Map(shapeDataRows(shape).map((row) => [rowKey(row.rowName, row.sectionIndex), row] as const));
  const writes: ShapeDataWrite[] = [];
  const checked: ShapeDataReceipt[] = [];
  for (const { column, rowName, sectionIndex } of bindings) {
    const cell = source[column];
    if (cell === undefined || cell.trim() === '') continue;
    const row = rows.get(rowKey(rowName, sectionIndex));
    let refusal: string | null = null;
    let input = cell;
    if (!row) refusal = 'The shape-data row no longer exists.';
    else if (row.type === 'fixed-list' && !formatOptions(row.format).includes(cell)) refusal = 'The value is not in the fixed list.';
    else if (row.type === 'boolean') {
      const value = cell.trim().toLowerCase();
      if (['true', '1', 'yes'].includes(value)) input = '1';
      else if (['false', '0', 'no'].includes(value)) input = '0';
      else refusal = 'The value is not a recognized boolean.';
    }
    checked.push({ pageId, shapeId, rowName, rowIndex: null, sectionIndex: sectionIndex ?? null, before: row?.formula ?? null, after: null, refusal });
    if (row && refusal === null) writes.push({ rowName, sectionIndex, formula: shapeDataValueFormula(row.type, input) });
  }
  const invalid = checked.filter((receipt) => receipt.refusal !== null);
  if (invalid.length > 0) return { receipts: checked, refusals: invalid, applied: false };
  if (writes.length === 0) return { receipts: [], refusals: [], applied: false };
  const receipts = handle.setShapeData(pageId, shapeId, writes);
  const refusals = receipts.filter((receipt) => receipt.refusal !== null);
  return { receipts, refusals, applied: refusals.length === 0 };
}
