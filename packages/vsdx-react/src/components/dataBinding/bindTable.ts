import { isShapeDataValueEditable, shapeDataRows, shapeDataValueFormula } from '@betteroffice/vsdx';
import type { DiagramHandle, ShapeDataReceipt, ShapeSnapshot } from '@betteroffice/vsdx';

/** A table imported from a spreadsheet: one header per column, then the data rows. */
export interface ImportedTable {
  headers: readonly string[];
  rows: readonly (readonly string[])[];
}

/** One column of an imported row mapped onto the shape-data row it writes. */
export interface ColumnBinding {
  column: number;
  rowName: string;
}

export interface BindOutcome {
  receipts: readonly ShapeDataReceipt[];
  refusals: readonly ShapeDataReceipt[];
  applied: boolean;
}

/**
 * Pairs each column with the shape-data row whose label matches its header.
 *
 * Rows the engine would refuse are left out, so a bind is never offered for one.
 */
export function matchColumnsToRows(table: ImportedTable, shape: ShapeSnapshot | null): ColumnBinding[] {
  const rows = shapeDataRows(shape);
  // null marks a name two rows answer to: binding it would silently pick one of them.
  const byLabel = new Map<string, string | null>();
  const claim = (key: string, rowName: string) => {
    const held = byLabel.get(key);
    if (held === undefined) byLabel.set(key, rowName);
    else if (held !== rowName) byLabel.set(key, null);
  };
  for (const row of rows) {
    if (row.rowName === null || !isShapeDataValueEditable(row)) continue;
    claim(normalize(row.label), row.rowName);
    claim(normalize(row.rowName), row.rowName);
  }
  const bindings: ColumnBinding[] = [];
  const used = new Set<string>();
  table.headers.forEach((header, column) => {
    const rowName = byLabel.get(normalize(header));
    if (rowName === undefined || rowName === null || used.has(rowName)) return;
    used.add(rowName);
    bindings.push({ column, rowName });
  });
  return bindings;
}

function normalize(text: string): string {
  return text.trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Writes one imported row into the shape's data as a single undo entry. */
export function bindRowToShape(
  handle: DiagramHandle,
  pageId: string,
  shapeId: string,
  shape: ShapeSnapshot,
  table: ImportedTable,
  rowIndex: number,
  bindings: readonly ColumnBinding[],
): BindOutcome {
  const source = table.rows[rowIndex];
  if (!source) throw new Error(`Imported table has no row ${rowIndex}.`);
  const types = new Map(shapeDataRows(shape).map((row) => [row.rowName, row.type] as const));
  const writes = bindings.flatMap(({ column, rowName }) => {
    const cell = source[column];
    if (cell === undefined || cell.trim() === '') return [];
    return [{ rowName, formula: shapeDataValueFormula(types.get(rowName) ?? 'string', cell) }];
  });
  if (writes.length === 0) return { receipts: [], refusals: [], applied: false };
  const receipts = handle.setShapeData(pageId, shapeId, writes);
  const refusals = receipts.filter((receipt) => receipt.refusal !== null);
  return { receipts, refusals, applied: refusals.length === 0 };
}
