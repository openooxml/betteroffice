import type { CellSnapshot, ShapeSnapshot } from './types';

export type ShapeDataType = 'string' | 'fixed-list' | 'number' | 'boolean' | 'variable-list' | 'date' | 'duration' | 'currency';

export interface ShapeDataRow {
  rowName: string | null;
  rowIndex: number | null;
  sectionIndex?: number;
  label: string;
  prompt: string | null;
  type: ShapeDataType;
  format: string | null;
  formula: string | null;
  value: string | null;
  displayValue: string;
  sortKey: string | null;
  invisible: boolean;
  ask: boolean;
}

export function shapeDataTypeFromValue(value: string | null | undefined): ShapeDataType {
  const normalized = (value ?? '').trim().replace(/^=/, '').trim();
  switch (normalized) {
    case '1': return 'fixed-list';
    case '2': return 'number';
    case '3': return 'boolean';
    case '4': return 'variable-list';
    case '5': return 'date';
    case '6': return 'duration';
    case '7': return 'currency';
    default: return 'string';
  }
}

function rowKey(cell: CellSnapshot): string | null {
  if (cell.locator.section !== 'Property' || !cell.locator.row) return null;
  const section = cell.locator.sectionIndex ?? 0;
  const row = 'name' in cell.locator.row ? `N:${cell.locator.row.name}` : `IX:${cell.locator.row.index}`;
  return `${section}\u{1f}${row}`;
}

function byName(cells: CellSnapshot[], name: string): CellSnapshot | undefined {
  return cells.find((cell) => cell.name === name);
}

function textOf(cell: CellSnapshot | undefined): string | null {
  if (!cell) return null;
  if (cell.value !== null && cell.value !== '') return cell.value;
  return unquoteFormula(cell.formula);
}

function truthyOf(cell: CellSnapshot | undefined): boolean {
  if (!cell) return false;
  const source = cell.value ?? cell.formula;
  if (source === null || source === undefined) return false;
  const normalized = source.trim().replace(/^=/, '').trim();
  if (/^true$/i.test(normalized)) return true;
  if (/^false$/i.test(normalized)) return false;
  const number = Number(normalized);
  if (normalized !== '' && Number.isFinite(number)) return number !== 0;
  return false;
}

export function unquoteFormula(formula: string | null | undefined): string | null {
  if (!formula) return null;
  const literal = formula.trim().replace(/^=/, '').trim();
  if (literal.length >= 2 && literal.startsWith('"') && literal.endsWith('"')) {
    return literal.slice(1, -1).replace(/""/g, '"');
  }
  return null;
}

export function quoteShapeDataValue(text: string): string {
  return `"${text.replace(/"/g, '""')}"`;
}

export function formatOptions(format: string | null | undefined): string[] {
  if (!format) return [];
  return format.split(';').map((option) => option.trim()).filter((option) => option !== '');
}

export function shapeDataValueFormula(type: ShapeDataType, input: string): string {
  const trimmed = input.trim();
  if (type === 'boolean') return trimmed !== '' && trimmed !== '0' && !/^false$/i.test(trimmed) ? '1' : '0';
  if (type === 'number') return trimmed;
  return quoteShapeDataValue(input);
}

export function shapeDataRows(shape: ShapeSnapshot | null | undefined): ShapeDataRow[] {
  if (!shape) return [];
  const groups = new Map<string, CellSnapshot[]>();
  for (const cell of shape.cells) {
    const key = rowKey(cell);
    if (key === null) continue;
    const group = groups.get(key);
    if (group) group.push(cell);
    else groups.set(key, [cell]);
  }
  const rows: ShapeDataRow[] = [];
  for (const cells of groups.values()) {
    const first = cells[0];
    const row = first.locator.row;
    const rowName = row && 'name' in row ? row.name : null;
    const rowIndex = row && 'index' in row ? row.index : null;
    const fallback = rowName ?? (rowIndex !== null ? `Row ${rowIndex}` : 'Row');
    const value = byName(cells, 'Value');
    const typeCell = byName(cells, 'Type');
    rows.push({
      rowName,
      rowIndex,
      sectionIndex: first.locator.sectionIndex ?? undefined,
      label: textOf(byName(cells, 'Label')) ?? fallback,
      prompt: textOf(byName(cells, 'Prompt')),
      type: shapeDataTypeFromValue(typeCell?.value ?? typeCell?.formula),
      format: textOf(byName(cells, 'Format')),
      formula: value?.formula ?? null,
      value: value?.value ?? null,
      displayValue: value?.value ?? unquoteFormula(value?.formula) ?? '',
      sortKey: textOf(byName(cells, 'SortKey')),
      invisible: truthyOf(byName(cells, 'Invisible')),
      ask: truthyOf(byName(cells, 'Ask')) || truthyOf(byName(cells, 'Verify')),
    });
  }
  rows.sort((left, right) => (left.sortKey ?? '').localeCompare(right.sortKey ?? ''));
  return rows;
}

export function visibleShapeDataRows(shape: ShapeSnapshot | null | undefined): ShapeDataRow[] {
  return shapeDataRows(shape).filter((row) => !row.invisible);
}
