import { normalizeRange } from '@betteroffice/xlsx';
import type { DisplayList, MergedRange, Selection } from '@betteroffice/xlsx';

export interface AnnouncementStrings {
  cellLabel: string;
  emptyCellLabel: string;
  hasFormula(label: string): string;
  rangeSelected(vars: { start: string; end: string; rows: number; columns: number }): string;
  sheetActivated(vars: { name: string; index: number; count: number }): string;
}

export interface AnnouncedCell {
  /** Painted cell text; null while the painted frame does not cover the cell. */
  text: string | null;
  /** Null while the cell input is still loading. */
  formula: boolean | null;
}

export function columnLetter(col: number): string {
  let out = '';
  for (let n = col; n >= 0; n = Math.floor(n / 26) - 1) out = String.fromCharCode(65 + (n % 26)) + out;
  return out;
}

export function cellAddress(row: number, col: number): string {
  return `${columnLetter(col)}${row + 1}`;
}

function localIndex(start: number, indices: number[] | undefined, count: number, address: number): number {
  const local = indices ? indices.indexOf(address) : address - start;
  return local >= 0 && local < count ? local : -1;
}

const originKey = (x: number, y: number) => `${Math.round(x * 1000)},${Math.round(y * 1000)}`;

/** The text painted in a cell, or null when the frame does not cover it. */
export function paintedCellText(displayList: DisplayList, row: number, col: number): string | null {
  const grid = displayList.grid;
  if (!grid) return null;
  const lr = localIndex(grid.startRow, grid.rowIndices, grid.rowOffsets.length - 1, row);
  const lc = localIndex(grid.startCol, grid.colIndices, grid.colOffsets.length - 1, col);
  if (lr < 0 || lc < 0) return null;
  const origin = originKey(grid.colOffsets[lc], grid.rowOffsets[lr]);
  const parts: string[] = [];
  for (const command of displayList.commands) {
    if (command.op === 'text' && command.clip && !command.ghost && !command.chart &&
      originKey(command.clip.x, command.clip.y) === origin) parts.push(command.text);
  }
  return parts.join(' ').trim();
}

export function selectionKey(sheet: number, selection: Selection): string {
  const range = normalizeRange(selection);
  return `${sheet}:${range.top},${range.left},${range.bottom},${range.right}:${selection.anchor.row},${selection.anchor.col}`;
}

/** One cell, or exactly the merged area the active cell belongs to. */
export function isSingleCell(selection: Selection, merged: readonly MergedRange[] = []): boolean {
  const range = normalizeRange(selection);
  if (range.top === range.bottom && range.left === range.right) return true;
  return merged.some(({ start, end }) => Math.min(start.row, end.row) === range.top &&
    Math.min(start.col, end.col) === range.left && Math.max(start.row, end.row) === range.bottom &&
    Math.max(start.col, end.col) === range.right);
}

const fill = (template: string, vars: Record<string, string>) =>
  template.replace(/\{(\w+)\}/g, (match, key: string) => vars[key] ?? match);

/** What a screen reader hears once the selection settles: the active cell, or the range and its size. */
export function selectionAnnouncement(selection: Selection, cell: AnnouncedCell, strings: AnnouncementStrings,
  sheet?: { name: string; index: number; count: number }, merged: readonly MergedRange[] = []): string {
  const parts: string[] = [];
  if (sheet) parts.push(strings.sheetActivated({ name: sheet.name, index: sheet.index + 1, count: sheet.count }));
  const range = normalizeRange(selection);
  if (isSingleCell(selection, merged)) {
    const address = cellAddress(selection.anchor.row, selection.anchor.col);
    const text = cell.text ?? '';
    const label = text ? fill(strings.cellLabel, { address, value: text }) : fill(strings.emptyCellLabel, { address });
    parts.push(cell.formula ? strings.hasFormula(label) : label);
  } else {
    parts.push(strings.rangeSelected({ start: cellAddress(range.top, range.left), end: cellAddress(range.bottom, range.right),
      rows: range.bottom - range.top + 1, columns: range.right - range.left + 1 }));
  }
  return parts.join(', ');
}
