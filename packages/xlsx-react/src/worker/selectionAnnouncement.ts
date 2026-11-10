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

function lastLocalAtOrBefore(start: number, indices: number[] | undefined, count: number, address: number): number {
  if (!indices) return Math.min(address - start, count - 1);
  let out = -1;
  for (let index = 0; index < count && index < indices.length; index++) if (indices[index] <= address) out = index;
  return out;
}

/** A point inside the cell a text run belongs to. */
function runAnchor(x: number, clip: { x: number; w: number }, align: 'left' | 'center' | 'right' | undefined): number {
  const inset = Math.min(0.5, clip.w / 2);
  if (align === 'left') return clip.x + inset;
  if (align === 'right') return clip.x + clip.w - inset;
  return Math.min(Math.max(x, clip.x + inset), clip.x + clip.w - inset);
}

/**
 * The text painted in a cell, or null when the frame does not cover it. Text
 * spilling over blank neighbours clips wider than its cell, toward the side its
 * alignment opens: a left-aligned run's clip starts in its own cell and a
 * right-aligned run's clip ends in it.
 */
export function paintedCellText(displayList: DisplayList, row: number, col: number,
  merged: readonly MergedRange[] = []): string | null {
  const grid = displayList.grid;
  if (!grid) return null;
  const columns = grid.colOffsets.length - 1;
  const lr = localIndex(grid.startRow, grid.rowIndices, grid.rowOffsets.length - 1, row);
  const lc = localIndex(grid.startCol, grid.colIndices, columns, col);
  if (lr < 0 || lc < 0) return null;
  const merge = merged.find(({ start, end }) =>
    Math.min(start.row, end.row) === row && Math.min(start.col, end.col) === col);
  const lastCol = merge ? Math.max(merge.start.col, merge.end.col) : col;
  const le = Math.max(lc, lastCol === col ? lc : lastLocalAtOrBefore(grid.startCol, grid.colIndices, columns, lastCol));
  const left = grid.colOffsets[lc];
  const right = grid.colOffsets[le + 1];
  const top = Math.round(grid.rowOffsets[lr] * 1000);
  const parts: string[] = [];
  for (const command of displayList.commands) {
    if (command.op !== 'text' || !command.clip || command.ghost || command.chart) continue;
    if (Math.round(command.clip.y * 1000) !== top) continue;
    const anchor = runAnchor(command.x, command.clip, command.align);
    if (anchor >= left && anchor < right) parts.push(command.text);
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
