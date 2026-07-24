/**
 * Pure a11y grid builder: turns a display list plus the current selection into a
 * serializable {@link A11yGrid} the chrome renders offscreen for screen readers.
 * Cell text is recovered from the frame's clipped text commands, keyed by cell
 * box origin; geometry and addresses come from the frame's grid metadata. No
 * DOM, no i18n import — label templates arrive as a param.
 */

import type { DisplayList, GridMeta } from '../display-list/types';
import { normalizeRange } from '../selection/index';
import type { CellRange, Selection } from '../selection/index';
import type { A11yCell, A11yColumnHeader, A11yGrid, A11yRow, A11yStrings } from './types';

export type { A11yCell, A11yColumnHeader, A11yGrid, A11yRow, A11yStrings } from './types';

// bijective base-26 column letter: 0 -> A, 25 -> Z, 26 -> AA.
function columnLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function cellAddress(row: number, col: number): string {
  return `${columnLetter(col)}${row + 1}`;
}

function fill(template: string, vars: Record<string, string | number>): string {
  return template.replace(/\{(\w+)\}/g, (_, key) => String(vars[key] ?? ''));
}

// round to the golden precision so float noise in clip origins keys stably.
function key(x: number, y: number): string {
  return `${Math.round(x * 1000)},${Math.round(y * 1000)}`;
}

// index each clipped text command by its clip's top-left, the cell box origin
// the renderer aligns cell text to. later commands win, matching paint order.
// ghost previews are skipped so proposed values never read as committed.
function textByCellOrigin(displayList: DisplayList): Map<string, string> {
  const map = new Map<string, string>();
  for (const cmd of displayList.commands) {
    if (cmd.op === 'text' && cmd.clip && !cmd.ghost) map.set(key(cmd.clip.x, cmd.clip.y), cmd.text);
  }
  return map;
}

function isSelected(range: CellRange | null, row: number, col: number): boolean {
  if (!range) return false;
  return row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
}

function trackAddress(start: number, indices: number[] | undefined, local: number): number {
  return indices?.[local] ?? start + local;
}

function buildCell(
  grid: GridMeta,
  lr: number,
  lc: number,
  text: string,
  selected: boolean,
  t: A11yStrings
): A11yCell {
  const row = trackAddress(grid.startRow, grid.rowIndices, lr);
  const col = trackAddress(grid.startCol, grid.colIndices, lc);
  const address = cellAddress(row, col);
  const empty = text.length === 0;
  const template = empty
    ? selected
      ? t.emptyCellLabelSelected
      : t.emptyCellLabel
    : selected
      ? t.cellLabelSelected
      : t.cellLabel;
  return {
    row,
    col,
    address,
    text,
    selected,
    label: fill(template, { address, value: text }),
  };
}

/**
 * Build the offscreen grid mirror for a frame. Returns an empty grid (no rows)
 * when the frame carries no grid metadata, so the chrome can render nothing
 * rather than guess geometry.
 */
export function buildA11yGrid(
  displayList: DisplayList,
  selection: Selection | null,
  sheetName: string,
  t: A11yStrings
): A11yGrid {
  const grid = displayList.grid;
  const label = fill(t.gridLabel, { sheet: sheetName });
  if (!grid) return { label, sheetName, columnHeaders: [], rows: [] };

  const cols = grid.colOffsets.length - 1;
  const rows = grid.rowOffsets.length - 1;
  const range = selection ? normalizeRange(selection) : null;
  const textAt = textByCellOrigin(displayList);

  const columnHeaders: A11yColumnHeader[] = [];
  for (let lc = 0; lc < cols; lc++) {
    const col = trackAddress(grid.startCol, grid.colIndices, lc);
    const letter = columnLetter(col);
    columnHeaders.push({ col, text: letter, label: fill(t.columnHeaderLabel, { column: letter }) });
  }

  const outRows: A11yRow[] = [];
  for (let lr = 0; lr < rows; lr++) {
    const row = trackAddress(grid.startRow, grid.rowIndices, lr);
    const cells: A11yCell[] = [];
    for (let lc = 0; lc < cols; lc++) {
      const origin = key(grid.colOffsets[lc], grid.rowOffsets[lr]);
      const text = textAt.get(origin) ?? '';
      const col = trackAddress(grid.startCol, grid.colIndices, lc);
      cells.push(buildCell(grid, lr, lc, text, isSelected(range, row, col), t));
    }
    outRows.push({ row, header: fill(t.rowHeaderLabel, { row: row + 1 }), cells });
  }

  return { label, sheetName, columnHeaders, rows: outRows };
}
