/**
 * Golden corpus for the pure editing seams (selection → a11y grid, hit-testing).
 *
 * Same contract as the display-list corpus: each scenario builds a serializable
 * snapshot that the committed `golden/<name>.json` freezes verbatim. The goldens
 * pin CURRENT output so refactors and the eventual Rust port prove byte-
 * identical results — they judge stability, not correctness. Reuses the
 * display-list canonicalizer so number rounding and key ordering match.
 */

import type { DisplayList, DrawCmd, GridMeta, Rect } from '../display-list/types';
import { buildA11yGrid } from '../a11y/index';
import type { A11yStrings } from '../a11y/index';
import { cellAtPoint, cellRect, rangeRect } from '../hittest/index';
import type { Selection } from '../selection/index';

/**
 * One pinned scenario; `build()` returns any serializable snapshot.
 */
export interface GoldenScenario {
  name: string;
  pins: string;
  build(): unknown;
}

// the a11y label templates, mirroring the en.json a11y keys the chrome passes.
const A11Y_STRINGS: A11yStrings = {
  gridLabel: 'Spreadsheet grid',
  cellLabel: '{address}, {value}',
  cellLabelSelected: '{address}, {value}, selected',
  emptyCellLabel: '{address}, empty',
  emptyCellLabelSelected: '{address}, empty, selected',
  rowHeaderLabel: 'Row {row}',
  columnHeaderLabel: 'Column {column}',
};

const COL_W = 80;
const ROW_H = 20;

function cellBox(col: number, row: number): Rect {
  return { x: col * COL_W, y: row * ROW_H, w: COL_W, h: ROW_H };
}

function labelCmd(col: number, row: number, value: string): DrawCmd {
  const clip = cellBox(col, row);
  return {
    op: 'text',
    x: clip.x + 2,
    y: clip.y + ROW_H - 4,
    text: value,
    fontSize: 12,
    color: '#000000',
    clip,
  };
}

// a display list for a `cols` x `rows` window with per-cell text and grid meta.
function gridDisplayList(
  cols: number,
  rows: number,
  startRow: number,
  startCol: number
): DisplayList {
  const commands: DrawCmd[] = [];
  const colOffsets: number[] = [];
  const rowOffsets: number[] = [];
  for (let c = 0; c <= cols; c++) colOffsets.push(c * COL_W);
  for (let r = 0; r <= rows; r++) rowOffsets.push(r * ROW_H);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      commands.push(labelCmd(c, r, `r${startRow + r}c${startCol + c}`));
    }
  }
  const grid: GridMeta = { startRow, startCol, rowOffsets, colOffsets };
  return { width: cols * COL_W, height: rows * ROW_H, commands, grid };
}

function a11yGridWithSelection(): GoldenScenario {
  return {
    name: 'a11y-grid-with-selection',
    pins: 'a selection over a labelled grid mirrors into an offscreen a11y grid',
    build() {
      const displayList = gridDisplayList(3, 2, 0, 0);
      const selection: Selection = { anchor: { row: 0, col: 1 }, focus: { row: 1, col: 2 } };
      return buildA11yGrid(displayList, selection, 'Sheet1', A11Y_STRINGS);
    },
  };
}

function a11yGridScrolled(): GoldenScenario {
  return {
    name: 'a11y-grid-scrolled',
    pins: 'a scrolled window (nonzero start row/col) addresses cells past Z with no selection',
    build() {
      const displayList = gridDisplayList(2, 2, 12, 26);
      return buildA11yGrid(displayList, null, 'Data', A11Y_STRINGS);
    },
  };
}

function hittestTable(): GoldenScenario {
  return {
    name: 'hittest-table',
    pins: 'point→cell, cell→rect, and clipped range→rect over a synthetic grid window',
    build() {
      const grid: GridMeta = {
        startRow: 10,
        startCol: 5,
        colOffsets: [0, 50, 150, 180],
        rowOffsets: [0, 10, 50],
      };
      const points = [
        { x: 0, y: 0 },
        { x: 49, y: 5 },
        { x: 50, y: 5 },
        { x: 179, y: 49 },
        { x: 180, y: 10 },
        { x: -1, y: 0 },
      ];
      // echo the window origin as scalars (not the offset arrays) so the
      // snapshot stays formatter-clean; the offsets live in this source.
      return {
        window: { startRow: grid.startRow, startCol: grid.startCol },
        hits: points.map((p) => ({ ...p, cell: cellAtPoint(grid, p.x, p.y) })),
        rects: [
          { row: 10, col: 5, rect: cellRect(grid, 10, 5) },
          { row: 11, col: 7, rect: cellRect(grid, 11, 7) },
          { row: 9, col: 5, rect: cellRect(grid, 9, 5) },
        ],
        ranges: [
          {
            range: { top: 10, left: 5, bottom: 11, right: 6 },
            rect: rangeRect(grid, { top: 10, left: 5, bottom: 11, right: 6 }),
          },
          {
            range: { top: 11, left: 6, bottom: 99, right: 99 },
            rect: rangeRect(grid, { top: 11, left: 6, bottom: 99, right: 99 }),
          },
        ],
      };
    },
  };
}

/**
 * The full editing-seam corpus. Order is stable for deterministic runs.
 */
export const corpus: GoldenScenario[] = [
  a11yGridWithSelection(),
  a11yGridScrolled(),
  hittestTable(),
];
