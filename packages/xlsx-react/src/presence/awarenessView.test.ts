import { describe, expect, test } from 'bun:test';
import * as collaboration from '@betteroffice/xlsx/collaboration';
import type { AwarenessCursor } from '@betteroffice/xlsx/collaboration';
import * as view from './awarenessView';

describe('presence awareness helpers', () => {
  test('match the collaboration exports', () => {
    expect(view.AWARENESS_LABEL_DURATION_MS).toBe(collaboration.AWARENESS_LABEL_DURATION_MS);
    const colors = ['#abc', '#A1B2C3', '#a1b2c3', '#12345', 'abc', '#ggg', '', null, undefined, 7, {}];
    for (let clientId = 0; clientId < 3 * collaboration.AWARENESS_COLORS.length; clientId += 1) {
      for (const color of colors) {
        expect(view.normalizeAwarenessColor(color, clientId))
          .toBe(collaboration.normalizeAwarenessColor(color, clientId));
      }
    }
    const sheetIds = ['s1', 's2'];
    const cell = (row: number, col: number) => ({ row, col });
    const cursors: AwarenessCursor[] = [
      { sheet: 's1', anchor: cell(0, 0), head: cell(0, 0) },
      { sheet: 's2', anchor: cell(9, 4), head: cell(2, 7) },
      { sheet: 's3', anchor: cell(0, 0), head: cell(1, 1) },
      { sheet: 's1', anchor: cell(-1, 0), head: cell(1, 1) },
      { sheet: 's1', anchor: cell(0, 0), head: cell(collaboration.XLSX_MAX_ROWS, 0) },
      { sheet: 's1', anchor: cell(0, collaboration.XLSX_MAX_COLUMNS - 1), head: cell(collaboration.XLSX_MAX_ROWS - 1, 0) },
      { sheet: 's1', anchor: cell(0, collaboration.XLSX_MAX_COLUMNS), head: cell(0, 0) },
      { sheet: 's1', anchor: cell(0.5, 0), head: cell(0, 0) },
      { sheet: 's1', anchor: cell(Number.NaN, 0), head: cell(0, 0) },
    ];
    for (const cursor of cursors) {
      expect(view.resolveAwarenessCursor(cursor, sheetIds))
        .toEqual(collaboration.resolveAwarenessCursor(cursor, sheetIds));
      for (const [rows, columns] of [[10, 5], [0, 5], [10, 0], [3, 8]]) {
        expect(view.resolveAwarenessCursor(cursor, sheetIds, rows, columns))
          .toEqual(collaboration.resolveAwarenessCursor(cursor, sheetIds, rows, columns));
      }
    }
  });
});
