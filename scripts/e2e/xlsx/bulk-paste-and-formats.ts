import { expect } from 'bun:test';

import { SHEET, displayedText, profiledDisplayList } from './context';
import type { XlsxScenario } from './context';

const TOP = 400;
const ROWS = 50;
const COLS = 10;

export const bulkPasteAndFormats: XlsxScenario = {
  name: 'bulk-paste-and-formats',
  description: 'Pastes 550 cells in one call, then number formats, style patches and a captured format replayed onto another range, verified through the display list and a reopen.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const a1 = (row: number, col: number) => handle.cell(SHEET, row, col).a1;
    const edits = [];
    for (let row = 0; row < ROWS; row += 1) {
      for (let col = 0; col < COLS; col += 1) {
        edits.push({ row: TOP + row, col, input: String(row * COLS + col) });
      }
    }
    for (let row = 0; row < ROWS; row += 1) {
      edits.push({ row: TOP + row, col: COLS, input: `=SUM(${a1(TOP + row, 0)}:${a1(TOP + row, COLS - 1)})` });
    }

    const pasted = recorder.op('editCells:paste550', () => handle.editCells(SHEET, edits));
    expect(pasted.applied).toBe(true);
    expect(handle.cell(SHEET, TOP, COLS - 1).input).toBe('9');
    expect(displayedText(handle, recorder, 'searchText:rowTotal', '45', TOP, COLS)).toContain('45');

    const styled = `${a1(TOP, 0)}:${a1(TOP + 4, 2)}`;
    expect(recorder.op('patchRangeStyle:bold', () => handle.patchRangeStyle(SHEET, styled, { bold: true, fillColor: '#ffee00', horizontalAlignment: 'center' })).applied).toBe(true);
    const formatting = recorder.op('selectionFormatting', () => handle.selectionFormatting(SHEET, styled));
    expect(formatting).toMatchObject({ bold: true, fillColor: '#ffee00', horizontalAlignment: 'center' });

    const percent = `${a1(TOP + 10, 0)}:${a1(TOP + 12, 2)}`;
    expect(recorder.op('setNumberFormat:percent', () => handle.setNumberFormat(SHEET, percent, 'percent')).applied).toBe(true);
    expect(recorder.op('selectionFormatting:percent', () => handle.selectionFormatting(SHEET, percent)).numberFormat).toBe('percent');

    const captured = recorder.op('captureFormat', () => handle.captureFormat(SHEET, a1(TOP, 0)));
    const target = `${a1(TOP + 20, 0)}:${a1(TOP + 22, 2)}`;
    expect(recorder.op('applyFormat', () => handle.applyFormat(SHEET, target, captured)).applied).toBe(true);
    expect(recorder.op('selectionFormatting:applied', () => handle.selectionFormatting(SHEET, target))).toMatchObject({ bold: true, fillColor: '#ffee00' });

    const painted = profiledDisplayList(handle, recorder, 'displayList:formatted', { x: 0, y: 0, width: 1280, height: 800 });
    expect(painted.commands.length).toBeGreaterThan(0);

    const saved = recorder.op('save', () => handle.save());
    const reopened = recorder.op('reopen', () => open(saved));
    expect(recorder.op('selectionFormatting:afterReopen', () => reopened.selectionFormatting(SHEET, styled))).toMatchObject({ bold: true });
    expect(reopened.selectionFormatting(SHEET, percent).numberFormat).toBe('percent');
    expect(reopened.cell(SHEET, TOP + ROWS - 1, COLS).input).toBe(`=SUM(${a1(TOP + ROWS - 1, 0)}:${a1(TOP + ROWS - 1, COLS - 1)})`);
  },
};
