import { describe, expect, it } from 'bun:test';
import type { DisplayList, Selection } from '@betteroffice/xlsx';
import { createT, en } from '@betteroffice/xlsx-i18n';
import { isSingleCell, paintedCellText, selectionAnnouncement, type AnnouncementStrings } from './selectionAnnouncement';

const t = createT(en);
const strings: AnnouncementStrings = {
  cellLabel: t('a11y.cellLabel'),
  emptyCellLabel: t('a11y.emptyCellLabel'),
  hasFormula: (label) => t('a11y.hasFormula', { label }),
  rangeSelected: (vars) => t('a11y.rangeSelected', vars),
  sheetActivated: (vars) => t('a11y.sheetActivated', vars),
};
const at = (row: number, col: number, focus = { row, col }): Selection => ({ anchor: { row, col }, focus });

const frame: DisplayList = {
  width: 300, height: 60,
  grid: { startRow: 4, startCol: 1, rowOffsets: [0, 20, 40], colOffsets: [0, 64, 128] },
  commands: [
    { op: 'text', x: 2, y: 14, text: '42', fontSize: 11, color: '#000000', clip: { x: 0, y: 0, w: 64, h: 20 } },
    { op: 'text', x: 66, y: 14, text: 'Total', fontSize: 11, color: '#000000', clip: { x: 64, y: 0, w: 64, h: 20 } },
    { op: 'text', x: 66, y: 34, text: 'later', fontSize: 11, color: '#000000', clip: { x: 64, y: 20, w: 64, h: 20 }, ghost: true },
  ],
} as DisplayList;

describe('selection announcement', () => {
  it('reads the painted text of a cell in the frame', () => {
    expect(paintedCellText(frame, 4, 1)).toBe('42');
    expect(paintedCellText(frame, 4, 2)).toBe('Total');
    expect(paintedCellText(frame, 5, 2)).toBe('');
    expect(paintedCellText(frame, 9, 1)).toBeNull();
  });

  it('announces the active cell with its value and a formula hint', () => {
    expect(selectionAnnouncement(at(4, 1), { text: '42', formula: false }, strings)).toBe('B5, 42');
    expect(selectionAnnouncement(at(4, 1), { text: '42', formula: true }, strings)).toBe('B5, 42, has formula');
    expect(selectionAnnouncement(at(5, 2), { text: '', formula: false }, strings)).toBe('C6, empty');
  });

  it('announces a range with its extent', () => {
    expect(selectionAnnouncement(at(0, 0, { row: 4, col: 2 }), { text: null, formula: null }, strings))
      .toBe('A1 to C5 selected, 5 rows by 3 columns');
    expect(selectionAnnouncement(at(0, 1, { row: 0, col: 0 }), { text: null, formula: null }, strings))
      .toBe('A1 to B1 selected, 1 row by 2 columns');
  });

  it('reads a merged area as its active cell', () => {
    const merged = [{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }];
    const selection: Selection = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 }, range: { top: 0, left: 0, bottom: 1, right: 1 } };
    expect(isSingleCell(selection, merged)).toBe(true);
    expect(selectionAnnouncement(selection, { text: 'Head', formula: false }, strings, undefined, merged)).toBe('A1, Head');
  });

  it('names the sheet first after a sheet switch', () => {
    expect(selectionAnnouncement(at(0, 0), { text: '', formula: false }, strings, { name: 'Q2', index: 1, count: 3 }))
      .toBe('Q2, sheet 2 of 3, A1, empty');
  });

  it('keeps cell text with braces verbatim', () => {
    expect(selectionAnnouncement(at(0, 0), { text: '{address}', formula: false }, strings)).toBe('A1, {address}');
  });
});
