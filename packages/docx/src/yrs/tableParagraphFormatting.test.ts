import { expect, test } from 'bun:test';
import type { Table, TableLook } from '../types/document';
import type { Style } from '../types/styles';
import { tableCellParagraphFormatting } from './tableParagraphFormatting';

test('distinguishes an omitted table look from explicit empty and zero settings', () => {
  const style: Style = {
    type: 'table',
    styleId: 'Grid',
    pPr: { spaceAfter: 10 },
    tblStylePr: [
      { type: 'firstRow', pPr: { spaceAfter: 110 } },
      { type: 'firstCol', pPr: { spaceAfter: 130 } },
      { type: 'band1Horz', pPr: { spaceAfter: 150 } },
      { type: 'band1Vert', pPr: { spaceAfter: 170 } },
      { type: 'band2Vert', pPr: { spaceAfter: 180 } },
    ],
  };
  for (const look of [undefined, {}, { value: '0000' }] as Array<TableLook | undefined>) {
    const table: Table = {
      type: 'table',
      formatting: { look },
      rows: Array.from({ length: 2 }, () => ({
        type: 'tableRow',
        cells: Array.from({ length: 2 }, () => ({ type: 'tableCell', content: [] })),
      })),
    };
    const after = (row: number, column: number) =>
      tableCellParagraphFormatting(table, style, row, column, column + 1, 2)?.spaceAfter;
    expect(after(0, 1)).toBe(look ? 180 : 110);
    expect(after(1, 0)).toBe(look ? 170 : 130);
    expect(after(1, 1)).toBe(look ? 180 : 150);
  }
});
