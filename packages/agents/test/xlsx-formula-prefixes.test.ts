import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openXlsx } from '../src/xlsx';
import { implicitFontWorkbook } from './format-fixtures';

async function excelSaved(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await implicitFontWorkbook());
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('</sheetData>', '<row r="6"><c r="A6"><f>_xlfn.LET(_xlpm.x,2,_xlpm.x*3)</f><v>6</v></c><c r="C6"><f>_xlfn.SEQUENCE(2,2)</f><v>1</v></c></row></sheetData>'));
  return zip.generateAsync({ type: 'uint8array' });
}

test('agent formula reads leave out storage prefixes and still guard their writes', async () => {
  const book = await openXlsx(await excelSaved());
  try {
    const sheet = book.list().items[0].sheetId;
    const cells = book.readCells({ sheet, range: 'A6:C6' }).cells;
    expect(cells.map(cell => cell.formula)).toEqual(['LET(x,2,x*3)', null, 'SEQUENCE(2,2)']);
    expect(book.read(cells[2].ref, { field: 'formula' }).text).toBe('SEQUENCE(2,2)');
    const pending = book.proposeCells({ author: 'test', edits: [{ cell: cells[0].cell, input: '=LET(x,2,x*4)' }] });
    const saved = await JSZip.loadAsync(await book.export(pending.id));
    expect(await saved.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('<f>_xlfn.LET(_xlpm.x,2,_xlpm.x*4)</f>');
    const reopened = await openXlsx(await book.export(pending.id));
    try {
      expect(reopened.readCells({ sheet, range: 'A6:C6' }).cells.map(cell => cell.formula)).toEqual(['LET(x,2,x*4)', null, 'SEQUENCE(2,2)']);
    } finally { reopened.close(); }
  } finally { book.close(); }
});
