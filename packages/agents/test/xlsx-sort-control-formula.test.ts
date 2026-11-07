import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openXlsx } from '../src/xlsx';
import { implicitFontWorkbook } from './format-fixtures';

test('XLSX sort moves a formula holding a control character as its CHAR form', async () => {
  const zip = await JSZip.loadAsync(await implicitFontWorkbook());
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('<c r="C2" s="1"><v>50</v></c>', '<c r="C2" t="str"><f>"a&#1;b"</f><v>x</v></c>'));
  const book = await openXlsx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const inputs = () => book.handle.rangeCells(0, 'A2:C4').map(row => row.map(cell => cell.input));
    expect(inputs()[0][2]).toBe('="a\u0001b"');
    const sorted = book.sortRange({ version: book.handle.version(), sheet: 'Sales', range: 'A2:C4', keys: [{ column: 'B', order: 'asc' }] });
    expect(sorted.applied).toBe(true);
    expect(inputs().map(row => row[0])).toEqual(['Bottle', 'Cap', 'Bag']);
    expect(inputs()[1][2]).toBe('=("a"&CHAR(1)&"b")');
    const saved = await JSZip.loadAsync(await book.export());
    const savedSheet = await saved.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(savedSheet).toContain('<f>(&quot;a&quot;&amp;CHAR(1)&amp;&quot;b&quot;)</f>');
    expect(savedSheet).not.toContain('&#1;');
  } finally {
    book.close();
  }
});
