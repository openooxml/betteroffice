import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openXlsx } from '../src/xlsx';
import { implicitFontWorkbook } from './format-fixtures';

const extension = '<extLst><ext uri="{6E1A2B3C-4D5E-4F60-8A9B-0C1D2E3F4A5B}"><x:keep xmlns:x="urn:protected-sort" value="kept"/></ext></extLst>';
const sheetPath = 'xl/worksheets/sheet1.xml';

async function fixture() {
  const zip = await JSZip.loadAsync(await implicitFontWorkbook());
  const styles = await zip.file('xl/styles.xml')!.async('string');
  zip.file('xl/styles.xml', styles.replace('<cellXfs count="2">', '<cellXfs count="4">').replace('</cellXfs>', `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0">${extension}</xf></cellXfs>`));
  const sheet = await zip.file(sheetPath)!.async('string');
  zip.file(sheetPath, sheet.replace('<c r="A2" t="inlineStr">', '<c r="A2" s="3" t="inlineStr">'));
  return zip;
}

function xfs(styles: string) {
  return styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)![1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!;
}

function cell(sheet: string, address: string) {
  return sheet.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>[\\s\\S]*?<\\/c>`))![0];
}

function styleIndex(cellXml: string) {
  return Number(cellXml.slice(0, cellXml.indexOf('>')).match(/\bs="(\d+)"/)?.[1] ?? 0);
}

for (const scenario of [
  { name: 'ascending', order: 'asc', header: false, range: 'A2:B4', sourceRows: [4, 2, 3] },
  { name: 'descending', order: 'desc', header: false, range: 'A2:B4', sourceRows: [3, 2, 4] },
  { name: 'with a header', order: 'asc', header: true, range: 'A1:B4', sourceRows: [1, 4, 2, 3] },
] as const) {
  test(`XLSX unprotected sort ${scenario.name} keeps each source cell's exact style XML`, async () => {
    const zip = await fixture();
    const originalSheet = await zip.file(sheetPath)!.async('string');
    expect(originalSheet).not.toContain('<sheetProtection');
    const before = xfs(await zip.file('xl/styles.xml')!.async('string'));
    const book = await openXlsx(await zip.generateAsync({ type: 'uint8array' }));
    try {
      const sorted = book.sortRange({ version: book.handle.version(), sheet: 'Sales', range: scenario.range, keys: [{ column: 'B', order: scenario.order }], header: scenario.header });
      expect(sorted.applied).toBe(true);
      const saved = await JSZip.loadAsync(await book.export());
      const savedSheet = await saved.file(sheetPath)!.async('string');
      const after = xfs(await saved.file('xl/styles.xml')!.async('string'));
      expect(after.slice(0, before.length)).toEqual(before);
      const extendedCells: string[] = [];
      for (const [index, sourceRow] of scenario.sourceRows.entries()) {
        const destRow = index + (scenario.header ? 1 : 2);
        for (const column of ['A', 'B']) {
          const source = cell(originalSheet, `${column}${sourceRow}`);
          const dest = cell(savedSheet, `${column}${destRow}`);
          const sourceStyle = styleIndex(source);
          const destStyle = styleIndex(dest);
          expect(after[destStyle]).toBe(before[sourceStyle]);
          expect(dest.match(/<(?:v|t)\b[^>]*>(.*?)<\/(?:v|t)>/)?.[1]).toBe(source.match(/<(?:v|t)\b[^>]*>(.*?)<\/(?:v|t)>/)?.[1]);
          if (sourceStyle === 0) expect(destStyle).toBe(0);
          if (after[destStyle].includes(extension)) extendedCells.push(`${column}${destRow}`);
        }
      }
      expect(extendedCells).toEqual(['A3']);
    } finally { book.close(); }
  });
}
