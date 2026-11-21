import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachXlsx, openXlsx, type XlsxAgentWorkbook } from '../src';
import { implicitFontWorkbook } from './format-fixtures';

const author = 'Agent';
const sheet = 'sheet:0';

const numericRow = (row: number, columns: string) => `<row r="${row}">${[...columns].map(column => `<c r="${column}${row}"><v>10</v></c>`).join('')}</row>`;

async function withXlsx(rows: string[], run: (book: XlsxAgentWorkbook) => Promise<void>) {
  const zip = await JSZip.loadAsync(await implicitFontWorkbook());
  const part = 'xl/worksheets/sheet1.xml';
  const xml = await zip.file(part)!.async('string');
  zip.file(part, xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${rows.join('')}</sheetData>`));
  const owner = await openXlsx(await zip.generateAsync({ type: 'uint8array' }));
  const book = attachXlsx(owner.handle);
  try { await run(book); } finally { book.close(); owner.close(); }
}

const proposeCell = (book: XlsxAgentWorkbook, range: string) =>
  book.proposeCells({ author, edits: [{ cell: book.readCells({ sheet, range }).cells[0].cell, input: '99' }] });

test('attached XLSX refuses to export, verify or accept a proposal after rows move its target', () => withXlsx([numericRow(2, 'A'), numericRow(3, 'A')], async book => {
  const { id } = proposeCell(book, 'A3');
  expect(book.handle.applyOps([{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]).applied).toBe(true);
  expect(book.review(id).stale).toBe(true);
  await expect(book.export(id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  await expect(book.verify(id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  await expect(book.accept(id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  expect(book.handle.searchText('99').filter(hit => hit.text === '99')).toEqual([]);
}));

test('attached XLSX lists pending proposals without a trial accept per proposal', () => {
  const columns = 'ABCDEFGHIJ';
  return withXlsx(Array.from({ length: 1000 }, (_, index) => numericRow(index + 1, columns)), async book => {
    for (let row = 1; row <= 32; row++) proposeCell(book, `A${row}`);
    const handle = book.handle as typeof book.handle & { isProposalStale(id: string): boolean };
    const original = handle.isProposalStale;
    let trials = 0;
    handle.isProposalStale = (id: string) => { trials++; return original.call(handle, id); };
    try {
      const samples = Array.from({ length: 5 }, () => {
        const start = performance.now();
        expect(book.listProposals()).toHaveLength(32);
        return performance.now() - start;
      }).sort((a, b) => a - b);
      expect(trials).toBe(0);
      expect(samples[2]).toBeLessThan(100);
    } finally { handle.isProposalStale = original; }
  });
});
