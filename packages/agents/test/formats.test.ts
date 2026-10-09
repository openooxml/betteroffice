import { expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { openPptx, openXlsx } from '../src';
import { pptxFixture, xlsxFixture } from './format-fixtures';

async function unchangedParts(before: Uint8Array, after: Uint8Array, changed: string[]) {
  const source = await JSZip.loadAsync(before);
  const result = await JSZip.loadAsync(after);
  for (const [path, file] of Object.entries(source.files)) {
    if (file.dir || changed.includes(path)) continue;
    if (path === '_rels/.rels') {
      const before = (await file.async('string')).replace(/>\s+</g, '><');
      const after = (await result.file(path)!.async('string')).replace(/>\s+</g, '><');
      expect(after, path).toBe(before);
    } else {
      expect(await result.file(path)!.async('uint8array'), path).toEqual(await file.async('uint8array'));
    }
  }
}

test('XLSX paginates cells, stages atomic inputs/formulas, exports privately, and refuses stale proposals', async () => {
  const source = await xlsxFixture();
  const book = await openXlsx(source);
  try {
    expect(book.overview()).toMatchObject({ format: 'xlsx', sheets: 3 });
    const baseline = await book.export();
    const sheet = book.list().items[0].sheetId;
    const page = book.readCells({ sheet, range: 'B3:E3', limit: 2 });
    expect(page.nextOffset).toBe(2);
    expect(page.cells[0].value).toEqual({ kind: 'number', value: 100 });
    const next = book.readCells({ sheet, range: 'B3:E3', offset: page.nextOffset! });
    expect(next.cells[0].formula).toBe('B3+C3');
    const pending = book.proposeCells({ author: 'test', edits: [{ cell: page.cells[0].cell, input: '1000' }, { cell: next.cells[1].cell, input: '=D3*2' }] });
    expect(book.read(page.cells[0].ref).value).toEqual({ kind: 'number', value: 100 });
    expect((await book.verify(pending.id)).reopened).toBe(true);
    const exported = await book.export(pending.id);
    const reopened = await openXlsx(exported);
    try {
      const cells = reopened.readCells({ sheet, range: 'B3:E3' }).cells;
      expect(cells.map(cell => cell.value)).toEqual([{ kind: 'number', value: 1000 }, { kind: 'number', value: 57 }, { kind: 'number', value: 1057 }, { kind: 'number', value: 2114 }]);
      expect(cells[3].formula).toBe('D3*2');
    } finally { reopened.close(); }
    await unchangedParts(baseline, exported, ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/workbook.xml', 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']);
    const stale = book.proposeCells({ author: 'test', edits: [{ cell: page.cells[1].cell, input: '9' }] });
    await expect(book.accept(pending.id, { tracked: true })).rejects.toMatchObject({ code: 'UNSUPPORTED' });
    await book.accept(pending.id);
    expect(book.review(stale.id).stale).toBe(true);
    await expect(book.export(stale.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
    expect(() => book.proposeCells({ author: 'test', edits: [{ cell: page.cells[1].cell, input: '8' }] })).toThrow('changed');
    expect(book.reject(stale.id).status).toBe('rejected');
    expect(book.read(page.cells[0].ref).value).toEqual({ kind: 'number', value: 1000 });
  } finally { book.close(); }
  expect(() => book.list()).toThrow('closed');
});

test('XLSX search paginates, excludes foreign handles, and refuses invalid ranges and unsafe writes', async () => {
  const bytes = await xlsxFixture();
  const book = await openXlsx(bytes);
  const other = await openXlsx(bytes);
  try {
    const first = book.grep({ query: 'Line item', limit: 2 });
    expect(first.matches).toHaveLength(2);
    expect(first.nextCursor).toBeString();
    const second = book.grep({ query: 'Line item', limit: 2, cursor: first.nextCursor! });
    expect(second.matches[0].a1).not.toBe(first.matches[0].a1);
    expect(() => book.grep({ query: 'different', cursor: first.nextCursor! })).toThrow('Search changed');
    expect(() => book.grep({ query: 'line', caseSensitive: false })).toThrow('case-sensitive');
    expect(() => other.proposeCells({ author: 'test', edits: [{ cell: first.matches[0].cell, input: 'changed' }] })).toThrow('another workbook');
    expect(() => book.readCells({ sheet: 'sheet:0', range: 'B3:A1' })).toThrow('ordered');
    expect(() => book.readCells({ sheet: 'sheet:0', range: 'XFE1' })).toThrow('Excel grid');
    expect(() => book.proposeCells({ author: 'test', edits: [{ cell: first.matches[0].cell, input: 'a' }, { cell: first.matches[0].cell, input: 'b' }] })).toThrow();
    const merged = book.readCells({ sheet: 'sheet:0', range: 'B1' }).cells[0];
    expect(() => book.proposeCells({ author: 'test', edits: [{ cell: merged.cell, input: 'unsafe' }] })).toThrow();
    expect(book.grep({ query: 'Line item', limit: 2 }).matches[0].text).toBe(first.matches[0].text);
  } finally { book.close(); other.close(); }
});

test('PPTX replaces one occurrence, preserves formatting and package parts, and rejects stale handles', async () => {
  const bytes = await pptxFixture();
  const deck = await openPptx(bytes);
  const other = await openPptx(bytes);
  try {
    expect(deck.overview()).toMatchObject({ format: 'pptx', slides: 1, paragraphs: 2 });
    const hits = deck.grep({ query: 'risk', limit: 1 });
    expect(hits.matches[0].text).toBe('Risk');
    const next = deck.grep({ query: 'risk', limit: 1, cursor: hits.nextCursor! });
    expect(next.matches[0].text).toBe('risk');
    expect(next.nextCursor).toBeNull();
    const edit = { match: next.matches[0].match, newText: 'opportunity' };
    expect(() => other.propose({ author: 'test', edits: [edit] })).toThrow('another deck');
    expect(() => deck.propose({ author: 'test', edits: [edit, edit] })).toThrow();
    const pending = deck.propose({ author: 'test', edits: [edit] });
    const stale = deck.propose({ author: 'test', edits: [{ match: hits.matches[0].match, newText: 'Exposure' }] });
    const proposed = await deck.export(pending.id);
    expect(deck.read(next.matches[0].ref).text).toContain('Risk risk.');
    const reopened = await openPptx(proposed);
    try { expect(reopened.grep({ query: 'Risk opportunity.' }).matches).toHaveLength(1); }
    finally { reopened.close(); }
    await unchangedParts(bytes, proposed, ['ppt/slides/slide1.xml']);
    const xml = await (await JSZip.loadAsync(proposed)).file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).toContain('b="1" sz="2400"');
    expect(xml).toContain('val="1565C0"');
    expect((await deck.verify(pending.id)).reopened).toBe(true);
    await deck.accept(pending.id);
    expect(deck.review(stale.id).stale).toBe(true);
    await expect(deck.accept(stale.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
    expect(() => deck.grep({ query: 'risk', cursor: hits.nextCursor! })).toThrow('Search changed');
    expect(() => deck.propose({ author: 'test', edits: [edit] })).toThrow('changed');
    expect(deck.reject(stale.id).status).toBe('rejected');
  } finally { deck.close(); other.close(); }
});

test('PPTX long reads and search stay bounded and reject paragraph crossings', async () => {
  const deck = await openPptx(await pptxFixture(`${'sample '.repeat(5000)}😀`));
  try {
    const hits = deck.grep({ query: 'sample', limit: 100 });
    expect(JSON.stringify(hits).length).toBeLessThan(17000);
    const ref = hits.matches[0].ref;
    const read = deck.read(ref, { length: 100 });
    expect(read.nextStart).toBe(100);
    const tail = deck.read(ref, { start: 35001, length: 1 });
    expect(tail.text).toBe('😀');
    expect(() => deck.propose({ author: 'test', edits: [{ match: hits.matches[0].match, newText: 'a\nb' }] })).toThrow('one paragraph');
    expect(() => deck.grep({ query: 'missing', story: 'missing' })).toThrow('story');
    await expect(deck.render()).rejects.toMatchObject({ code: 'UNSUPPORTED' });
  } finally { deck.close(); }
});


test('XLSX pages and search hydrate cell handles with one engine read', async () => {
  const book = await openXlsx(await xlsxFixture());
  const internal = book as unknown as { handle: import('@betteroffice/xlsx').WorkbookHandle };
  const reads = spyOn(internal.handle, 'readCells');
  try {
    const page = book.readCells({ sheet: 'sheet:0', range: 'A3:E30', offset: 3, limit: 100 });
    expect(reads).toHaveBeenCalledTimes(1);
    expect(reads.mock.calls[0][0].ranges).toHaveLength(100);
    expect(page.cells[0].a1).toBe('D3');
    expect(page.nextOffset).toBe(3 + page.cells.length);
    expect(JSON.stringify(page).length).toBeLessThan(17000);
    expect(new Set(page.cells.map(cell => cell.version)).size).toBe(1);
    reads.mockClear();
    const matches = book.grep({ query: 'Line item', limit: 20 }).matches;
    expect(reads).toHaveBeenCalledTimes(1);
    expect(reads.mock.calls[0][0].ranges).toHaveLength(matches.length);
    expect(() => book.readCells({ sheet: 'missing', range: 'A1', offset: 1 })).toThrow('sheetId');
  } finally { reads.mockRestore(); book.close(); }
});
