import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { attachDocx, attachXlsx, openDocx, openXlsx, type DocxAgentDocument, type XlsxAgentWorkbook } from '../src';
import { fixture } from './fixture';
import { implicitFontWorkbook } from './format-fixtures';

const author = 'Agent';
const sheet = 'sheet:0';

async function withDocx(texts: string[], run: (doc: DocxAgentDocument) => Promise<void>) {
  const zip = await JSZip.loadAsync(await fixture());
  const part = 'word/document.xml';
  const xml = await zip.file(part)!.async('string');
  zip.file(part, xml.replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${texts.map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join('')}`));
  const owner = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
  const doc = attachDocx(owner.session);
  try { await run(doc); } finally { doc.close(); owner.close(); }
}

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

const paragraphTarget = (doc: DocxAgentDocument, index: number) => {
  const { ref, revision } = doc.list({ story: 'body' }).items[index];
  return { ref, revision };
};

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

test('attached DOCX accept is one undo step that leaves earlier typing in place', () => withDocx(['', 'old', 'old2'], async doc => {
  const { session } = doc;
  const [{ paraId }] = session.paragraphs('body');
  session.addUndoBoundary();
  session.insertText({ story: 'body', paraId, offset: 0 }, 'hello');
  session.addUndoBoundary();
  expect(session.paragraphs('body')[0]?.text).toBe('hello');
  const { id } = doc.propose({ author, edits: [
    { ...paragraphTarget(doc, 1), start: 0, oldText: 'old', newText: 'new' },
    { ...paragraphTarget(doc, 2), start: 0, oldText: 'old2', newText: 'new2' },
  ] });
  await doc.accept(id);
  expect(session.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['hello', 'new', 'new2']);
  expect(session.canUndo()).toBe(true);
  session.addUndoBoundary();
  session.insertText({ story: 'body', paraId, offset: 0 }, 'X');
  session.addUndoBoundary();
  expect(session.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Xhello', 'new', 'new2']);
  expect(session.canUndo()).toBe(true);
  session.undo();
  expect(session.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['hello', 'new', 'new2']);
  expect(session.canUndo()).toBe(true);
  session.undo();
  expect(session.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['hello', 'old', 'old2']);
  expect(session.canUndo()).toBe(true);
}));

test('attached DOCX tracked accept takes the given date', () => withDocx(['old'], async doc => {
  const date = '2026-01-01T00:00:00Z';
  const { id } = doc.propose({ author, edits: [{ ...paragraphTarget(doc, 0), start: 0, oldText: 'old', newText: 'new' }] });
  expect(await doc.accept(id, { tracked: true, date })).toMatchObject({ id, status: 'accepted', tracked: true });
  expect(doc.review(id).status).toBe('accepted');
  const revisions = doc.session.listRevisions();
  expect(revisions.length).toBeGreaterThan(0);
  expect(revisions.every(revision => Date.parse(revision.date ?? '') === Date.parse(date))).toBe(true);
}));

test('attached DOCX refuses a second proposal in a paragraph an accept edited back to the same text', () => withDocx(['a a a a'], async doc => {
  const target = paragraphTarget(doc, 0);
  const first = doc.propose({ author, edits: [{ ...target, start: 0, oldText: '', newText: 'a ' }, { ...target, start: 5, oldText: ' a', newText: '' }] });
  const second = doc.propose({ author, edits: [{ ...target, start: 2, oldText: 'a', newText: 'B' }] });
  await doc.accept(first.id);
  expect(doc.session.paragraphs('body')[0]?.text).toBe('a a a a');
  await expect(doc.accept(second.id)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  expect(doc.session.paragraphs('body')[0]?.text).toBe('a a a a');
  expect(doc.review(second.id)).toMatchObject({ status: 'pending' });
  expect(doc.review(second.id).staleRefs.length).toBeGreaterThan(0);
  expect(doc.reject(second.id)).toMatchObject({ status: 'rejected' });
  expect(doc.session.paragraphs('body')[0]?.text).toBe('a a a a');
}));
