import { expect, test } from 'bun:test';
import { MAX_PASTE_FORMATTED_RUNS, type DocxClipboardFragment, type DocxTableFragment } from '@betteroffice/docx/yrs';
import { clipboardFragmentPlainText, pasteOptionsOffered, plainTextClipboardFragment } from './pasteText';

const pasteOptionsApply = (fragment: DocxClipboardFragment): boolean => pasteOptionsOffered(fragment).length > 0;

test('keep text only splits on the same line separators as a plain paste', () => {
  expect(plainTextClipboardFragment('a\u000Bb\u{2028}c\u{2029}d\r\ne\rf')).toEqual({
    paragraphs: ['a', 'b', 'c', 'd', 'e', 'f'].map((text) => ({ text })),
  });
});

test('plain text of a table-only fragment keeps the cell text', () => {
  const cell = (text: string) => ({ tcPr: {}, content: [{ kind: 'text' as const, text }, { kind: 'pilcrow' as const }] });
  expect(clipboardFragmentPlainText({
    paragraphs: [{ text: '', tables: [{ tblPr: {}, grid: [], rows: [{ trPr: {}, cells: [cell('a'), cell('b')] }, { trPr: {}, cells: [cell('c'), cell('d')] }] }] }],
  })).toBe('a\tb\nc\td');
});

test('paste options are offered only for formatted text without pictures', () => {
  const bold = { paragraphs: [{ text: 'Bold', runs: [{ text: 'Bold', format: { other: { bold: true } } }] }] };
  expect(pasteOptionsApply(bold)).toBe(true);
  expect(pasteOptionsApply({ paragraphs: [{ text: '  ', runs: [{ text: '  ', format: { other: { bold: true } } }] }] })).toBe(false);
  expect(pasteOptionsApply({ paragraphs: [{ text: 'Plain' }] })).toBe(false);
  expect(pasteOptionsApply({ paragraphs: [{ text: '', inlines: [{ image: { src: 'data:image/png;base64,AA==' } }] }] })).toBe(false);
});

test('paste options are offered after an in-editor copy only when its runs carry formatting', () => {
  const marked = { paragraphs: [{ text: 'Bold', runs: [{ text: 'Bold', format: { other: { bold: true } } }] }] };
  expect(pasteOptionsApply(marked)).toBe(true);
  const unmarked = { paragraphs: [{ text: 'Plain', runs: [{ text: 'Plain', format: { other: {} } }] }] };
  expect(pasteOptionsApply(unmarked)).toBe(false);
});

test('a formatted paste that holds a picture offers no paste options and a plain table paste offers them', () => {
  const bold = { text: 'Bold', format: { other: { bold: true } } };
  expect(pasteOptionsApply({ paragraphs: [{ text: 'Bold', runs: [bold] }] })).toBe(true);
  expect(pasteOptionsApply({ paragraphs: [{ text: 'Bold', runs: [bold],
    inlines: [{ text: 'Bold' }, { image: { src: 'data:image/png;base64,AA==' } }] }] })).toBe(false);
  const cell = (text: string) => ({ tcPr: {}, content: [{ kind: 'text' as const, text }, { kind: 'pilcrow' as const }] });
  const table = { paragraphs: [{ text: '', tables: [{ tblPr: {}, grid: [], rows: [{ trPr: {}, cells: [cell('a'), cell('b')] }] }] }] };
  expect(clipboardFragmentPlainText(table)).toBe('a\tb');
  expect(pasteOptionsApply(table)).toBe(true);
});

type CellContent = DocxTableFragment['rows'][number]['cells'][number]['content'];

function rowTable(...cells: CellContent[]): DocxTableFragment {
  return { tblPr: {}, grid: [], rows: [{ trPr: {}, cells: cells.map((content) => ({ tcPr: {}, content })) }] };
}

function cellText(content: CellContent): string {
  return clipboardFragmentPlainText({ paragraphs: [{ text: '', tables: [rowTable(content)] }] });
}

test('keep text only keeps the leading spaces of a table cell', () => {
  expect(cellText([{ kind: 'text', text: '  indented' }, { kind: 'pilcrow' }])).toBe('  indented');
});

test('keep text only keeps a tab that starts a table cell', () => {
  expect(cellText([{ kind: 'tab' }, { kind: 'text', text: 'tabbed' }, { kind: 'pilcrow' }])).toBe('\ttabbed');
});

test('keep text only puts a table nested in a cell on its own lines between the text before and after it', () => {
  const nested = rowTable([{ kind: 'text', text: 'a' }, { kind: 'pilcrow' }], [{ kind: 'text', text: 'b' }, { kind: 'pilcrow' }]);
  expect(cellText([{ kind: 'text', text: 'pre' }, { kind: 'table', table: nested }, { kind: 'text', text: 'post' }, { kind: 'pilcrow' }]))
    .toBe('pre\na\tb\npost');
});

test('keep text only joins the paragraphs of a table cell with one space and adds none at the end', () => {
  expect(cellText([{ kind: 'text', text: 'first' }, { kind: 'pilcrow' }, { kind: 'text', text: 'second' }, { kind: 'pilcrow' }]))
    .toBe('first second');
});

test('keep text only keeps a space typed at the end of a table cell', () => {
  expect(cellText([{ kind: 'text', text: 'end ' }, { kind: 'pilcrow' }])).toBe('end ');
});

test('keep text only is offered at 1000 lines of text and not at 1001', () => {
  const lines = (count: number): DocxClipboardFragment => ({
    paragraphs: Array.from({ length: count }, (_, index) => index === 0 ? { text: 'x', runs: [{ text: 'x', format: { other: { bold: true } } }] } : { text: 'x' }),
  });
  expect(pasteOptionsOffered(lines(1000))).toEqual(['keep-source', 'merge', 'text-only']);
  expect(pasteOptionsOffered(lines(1001))).toEqual(['keep-source', 'merge']);
});

test('paste options are not offered for a paste over the formatted run limit', () => {
  const fragment = (count: number): DocxClipboardFragment => {
    const runs = Array.from({ length: count }, (_, i) => ({ text: 'x', format: { other: i % 2 === 0 ? { bold: true } : {} } }));
    return { paragraphs: [{ text: 'x'.repeat(count), runs }] };
  };
  expect(pasteOptionsApply(fragment(MAX_PASTE_FORMATTED_RUNS))).toBe(true);
  expect(pasteOptionsApply(fragment(MAX_PASTE_FORMATTED_RUNS + 1))).toBe(false);
});
