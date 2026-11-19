import { expect, test } from 'bun:test';
import type { DocxClipboardFragment } from '@betteroffice/docx/yrs';
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
