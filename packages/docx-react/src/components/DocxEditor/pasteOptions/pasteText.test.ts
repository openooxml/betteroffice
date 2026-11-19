import { expect, test } from 'bun:test';
import { clipboardFragmentPlainText, pasteOptionsApply, plainTextClipboardFragment } from './pasteText';

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
  expect(pasteOptionsApply(bold, 'Bold')).toBe(true);
  expect(pasteOptionsApply(bold, '  ')).toBe(false);
  expect(pasteOptionsApply({ paragraphs: [{ text: 'Plain' }] }, 'Plain')).toBe(false);
  expect(pasteOptionsApply({ paragraphs: [{ text: '', inlines: [{ image: { src: 'data:image/png;base64,AA==' } }] }] }, 'x')).toBe(false);
});

test('paste options are offered after an in-editor copy only when its runs carry formatting', () => {
  const marked = { paragraphs: [{ text: 'Bold', runs: [{ text: 'Bold', format: { other: { bold: true } } }] }] };
  expect(pasteOptionsApply(marked, 'Bold')).toBe(true);
  const unmarked = { paragraphs: [{ text: 'Plain', runs: [{ text: 'Plain', format: { other: {} } }] }] };
  expect(pasteOptionsApply(unmarked, 'Plain')).toBe(false);
});
