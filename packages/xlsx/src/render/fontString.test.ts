import { expect, test } from 'bun:test';
import { cellScriptFamilies, cellScriptFontKey, type CellScriptFamilies } from '../../../../shared/office-session/xlsxScriptFonts';
import { fontString } from './fontString';

test('keeps the original cell font stack for CJK punctuation', () => {
  const style = { fontFamily: 'Calibri', bold: false, italic: false };
  const families = new Map([
    [cellScriptFontKey({ ...style, script: 'cjk-sc' }), 'Noto Sans SC'],
    [cellScriptFontKey({ ...style, script: 'arabic' }), 'Noto Sans Arabic'],
  ]);
  const resolve: CellScriptFamilies = (cmd) => cellScriptFamilies(cmd, families);
  const cell = { op: 'text', x: 0, y: 0, color: '#000000', fontSize: 12, ...style } as const;
  for (const text of ['。', '（注）']) {
    expect(fontString({ ...cell, text }, resolve)).toBe('16px "Calibri", sans-serif');
  }
  expect(fontString({ ...cell, text: 'مرحبا。' }, resolve)).toBe('16px "Calibri", "Noto Sans Arabic", sans-serif');
});
