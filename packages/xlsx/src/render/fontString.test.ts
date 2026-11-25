import { expect, test } from 'bun:test';
import { fontString } from './fontString';

test('keeps the original cell font stack for CJK punctuation', () => {
  for (const text of ['。', '（注）']) {
    expect(fontString({
      op: 'text', x: 0, y: 0, text, color: '#000000', fontSize: 12, fontFamily: 'Calibri',
    })).toBe('16px "Calibri", sans-serif');
  }
});
