import { expect, test } from 'bun:test';
import { fontString } from './fontString';

test('puts the Arabic pack family between the cell font and the generic fallback', () => {
  expect(fontString({
    op: 'text', x: 0, y: 0, text: 'مرحبا', color: '#000000', fontSize: 12, fontFamily: 'Arial',
  })).toBe('16px "Arial", "Noto Sans Arabic", sans-serif');
});
