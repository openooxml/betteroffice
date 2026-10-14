import { expect, test } from 'bun:test';
import { normalFontIndex } from './xlsx-styles';

test('resolves localized Normal through its builtin ID and xfId', () => {
  expect(
    normalFontIndex(
      [
        { name: 'Normal', xfId: 0 },
        { name: '標準', builtinId: 0, xfId: 2 },
      ],
      [3, 1, 4]
    )
  ).toBe(4);
});

test('uses the named Normal style when the builtin ID is omitted', () => {
  expect(normalFontIndex([{ name: 'Normal', xfId: 1 }], [3, 2])).toBe(2);
});

test('falls back to style XF zero, then font zero when Normal data is absent', () => {
  expect(normalFontIndex([], [3, 2])).toBe(3);
  expect(normalFontIndex([], [])).toBe(0);
  expect(normalFontIndex([{ builtinId: 0, xfId: 99 }], [3])).toBe(0);
});
