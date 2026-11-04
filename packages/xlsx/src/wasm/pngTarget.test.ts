import { describe, expect, test } from 'bun:test';
import { cappedExtent, rangeA1, resolvePngTarget, sheetIndex } from './pngResolve';

const names = ['Data', 'Other'];
const viewport = { x: 0, y: 0, width: 300, height: 120 };

describe('PNG targets', () => {
  test('a range is an A1 string, an A1 object or zero-based row and column corners', () => {
    expect(rangeA1('B2:C3')).toBe('B2:C3');
    expect(rangeA1({ kind: 'a1', a1: 'b2' })).toBe('b2');
    expect(rangeA1({ kind: 'rowCol', start: { row: 1, col: 1 }, end: { row: 2, col: 27 } })).toBe('B2:AB3');
    expect(rangeA1({ kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 } })).toBe('A1:A1');
    expect(rangeA1({ kind: 'rowCol', start: { row: 0, col: 701 }, end: { row: 0, col: 702 } })).toBe('ZZ1:AAA1');
  });

  test.each([
    [{ kind: 'a1', a1: 'A1', extra: 1 }],
    [{ kind: 'a1' }],
    [{ kind: 'rowCol', start: { row: 0, col: 0 } }],
    [{ kind: 'rowCol', start: { row: 0, col: 0, extra: 1 }, end: { row: 1, col: 1 } }],
    [{ kind: 'rowCol', start: { row: -1, col: 0 }, end: { row: 1, col: 1 } }],
    [{ kind: 'rowCol', start: { row: 0.5, col: 0 }, end: { row: 1, col: 1 } }],
    [{ kind: 'rowCol', start: { row: true, col: 0 }, end: { row: 1, col: 1 } }],
    [{ kind: 'other' }],
    [7],
    [null],
    [['A1']],
  ])('refuses the range %j', (range) => {
    expect(() => rangeA1(range)).toThrow('range must be an A1 string or a range address');
  });

  test('refuses a range whose start is below or right of its end', () => {
    expect(() => rangeA1({ kind: 'rowCol', start: { row: 2, col: 0 }, end: { row: 1, col: 1 } }))
      .toThrow('range start must be above and left of its end');
  });

  test('a sheet is a zero-based index or a tab name', () => {
    expect(sheetIndex(1, names)).toBe(1);
    expect(sheetIndex('Other', names)).toBe(1);
    expect(() => sheetIndex(2, names)).toThrow('sheet index 2 out of range for 2 sheet(s)');
    expect(() => sheetIndex('Missing', names)).toThrow('no sheet named "Missing"');
    expect(() => sheetIndex(true, names)).toThrow('sheet must be a name (string) or an index (number)');
    expect(() => sheetIndex(-1, names)).toThrow('sheet must be a name (string) or an index (number)');
  });

  test('resolves targets and checks scale and size limits before any engine call', () => {
    expect(resolvePngTarget({ sheet: 'Other', viewport }, {}, names)).toEqual({ sheet: 1, viewport });
    expect(resolvePngTarget({ sheet: 0, range: { kind: 'a1', a1: 'A1:B2' } }, { scale: 2 }, names)).toEqual({ sheet: 0, range: 'A1:B2' });
    expect(() => resolvePngTarget({ sheet: 0, viewport }, { scale: 0 }, names)).toThrow('scale must be a positive number');
    expect(() => resolvePngTarget({ sheet: 0, viewport }, { scale: Number.NaN }, names)).toThrow('scale must be a positive number');
    expect(() => resolvePngTarget({ sheet: 0, viewport }, { maxWidth: 0 }, names)).toThrow('maxWidth must be a positive integer');
    expect(() => resolvePngTarget({ sheet: 0, range: 'A1' }, { maxHeight: 1.5 }, names)).toThrow('maxHeight must be a positive integer');
    expect(() => resolvePngTarget({ sheet: 0, viewport: { ...viewport, width: Number.POSITIVE_INFINITY } }, {}, names)).toThrow('finite');
    expect(() => resolvePngTarget({ sheet: 0, viewport, range: 'A1' }, {}, names)).toThrow('target must be');
    expect(() => resolvePngTarget({ sheet: 0 }, {}, names)).toThrow('target must be');
  });

  test('caps an extent as the engine does, never overshooting by a pixel at fractional scales', () => {
    expect(cappedExtent(300, undefined, 1.3)).toBe(Math.fround(300));
    expect(cappedExtent(10, 100, 1)).toBe(10);
    for (const scale of [0.75, 1, 1.3, 2.7, 3]) {
      for (const cap of [1, 23, 37, 97, 1000]) {
        const extent = cappedExtent(5000, cap, scale);
        expect(Math.ceil(Math.fround(extent * Math.fround(scale)))).toBeLessThanOrEqual(cap);
        expect(Math.ceil(Math.fround(extent * Math.fround(scale)))).toBeGreaterThanOrEqual(cap - 1);
      }
    }
  });
});
