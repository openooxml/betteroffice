import { expect, test } from 'bun:test';
import type { ShadingProperties } from '../types/colors';
import { tableCellAttrsToFormatting } from './saveFormatting';

test('unchanged empty resolved fill preserves authored automatic shading', () => {
  const shading: ShadingProperties = { pattern: 'clear', color: { auto: true }, fill: { auto: true } };
  expect(tableCellAttrsToFormatting({
    colspan: 2, rowspan: 1,
    _originalFormatting: { shading },
  })?.shading).toEqual(shading);
  expect(tableCellAttrsToFormatting({
    colspan: 1, rowspan: 1,
    _originalFormatting: { shading: { fill: { rgb: 'FF0000' } } },
    _originalResolvedFill: 'FF0000',
  })?.shading).toBeUndefined();
});
