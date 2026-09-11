import { expect, test } from 'bun:test';
import { capturePageExtent, validatePageBounds } from './page-bounds';

const a4 = { width_pt: 595.25, height_pt: 842, width_px: 1241, height_px: 1755 };
const letter = { width_pt: 612, height_pt: 792, width_px: 1275, height_px: 1650 };
const profile = (pages = [a4]) => ({ kind: 'office-page-bounds', pages });

test('uses the recorded PDF extent for one-pixel A4 canvas differences', () => {
  const bounds = validatePageBounds(profile());
  expect(capturePageExtent(bounds, 0, 1240, 1754)).toEqual({
    page: 1,
    raw: { width_px: 1240, height_px: 1754 },
    output: { width_px: 1241, height_px: 1755 },
    adjustment: 'office-page-bounds',
  });
  expect(capturePageExtent(bounds, 0, 1242, 1756).output).toEqual({
    width_px: 1241,
    height_px: 1755,
  });
});

test('leaves matching Letter pages and captures without a profile unchanged', () => {
  for (const bounds of [validatePageBounds(profile([letter])), null]) {
    const extent = capturePageExtent(bounds, 0, 1275, 1650);
    expect(extent.output).toEqual(extent.raw);
    expect(extent.adjustment).toBe('none');
  }
  expect(validatePageBounds(undefined)).toBeNull();
  expect(validatePageBounds(null)).toBeNull();
});

test('preserves extra actual pages at native size', () => {
  const extent = capturePageExtent(validatePageBounds(profile()), 1, 2550, 1650);
  expect(extent).toEqual({
    page: 2,
    raw: { width_px: 2550, height_px: 1650 },
    output: { width_px: 2550, height_px: 1650 },
    adjustment: 'none',
  });
});

test('rejects larger geometry differences on either axis', () => {
  const bounds = validatePageBounds(profile());
  for (const [width, height] of [
    [1239, 1755],
    [1243, 1755],
    [1241, 1753],
    [1241, 1757],
    [1755, 1241],
  ]) {
    expect(() => capturePageExtent(bounds, 0, width, height)).toThrow('more than 1 pixel');
  }
});

test('validates physical sizes and exact 150-DPI raster extents', () => {
  for (const page of [
    { ...a4, width_pt: NaN },
    { ...a4, height_pt: Infinity },
    { ...a4, width_pt: 0 },
    { ...a4, height_pt: -1 },
    { ...a4, width_pt: 10000 },
    { ...a4, width_px: 0 },
    { ...a4, height_px: 1755.5 },
    { ...a4, width_px: 8193 },
    { ...a4, width_px: 1240 },
    { ...a4, height_px: 1754 },
    { ...letter, width_px: 1224, height_px: 1584 },
  ]) {
    expect(() => validatePageBounds(profile([page]))).toThrow('150 DPI');
  }
  const fractional = { width_pt: 600.01, height_pt: 800.01, width_px: 1251, height_px: 1667 };
  expect(validatePageBounds(profile([fractional]))?.pages).toEqual([fractional]);
});

test('rejects unknown or malformed profiles and excessive page counts', () => {
  for (const value of [
    false,
    'office-page-bounds',
    {},
    { kind: 'other', pages: [a4] },
    { kind: 'office-page-bounds', pages: null },
    profile([]),
    profile(Array.from({ length: 101 }, () => a4)),
  ]) {
    expect(() => validatePageBounds(value)).toThrow('Invalid Office page bounds profile');
  }
  expect(() => validatePageBounds({ kind: 'office-page-bounds', pages: [null] })).toThrow(
    '150 DPI'
  );
  expect(validatePageBounds(profile(Array.from({ length: 100 }, () => a4)))?.pages).toHaveLength(100);
});

test('rejects invalid page indices and native canvas dimensions', () => {
  for (const [index, width, height] of [
    [-1, 1275, 1650],
    [0.5, 1275, 1650],
    [0, 0, 1650],
    [0, 1275, Infinity],
  ]) {
    expect(() => capturePageExtent(null, index, width, height)).toThrow('native canvas');
  }
});
