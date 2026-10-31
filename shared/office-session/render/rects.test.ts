import { expect, test } from 'bun:test';
import { clampToSurface, intersects, isEmpty, snapOut, union } from './rects';

test('fractional positive and negative bounds snap outward', () => {
  expect(snapOut({ x: -0.2, y: 1.4, width: 3.1, height: 2.2 }))
    .toEqual({ x: -1, y: 1, width: 4, height: 3 });
  expect(isEmpty(snapOut({ x: 0.5, y: 0.5, width: 0, height: 2 }))).toBe(true);
});

test('union, strict intersection, surface clipping and empty bounds', () => {
  const a = { x: 1, y: 2, width: 3, height: 4 }, b = { x: 3, y: 1, width: 4, height: 2 };
  expect(union(a, b)).toEqual({ x: 1, y: 1, width: 6, height: 5 });
  expect(intersects(a, b)).toBe(true);
  expect(intersects(a, { ...a, x: 4 })).toBe(false);
  expect(intersects(a, { ...a, width: 0 })).toBe(false);
  expect(union({ ...a, width: 0 }, b)).toEqual(b);
  expect(union(a, { ...b, height: 0 })).toEqual(a);
  expect(clampToSurface({ x: -1, y: 2, width: 10, height: 8 }, 5, 6))
    .toEqual({ x: 0, y: 2, width: 5, height: 4 });
  expect(isEmpty(clampToSurface({ ...a, x: 20 }, 5, 6))).toBe(true);
  expect(isEmpty({ ...a, height: -1 })).toBe(true);
});
