import { describe, expect, test } from 'bun:test';
import { ACTIVE_CELL_WIDTH, MARQUEE_WIDTH, marqueeBox, marqueeFills } from './selectionOutline';
import type { MarqueeBox, PixelRect } from './selectionOutline';

/**
 * The one-pixel ring covers the canvas gridline on every edge: boundary `b`
 * paints `[b - 1, b]`, so the box starts one pixel before the snapped rect and
 * ends on its far edge.
 */
function expectAlignedRing(box: MarqueeBox) {
  const { inset } = box;
  expect(box.left).toBe(inset.x - 1);
  expect(box.top).toBe(inset.y - 1);
  expect(box.left + box.width).toBe(inset.x + inset.w);
  expect(box.top + box.height).toBe(inset.y + inset.h);
  expect(box.width).toBe(inset.w + 1);
  expect(box.height).toBe(inset.h + 1);
}

function expectNonNegativeBox(box: MarqueeBox) {
  for (const value of [box.left, box.top, box.width, box.height, box.inset.x, box.inset.y, box.inset.w, box.inset.h]) {
    expect(Number.isFinite(value)).toBe(true);
  }
  expect(box.inset.w).toBeGreaterThanOrEqual(0);
  expect(box.inset.h).toBeGreaterThanOrEqual(0);
  expect(box.width).toBeGreaterThanOrEqual(0);
  expect(box.height).toBeGreaterThanOrEqual(0);
}

const overlaps = (a: PixelRect, b: PixelRect) =>
  a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

function intersect(a: PixelRect, b: PixelRect): PixelRect {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  return {
    x,
    y,
    w: Math.max(0, Math.min(a.x + a.w, b.x + b.w) - x),
    h: Math.max(0, Math.min(a.y + a.h, b.y + b.h) - y),
  };
}

describe('marqueeBox', () => {
  test('snaps an integer rect to itself and keeps the ring on its edges', () => {
    const box = marqueeBox({ x: 100, y: 50, w: 200, h: 40 });
    // the ring covers the gridlines at [99, 100] and [299, 300]; a symmetric
    // half-pixel offset would blur at most zooms.
    expect(box).toEqual({ left: 99, top: 49, width: 201, height: 41, inset: { x: 100, y: 50, w: 200, h: 40 } });
    expectAlignedRing(box);
  });

  test('snaps fractional edges with Math.round, rounding the sum for the right/bottom edge', () => {
    const rect = { x: 100.4, y: 50.6, w: 199.3, h: 39.7 };
    // Math.round(100.4) = 100, Math.round(50.6) = 51,
    // Math.round(100.4 + 199.3) = Math.round(299.7) = 300, Math.round(50.6 + 39.7) = Math.round(90.3) = 90.
    expect(Math.round(rect.x)).toBe(100);
    expect(Math.round(rect.y)).toBe(51);
    expect(Math.round(rect.x + rect.w)).toBe(300);
    expect(Math.round(rect.y + rect.h)).toBe(90);
    // Width/height come from the snapped edges, not from rounding the size:
    // Math.round(199.3) = 199 and Math.round(39.7) = 40 are both off by one.
    expect(Math.round(rect.w)).not.toBe(200);
    expect(Math.round(rect.h)).not.toBe(39);

    const box = marqueeBox(rect);
    expect(box.inset).toEqual({ x: 100, y: 51, w: 200, h: 39 });
    expect(box.left).toBe(99);
    expect(box.top).toBe(50);
    expect(box.width).toBe(201);
    expect(box.height).toBe(40);
    expectAlignedRing(box);
  });

  test('covers the gridline pixel on every snapped edge', () => {
    const box = marqueeBox({ x: 100.4, y: 50.6, w: 199.3, h: 39.7 });
    // boundary `b` is drawn `[b - 1, b]`; the ring's four pixels are exactly
    // those four gridline pixels, so no edge reads as inset while another is
    // on the line.
    expect([box.left, box.left + 1]).toEqual([box.inset.x - 1, box.inset.x]);
    expect([box.left + box.width - 1, box.left + box.width])
      .toEqual([box.inset.x + box.inset.w - 1, box.inset.x + box.inset.w]);
    expect([box.top, box.top + 1]).toEqual([box.inset.y - 1, box.inset.y]);
    expect([box.top + box.height - 1, box.top + box.height])
      .toEqual([box.inset.y + box.inset.h - 1, box.inset.y + box.inset.h]);
    expect(MARQUEE_WIDTH).toBe(1);
  });

  test('returns a non-negative box for zero-size and negative-size rects', () => {
    for (const rect of [
      { x: 100, y: 50, w: 0, h: 0 },
      { x: 100, y: 50, w: -40, h: -20 },
      { x: 100.4, y: 50.6, w: -199.3, h: 0 },
    ]) {
      const box = marqueeBox(rect);
      expectNonNegativeBox(box);
      expectAlignedRing(box);
    }
    expect(marqueeBox({ x: 100, y: 50, w: 0, h: 0 })).toEqual({
      left: 99, top: 49, width: 1, height: 1, inset: { x: 100, y: 50, w: 0, h: 0 },
    });
  });

  test('straddles with an even width, the active-cell border', () => {
    expect(ACTIVE_CELL_WIDTH).toBe(2);
    const box = marqueeBox({ x: 100, y: 50, w: 200, h: 40 }, ACTIVE_CELL_WIDTH);
    expect(box).toEqual({ left: 99, top: 49, width: 202, height: 42, inset: { x: 100, y: 50, w: 200, h: 40 } });
  });

  test('uses a one-pixel marquee ring sized from the inset', () => {
    expect(MARQUEE_WIDTH).toBe(1);
    for (const rect of [
      { x: 0, y: 0, w: 90, h: 60 },
      { x: 12.4, y: 7.8, w: 45.6, h: 21.2 },
      { x: -5.5, y: -2.5, w: 10, h: 10 },
    ]) {
      const box = marqueeBox(rect);
      expect(box.width).toBe(box.inset.w + 1);
      expect(box.height).toBe(box.inset.h + 1);
      expectAlignedRing(box);
    }
  });
});

describe('marqueeFills', () => {
  test('returns the whole snapped range when the hole is null', () => {
    expect(marqueeFills({ x: 0, y: 0, w: 90, h: 60 }, null)).toEqual([{ x: 0, y: 0, w: 90, h: 60 }]);
    expect(marqueeFills({ x: 10.4, y: 20.6, w: 50.3, h: 10.7 }, null)).toEqual([
      marqueeBox({ x: 10.4, y: 20.6, w: 50.3, h: 10.7 }).inset,
    ]);
  });

  test('returns the whole snapped range when the hole does not intersect it', () => {
    const range = { x: 0, y: 0, w: 90, h: 60 };
    for (const hole of [
      { x: 90, y: 0, w: 20, h: 20 },
      { x: -30, y: 0, w: 30, h: 20 },
      { x: 10, y: 60, w: 10, h: 10 },
      { x: 10, y: -20, w: 10, h: 20 },
    ]) {
      expect(marqueeFills(range, hole)).toEqual([{ x: 0, y: 0, w: 90, h: 60 }]);
    }
  });

  test('returns no fills when the hole covers the entire range', () => {
    expect(marqueeFills({ x: 0, y: 0, w: 30, h: 20 }, { x: 0, y: 0, w: 30, h: 20 })).toEqual([]);
    expect(marqueeFills({ x: 10.4, y: 20.6, w: 30.2, h: 20.3 }, { x: 10.4, y: 20.6, w: 30.2, h: 20.3 })).toEqual([]);
  });

  test('splits a top-left hole into a bottom band and a right column', () => {
    const range = { x: 0, y: 0, w: 90, h: 60 };
    const hole = { x: 0, y: 0, w: 30, h: 20 };
    const strips = marqueeFills(range, hole);
    expect(strips).toEqual([
      { x: 0, y: 20, w: 90, h: 40 },
      { x: 30, y: 0, w: 60, h: 20 },
    ]);
    // Union plus hole tiles the range without overlap.
    expect(strips.reduce((sum, strip) => sum + strip.w * strip.h, 0) + hole.w * hole.h).toBe(range.w * range.h);
    for (const strip of strips) expect(overlaps(strip, hole)).toBe(false);
    for (let i = 0; i < strips.length; i++) {
      for (let j = i + 1; j < strips.length; j++) expect(overlaps(strips[i], strips[j])).toBe(false);
    }
  });

  test('splits a middle hole in a 3x3 range into four strips without gaps or overlap', () => {
    const range = { x: 0, y: 0, w: 90, h: 90 };
    const hole = { x: 30, y: 30, w: 30, h: 30 };
    const strips = marqueeFills(range, hole);
    expect(strips).toEqual([
      { x: 0, y: 0, w: 90, h: 30 },
      { x: 0, y: 60, w: 90, h: 30 },
      { x: 0, y: 30, w: 30, h: 30 },
      { x: 60, y: 30, w: 30, h: 30 },
    ]);
    expect(strips.reduce((sum, strip) => sum + strip.w * strip.h, 0)).toBe(range.w * range.h - hole.w * hole.h);
    for (const strip of strips) expect(overlaps(strip, hole)).toBe(false);
    for (let i = 0; i < strips.length; i++) {
      for (let j = i + 1; j < strips.length; j++) expect(overlaps(strips[i], strips[j])).toBe(false);
    }
  });

  test('clips a hole partially outside the range and never returns negative sizes', () => {
    const range = { x: 0, y: 0, w: 90, h: 60 };
    const hole = { x: -10, y: -10, w: 30, h: 30 };
    const clippedHole = intersect(range, marqueeBox(hole).inset);
    expect(clippedHole).toEqual({ x: 0, y: 0, w: 20, h: 20 });

    const strips = marqueeFills(range, hole);
    for (const strip of strips) {
      expect(strip.w).toBeGreaterThan(0);
      expect(strip.h).toBeGreaterThan(0);
      expect(strip.x).toBeGreaterThanOrEqual(range.x);
      expect(strip.y).toBeGreaterThanOrEqual(range.y);
      expect(strip.x + strip.w).toBeLessThanOrEqual(range.x + range.w);
      expect(strip.y + strip.h).toBeLessThanOrEqual(range.y + range.h);
      expect(overlaps(strip, clippedHole)).toBe(false);
    }
    expect(strips.reduce((sum, strip) => sum + strip.w * strip.h, 0)).toBe(range.w * range.h - clippedHole.w * clippedHole.h);
  });

  test('keeps every fractional strip inside the snapped range and clear of the snapped hole', () => {
    const cases = [
      { range: { x: 100.4, y: 50.6, w: 199.3, h: 39.7 }, hole: { x: 120.2, y: 60.1, w: 30.6, h: 10.4 } },
      { range: { x: 10.25, y: 20.75, w: 60.5, h: 40.25 }, hole: { x: 5.5, y: 15.5, w: 20.4, h: 20.2 } },
      { range: { x: -3.5, y: -7.25, w: 100.5, h: 50.75 }, hole: { x: 40.5, y: 10.5, w: 300.4, h: 5.2 } },
      { range: { x: 0.5, y: 0.5, w: 1.2, h: 0.8 }, hole: { x: 0.6, y: 0.4, w: 0.7, h: 0.9 } },
      { range: { x: 0.5, y: 0.5, w: 1, h: 1 }, hole: { x: 0.5, y: 0.5, w: 0.6, h: 0.6 } },
    ] as const;

    for (const { range, hole } of cases) {
      const outer = marqueeBox(range).inset;
      const clippedHole = intersect(outer, marqueeBox(hole).inset);
      const strips = marqueeFills(range, hole);
      for (const strip of strips) {
        expect(strip.w).toBeGreaterThan(0);
        expect(strip.h).toBeGreaterThan(0);
        expect(strip.x).toBeGreaterThanOrEqual(outer.x);
        expect(strip.y).toBeGreaterThanOrEqual(outer.y);
        expect(strip.x + strip.w).toBeLessThanOrEqual(outer.x + outer.w);
        expect(strip.y + strip.h).toBeLessThanOrEqual(outer.y + outer.h);
        expect(overlaps(strip, clippedHole)).toBe(false);
      }
      expect(strips.reduce((sum, strip) => sum + strip.w * strip.h, 0) + clippedHole.w * clippedHole.h).toBe(outer.w * outer.h);
    }
  });
});
