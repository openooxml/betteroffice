import { describe, expect, it } from 'bun:test';
import { cellAtPoint, cellRect, chartRegionAtPoint, rangeRect } from './index';
import type { ChartRegion, GridMeta, Rect } from '../display-list/types';

// a 3-col x 3-row window starting at sheet cell (10, 5). columns are 80px wide,
// rows 20px tall, both in viewport-local px from the frame origin.
const grid: GridMeta = {
  startRow: 10,
  startCol: 5,
  colOffsets: [0, 80, 160, 240],
  rowOffsets: [0, 20, 40, 60],
};

describe('cellAtPoint', () => {
  it('maps a point to the enclosing sheet cell', () => {
    expect(cellAtPoint(grid, 0, 0)).toEqual({ row: 10, col: 5 });
    expect(cellAtPoint(grid, 90, 25)).toEqual({ row: 11, col: 6 });
    expect(cellAtPoint(grid, 239, 59)).toEqual({ row: 12, col: 7 });
  });

  it('places a point on a boundary in the trailing track', () => {
    // x=80 is the shared edge of cols 0 and 1; the leading edge wins col 1.
    expect(cellAtPoint(grid, 80, 20)).toEqual({ row: 11, col: 6 });
  });

  it('returns null outside the visible grid', () => {
    expect(cellAtPoint(grid, -1, 10)).toBeNull();
    expect(cellAtPoint(grid, 240, 10)).toBeNull();
    expect(cellAtPoint(grid, 10, 60)).toBeNull();
  });

  it('returns null when there is no grid metadata', () => {
    expect(cellAtPoint(undefined, 10, 10)).toBeNull();
  });
});

describe('cellRect', () => {
  it('returns the pixel box of a visible cell', () => {
    expect(cellRect(grid, 10, 5)).toEqual({ x: 0, y: 0, w: 80, h: 20 });
    expect(cellRect(grid, 12, 7)).toEqual({ x: 160, y: 40, w: 80, h: 20 });
  });

  it('returns null for a cell outside the window', () => {
    expect(cellRect(grid, 9, 5)).toBeNull();
    expect(cellRect(grid, 13, 5)).toBeNull();
    expect(cellRect(grid, 10, 8)).toBeNull();
    expect(cellRect(undefined, 10, 5)).toBeNull();
  });
});

describe('rangeRect', () => {
  it('unions the cells of a fully visible range', () => {
    const rect = rangeRect(grid, { top: 10, left: 5, bottom: 11, right: 6 });
    expect(rect).toEqual({ x: 0, y: 0, w: 160, h: 40 });
  });

  it('clips a range that overflows the visible window', () => {
    // range extends past the last visible row/col; result is clamped to it.
    const rect = rangeRect(grid, { top: 11, left: 6, bottom: 99, right: 99 });
    expect(rect).toEqual({ x: 80, y: 20, w: 160, h: 40 });
  });

  it('clips a range that starts before the window', () => {
    const rect = rangeRect(grid, { top: 0, left: 0, bottom: 10, right: 5 });
    expect(rect).toEqual({ x: 0, y: 0, w: 80, h: 20 });
  });

  it('returns null for a range entirely outside the window', () => {
    expect(rangeRect(grid, { top: 0, left: 0, bottom: 9, right: 4 })).toBeNull();
    expect(rangeRect(grid, { top: 13, left: 8, bottom: 20, right: 20 })).toBeNull();
    expect(rangeRect(undefined, { top: 10, left: 5, bottom: 11, right: 6 })).toBeNull();
  });
});

describe('non-uniform tracks', () => {
  // widths 50, 100, 30; heights 10, 40 — hit-test must not assume uniformity.
  const g: GridMeta = {
    startRow: 0,
    startCol: 0,
    colOffsets: [0, 50, 150, 180],
    rowOffsets: [0, 10, 50],
  };

  it('finds the right track under variable widths', () => {
    expect(cellAtPoint(g, 49, 5)).toEqual({ row: 0, col: 0 });
    expect(cellAtPoint(g, 50, 5)).toEqual({ row: 0, col: 1 });
    expect(cellAtPoint(g, 149, 45)).toEqual({ row: 1, col: 1 });
    expect(cellAtPoint(g, 150, 45)).toEqual({ row: 1, col: 2 });
  });
});

describe('frozen pane address maps', () => {
  const g: GridMeta = {
    startRow: 0,
    startCol: 0,
    rowIndices: [0, 4, 5],
    colIndices: [0, 3, 4],
    colOffsets: [0, 80, 160, 240],
    rowOffsets: [0, 20, 40, 60],
  };

  it('maps pinned and scrolled tracks to their absolute addresses', () => {
    expect(cellAtPoint(g, 10, 10)).toEqual({ row: 0, col: 0 });
    expect(cellAtPoint(g, 90, 25)).toEqual({ row: 4, col: 3 });
    expect(cellRect(g, 5, 4)).toEqual({ x: 160, y: 40, w: 80, h: 20 });
    expect(cellRect(g, 2, 2)).toBeNull();
  });

  it('unions only visible tracks from a range crossing the pane gap', () => {
    expect(rangeRect(g, { top: 0, left: 0, bottom: 4, right: 3 })).toEqual({
      x: 0,
      y: 0,
      w: 160,
      h: 40,
    });
    expect(rangeRect(g, { top: 1, left: 1, bottom: 3, right: 2 })).toBeNull();
  });
});

describe('chartRegionAtPoint', () => {
  const region = (id: string, clip: Rect, rect: Rect = clip): ChartRegion => ({
    id,
    label: '',
    rect,
    clip,
    movable: true,
  });

  it('bounds the hit by the clipped region, half-open like the engine', () => {
    const charts = [region('a', { x: 10, y: 20, w: 100, h: 50 })];
    expect(chartRegionAtPoint(charts, 10, 20)?.id).toBe('a');
    expect(chartRegionAtPoint(charts, 109.9, 69.9)?.id).toBe('a');
    expect(chartRegionAtPoint(charts, 110, 40)).toBeNull();
    expect(chartRegionAtPoint(charts, 9.9, 40)).toBeNull();
    expect(chartRegionAtPoint(charts, Number.NaN, 40)).toBeNull();
    expect(chartRegionAtPoint(undefined, 10, 20)).toBeNull();
  });

  // a chart the renderer could not draw still occupies its space, so it stays
  // an object the pointer can reach.
  it('answers for a chart that degraded to a placeholder', () => {
    const box = { x: 0, y: 0, w: 50, h: 40 };
    const charts = [{ ...region('undrawable', box), placeholder: true }];
    expect(chartRegionAtPoint(charts, 25, 20)?.id).toBe('undrawable');
  });

  it('resolves overlapping regions to the last painted', () => {
    const charts = [
      region('under', { x: 0, y: 0, w: 100, h: 100 }),
      region('over', { x: 50, y: 50, w: 100, h: 100 }),
    ];
    expect(chartRegionAtPoint(charts, 60, 60)?.id).toBe('over');
    expect(chartRegionAtPoint(charts, 10, 10)?.id).toBe('under');
  });

  it('hits the full rect only where the clip allows', () => {
    const charts = [
      region('clipped', { x: 80, y: 0, w: 120, h: 100 }, { x: 0, y: 0, w: 200, h: 100 }),
    ];
    expect(chartRegionAtPoint(charts, 40, 50)).toBeNull();
    expect(chartRegionAtPoint(charts, 100, 50)?.id).toBe('clipped');
  });

  // the far edge is where f32 and f64 disagree. these widths make the f32 sum
  // round DOWN, so the engine's right edge sits below the f64 one and a point
  // exactly on it is outside for the engine but inside for a naive predicate
  // that adds in f64 — the reachable-coordinate hazard, in one assertion.
  it('computes the far edge at f32, as the engine does', () => {
    const x = Math.fround(2.5899999141693115);
    const w = Math.fround(3.7699999809265137);
    const engineRight = Math.fround(x + w);
    expect(engineRight).toBeLessThan(x + w);

    const charts = [region('narrow', { x, y: 0, w, h: 10 })];
    expect(chartRegionAtPoint(charts, engineRight, 5)).toBeNull();
    expect(chartRegionAtPoint(charts, Math.fround(engineRight - 1e-6), 5)?.id).toBe('narrow');
  });

  // serde writes an f32 as the shortest decimal that round-trips *as f32*, and
  // JS parses that to a nearby but different f64 — so the width off the wire is
  // not the width the engine added. 3.32 parses 4.8e-7 away from its own f32.
  it('rounds the rect fields off the wire before adding them', () => {
    const wire = 3.32;
    const edge = Math.fround(1.0399999618530273 + Math.fround(wire));
    expect(edge).not.toBe(Math.fround(1.0399999618530273 + wire));

    const charts = [
      region('wire', { x: 1.0399999618530273, y: 1.0399999618530273, w: wire, h: wire }),
    ];
    expect(chartRegionAtPoint(charts, 2, 2)?.id).toBe('wire');
    expect(chartRegionAtPoint(charts, edge, 2)).toBeNull();
    expect(chartRegionAtPoint(charts, 2, edge)).toBeNull();
  });

  // the point arrives as an f64 from client px over zoom; the wasm boundary
  // narrows it to f32, so a point just below an edge rounds onto it.
  it('narrows the point to f32 the way the wasm boundary does', () => {
    const left = Math.fround(2.59);
    const charts = [region('edge', { x: left, y: 0, w: 10, h: 10 })];
    const justBelow = left - 1e-9;
    expect(justBelow).toBeLessThan(left);
    expect(Math.fround(justBelow)).toBe(left);
    expect(chartRegionAtPoint(charts, justBelow, 5)?.id).toBe('edge');
  });
});
