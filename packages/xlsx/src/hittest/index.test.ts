import { describe, expect, it } from 'bun:test';
import { cellAtPoint, cellRect, rangeRect } from './index';
import type { GridMeta } from '../display-list/types';

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
