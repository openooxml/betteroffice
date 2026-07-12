import { describe, expect, it } from 'bun:test';
import { clampScroll, uniformOffsets, visibleRange, visibleCells } from './index';
import type { ViewportState } from './index';

describe('visibleRange', () => {
  const offsets = uniformOffsets(20, 80); // 20 tracks, 80px each, total 1600

  it('returns leading tracks at scroll 0', () => {
    expect(visibleRange(offsets, 0, 0, 320)).toEqual({ first: 0, last: 4 });
  });

  it('clips leading tracks when scrolled', () => {
    // scroll 250 lands inside track 3 (240..320); window 250..570 covers 3..7.
    expect(visibleRange(offsets, 0, 250, 320)).toEqual({ first: 3, last: 7 });
  });

  it('excludes frozen tracks and shortens the window by their width', () => {
    // 2 frozen (160px) pinned; body window starts at 160 + scroll.
    const r = visibleRange(offsets, 2, 200, 320);
    expect(r.first).toBeGreaterThanOrEqual(2);
    expect(r.last).toBeGreaterThanOrEqual(r.first);
  });

  it('is empty for a zero-track axis', () => {
    expect(visibleRange([0], 0, 0, 100)).toEqual({ first: 0, last: -1 });
  });
});

describe('clampScroll', () => {
  const colOffsets = uniformOffsets(10, 80); // total 800
  const rowOffsets = uniformOffsets(10, 20); // total 200
  const base: ViewportState = {
    scrollX: 0,
    scrollY: 0,
    width: 400,
    height: 100,
    dpr: 1,
    frozenRows: 0,
    frozenCols: 0,
  };

  it('returns the same object when already in range', () => {
    expect(clampScroll(base, colOffsets, rowOffsets)).toBe(base);
  });

  it('clamps overscroll to the max offset', () => {
    const state = { ...base, scrollX: 10_000, scrollY: 10_000 };
    const clamped = clampScroll(state, colOffsets, rowOffsets);
    expect(clamped.scrollX).toBe(400); // 800 content - 400 viewport
    expect(clamped.scrollY).toBe(100); // 200 content - 100 viewport
  });

  it('clamps negative scroll to zero', () => {
    const clamped = clampScroll({ ...base, scrollX: -50 }, colOffsets, rowOffsets);
    expect(clamped.scrollX).toBe(0);
  });
});

describe('visibleCells', () => {
  it('computes both axes together', () => {
    const state: ViewportState = {
      scrollX: 0,
      scrollY: 0,
      width: 320,
      height: 60,
      dpr: 1,
      frozenRows: 0,
      frozenCols: 0,
    };
    const result = visibleCells(state, uniformOffsets(20, 80), uniformOffsets(20, 20));
    expect(result.cols).toEqual({ first: 0, last: 4 });
    expect(result.rows).toEqual({ first: 0, last: 3 });
  });
});
