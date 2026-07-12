import { describe, expect, it } from 'bun:test';
import {
  cellCount,
  extendTo,
  moveFocus,
  normalizeRange,
  rangeContains,
  selectionAt,
} from './index';
import type { Selection, SelectionLimits } from './index';

const limits: SelectionLimits = { rows: 100, cols: 26, rowsPerPage: 20 };

describe('selectionAt', () => {
  it('collapses anchor and focus onto one cell', () => {
    const sel = selectionAt({ row: 3, col: 4 });
    expect(sel).toEqual({ anchor: { row: 3, col: 4 }, focus: { row: 3, col: 4 } });
  });

  it('copies the address so later mutation cannot alias it', () => {
    const addr = { row: 1, col: 1 };
    const sel = selectionAt(addr);
    addr.row = 9;
    expect(sel.anchor.row).toBe(1);
  });
});

describe('normalizeRange', () => {
  it('orders an inverted anchor/focus into a top-left rectangle', () => {
    const sel: Selection = { anchor: { row: 5, col: 7 }, focus: { row: 2, col: 3 } };
    expect(normalizeRange(sel)).toEqual({ top: 2, left: 3, bottom: 5, right: 7 });
  });

  it('is a point for a collapsed selection', () => {
    expect(normalizeRange(selectionAt({ row: 2, col: 2 }))).toEqual({
      top: 2,
      left: 2,
      bottom: 2,
      right: 2,
    });
  });
});

describe('rangeContains / cellCount', () => {
  const sel: Selection = { anchor: { row: 1, col: 1 }, focus: { row: 3, col: 4 } };

  it('detects membership inside the normalized range', () => {
    expect(rangeContains(sel, { row: 2, col: 3 })).toBe(true);
    expect(rangeContains(sel, { row: 0, col: 3 })).toBe(false);
    expect(rangeContains(sel, { row: 2, col: 5 })).toBe(false);
  });

  it('counts the covered cells', () => {
    expect(cellCount(sel)).toBe(3 * 4);
    expect(cellCount(selectionAt({ row: 0, col: 0 }))).toBe(1);
  });
});

describe('extendTo', () => {
  it('holds the anchor and moves the focus', () => {
    const sel = selectionAt({ row: 2, col: 2 });
    expect(extendTo(sel, { row: 5, col: 6 })).toEqual({
      anchor: { row: 2, col: 2 },
      focus: { row: 5, col: 6 },
    });
  });

  it('clamps the focus into the grid when limits are given', () => {
    const sel = selectionAt({ row: 2, col: 2 });
    const out = extendTo(sel, { row: 999, col: 999 }, limits);
    expect(out.focus).toEqual({ row: 99, col: 25 });
  });
});

describe('moveFocus', () => {
  const sel: Selection = { anchor: { row: 4, col: 4 }, focus: { row: 4, col: 4 } };

  it('collapses onto the new cell when not extending', () => {
    const out = moveFocus(sel, 'down', { limits });
    expect(out).toEqual({ anchor: { row: 5, col: 4 }, focus: { row: 5, col: 4 } });
  });

  it('holds the anchor when extending', () => {
    const out = moveFocus(sel, 'right', { extend: true, limits });
    expect(out).toEqual({ anchor: { row: 4, col: 4 }, focus: { row: 4, col: 5 } });
  });

  it('clamps at the grid edge instead of overflowing', () => {
    const atEdge = selectionAt({ row: 0, col: 0 });
    expect(moveFocus(atEdge, 'up', { limits }).focus).toEqual({ row: 0, col: 0 });
    expect(moveFocus(atEdge, 'left', { limits }).focus).toEqual({ row: 0, col: 0 });
  });

  it('steps by more than one cell for paging', () => {
    const out = moveFocus(sel, 'down', { step: 20, limits });
    expect(out.focus).toEqual({ row: 24, col: 4 });
  });
});
