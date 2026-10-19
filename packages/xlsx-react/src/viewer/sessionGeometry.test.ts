import { describe, expect, it } from 'bun:test';
import type { DisplayList, WorkbookSheetView } from '@betteroffice/xlsx';
import { cellInViewport, deriveLimits, extendScrollExtent, laidOutCellRect, nearestCellScroll } from './sessionGeometry';

describe('sheet navigation geometry', () => {
  it('keeps sheet bounds independent of rendered size, track widths, hidden spans and zoom', () => {
    for (const height of [48, 240_024]) {
      for (const rowHeight of [12, 24, 48]) {
        const frame: DisplayList = { width: 67_488, height, commands: [], grid: {
          startRow: 9_998, startCol: 700,
          rowOffsets: [0, rowHeight, rowHeight * 2], colOffsets: [0, 80, 240],
          hiddenRowSpans: [[10_000, 900_000]], hiddenColSpans: [[702, 10_000]],
        } };
        for (const zoom of [1, 1.5]) {
          expect(deriveLimits(frame, 600 / zoom)).toEqual({
            rows: 1_048_576, cols: 16_384, rowsPerPage: Math.floor(600 / zoom / rowHeight),
          });
        }
      }
    }
    expect(deriveLimits(null, 0)).toEqual({ rows: 1_048_576, cols: 16_384, rowsPerPage: 1 });
  });

  it('extends the scroller before revealing distant cells and preserves the extent when zooming', () => {
    const content = { style: { width: '', height: '' } };
    const scroll = { firstElementChild: content } as unknown as HTMLElement;
    const initial = { sheet: 2, contentWidth: 4_000, contentHeight: 6_000 };
    const far = { x: 701 * 96, y: 9_999 * 24, w: 96, h: 24 };
    const grown = extendScrollExtent(scroll, initial, far, 1);
    expect(grown).toEqual({ sheet: 2, contentWidth: 67_392, contentHeight: 240_000 });
    expect(content.style).toEqual({ width: '67392px', height: '240000px' });
    expect(extendScrollExtent(scroll, grown, { x: 0, y: 0, w: 96, h: 24 }, 1.5)).toEqual(grown);
    expect(content.style).toEqual({ width: '101088px', height: '360000px' });
    expect(initial).toEqual({ sheet: 2, contentWidth: 4_000, contentHeight: 6_000 });
  });

  it('preserves fully visible cells and reveals only the nearest edge, including after zoom', () => {
    const view = { frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0 } as WorkbookSheetView;
    const viewport = { x: 100, y: 200, width: 800, height: 600 };
    const cell = { row: 15, col: 4 };
    const rect = { x: 384, y: 360, w: 96, h: 24 };
    expect(nearestCellScroll(rect, cell, view, viewport)).toEqual({ x: 100, y: 200 });
    expect(nearestCellScroll({ ...rect, x: 850, y: 790 }, cell, view, viewport))
      .toEqual({ x: 146, y: 214 });
    expect(nearestCellScroll({ ...rect, x: 0, y: 0 }, cell, view, viewport)).toEqual({ x: 0, y: 0 });
    const far = { x: 701 * 96, y: 9_999 * 24, w: 96, h: 24 };
    const first = nearestCellScroll(far, cell, view, { x: 0, y: 0, width: 800, height: 600 });
    expect(first).toEqual({ x: 67_392 - 800, y: 240_000 - 600 });
    expect(nearestCellScroll(far, cell, view, {
      x: first.x / 1.5, y: first.y / 1.5, width: 800 / 1.5, height: 600 / 1.5,
    })).toEqual({ x: 67_392 - 800 / 1.5, y: 240_000 - 600 / 1.5 });
  });

  it('reuses laid-out sheet coordinates and respects frozen panes', () => {
    const view = { frozenRows: 1, frozenCols: 1, frozenWidth: 96, frozenHeight: 24 } as WorkbookSheetView;
    const viewport = { x: 1_000, y: 2_000, width: 800, height: 600 };
    const frame: DisplayList = { width: 800, height: 600, commands: [], grid: {
      startRow: 0, startCol: 0, rowIndices: [0, 84], colIndices: [0, 11],
      rowOffsets: [0, 24, 48], colOffsets: [0, 96, 192],
    } };
    expect(laidOutCellRect(frame, viewport, view, { row: 0, col: 0 }))
      .toEqual({ x: 0, y: 0, w: 96, h: 24 });
    expect(laidOutCellRect(frame, viewport, view, { row: 84, col: 11 }))
      .toEqual({ x: 1_096, y: 2_024, w: 96, h: 24 });
    expect(laidOutCellRect(frame, viewport, view, { row: 85, col: 11 })).toBeNull();
    expect(nearestCellScroll({ x: 0, y: 0, w: 96, h: 24 }, { row: 0, col: 0 }, view, viewport))
      .toEqual({ x: 1_000, y: 2_000 });
    expect(nearestCellScroll({ x: 1_750, y: 2_590, w: 96, h: 24 }, { row: 90, col: 18 }, view, viewport))
      .toEqual({ x: 1_046, y: 2_014 });
  });

  it('keeps the full outline inside the viewport at fractional zoom and scroll positions', () => {
    const view = { frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0 } as WorkbookSheetView;
    const cell = { row: 40, col: 20 };
    const rect = { x: 1_920.25, y: 960.125, w: 96, h: 24 };
    for (const zoom of [0.75, 1, 1.25, 1.5]) {
      const viewport = { x: 0, y: 0, width: 800 / zoom, height: 600 / zoom };
      const target = nearestCellScroll(rect, cell, view, viewport, 2 / zoom);
      const right = (rect.x + rect.w) * zoom - Math.round(target.x * zoom);
      const bottom = (rect.y + rect.h) * zoom - Math.round(target.y * zoom);
      expect(right).toBeLessThanOrEqual(798.5);
      expect(bottom).toBeLessThanOrEqual(598.5);
      const back = nearestCellScroll(rect, cell, view, { ...viewport, x: rect.x + 10, y: rect.y + 10 }, 2 / zoom);
      expect(rect.x * zoom - Math.round(back.x * zoom)).toBeGreaterThanOrEqual(1.5);
      expect(rect.y * zoom - Math.round(back.y * zoom)).toBeGreaterThanOrEqual(1.5);
      expect(nearestCellScroll({ x: 0, y: 0, w: 96, h: 24 }, { row: 0, col: 0 }, view, viewport, 2 / zoom))
        .toEqual({ x: 0, y: 0 });
    }
  });

  it('checks the previous zoom viewport before revealing a selected cell again', () => {
    const scroll = { scrollLeft: 67_000, scrollTop: 239_500, clientWidth: 800, clientHeight: 600 } as HTMLElement;
    const far = { x: 701 * 96, y: 9_999 * 24, w: 96, h: 24 };
    expect(cellInViewport(scroll, far, 1)).toBe(true);
    expect(cellInViewport(scroll, { x: 0, y: 0, w: 96, h: 24 }, 1)).toBe(false);
    expect(cellInViewport(scroll, far, 1.5)).toBe(false);
  });
});
