import { describe, expect, it } from 'bun:test';
import {
  chartResizeEdges,
  chartResizeRect,
  chartResizeSpace,
  type ChartResizeViewport,
} from './chartResize';

const viewport: ChartResizeViewport = { x: 200, y: 120, frozenRows: 1, frozenCols: 1 };

describe('chart resize geometry', () => {
  it('maps a handle to the edges it actually moves', () => {
    expect(chartResizeEdges('se', 12, 7)).toEqual({ left: 0, top: 0, right: 12, bottom: 7 });
    expect(chartResizeEdges('nw', 12, 7)).toEqual({ left: 12, top: 7, right: 0, bottom: 0 });
    expect(chartResizeEdges('n', 12, 7)).toEqual({ left: 0, top: 7, right: 0, bottom: 0 });
  });

  it('keeps untouched edges still when the chart is partly scrolled off', () => {
    // the painted rect is viewport-local: x=-50 at scroll x=200 is sheet x=150.
    const rect = { x: -50, y: 10, w: 100, h: 50 };
    // an unpinned pane clips from the frozen edge, so x>0 here.
    const space = chartResizeSpace({ clip: { x: 210, y: 130, w: 10, h: 10 } }, viewport);
    expect(chartResizeRect(rect, 'se', 20, 10, space)).toEqual({ x: -50, y: 10, w: 120, h: 60 });
    expect(chartResizeRect(rect, 'nw', 20, 10, space)).toEqual({ x: -30, y: 20, w: 80, h: 40 });
  });

  it('clamps a west edge at the sheet origin, not at the viewport edge', () => {
    // sheet x starts at 30; pulling the west edge 100px left stops at 0,
    // which is -200 in viewport coordinates, not 0.
    const rect = { x: -170, y: 10, w: 200, h: 50 };
    const space = chartResizeSpace({ clip: { x: 210, y: 130, w: 10, h: 10 } }, viewport);
    expect(chartResizeRect(rect, 'w', -100, 0, space)).toEqual({ x: -200, y: 10, w: 230, h: 50 });
  });

  it('clamps right and bottom at the worksheet far edge like the engine', () => {
    const space = chartResizeSpace({ clip: { x: 10, y: 10, w: 10, h: 10 } }, {
      x: 0, y: 0, frozenRows: 0, frozenCols: 0, sheetWidth: 1000, sheetHeight: 800,
    });
    const rect = { x: 900, y: 700, w: 80, h: 80 };
    expect(chartResizeRect(rect, 'se', 100, 100, space)).toEqual({ x: 900, y: 700, w: 100, h: 100 });
  });

  it('keeps the engine min-size order when a west/north edge hits the origin', () => {
    const space = chartResizeSpace({ clip: { x: 10, y: 10, w: 10, h: 10 } }, {
      x: 0, y: 0, frozenRows: 0, frozenCols: 0, sheetWidth: 1000, sheetHeight: 800,
    });
    // both dragged edges clamp to 0 and the untouched far edges give way to
    // the 1px minimum, so the frame starts at the origin instead of the
    // untouched edge staying at 0.5.
    expect(chartResizeRect({ x: 0, y: 0, w: 0.5, h: 0.5 }, 'nw', 10, 10, space))
      .toEqual({ x: 0, y: 0, w: 1, h: 1 });
  });

  it('treats a chart pinned in a frozen pane as already in sheet space', () => {
    const pinned = chartResizeSpace({ clip: { x: 0, y: 0, w: 40, h: 30 } }, viewport);
    expect(pinned).toEqual({ originX: 0, originY: 0, width: Infinity, height: Infinity });
    const scrolled = chartResizeSpace({ clip: { x: 210, y: 130, w: 40, h: 30 } }, viewport);
    expect(scrolled).toEqual({ originX: 200, originY: 120, width: Infinity, height: Infinity });
  });
});
