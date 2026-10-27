import { describe, expect, test } from 'bun:test';
import type { DisplayList, GridMeta } from '@betteroffice/xlsx';
import { createPeerlessPluginGeometry } from './createPeerlessPluginDocument';
import {
  createPluginGeometry,
  frozenGridExtent,
  gridRectToOverlay,
  pluginLayout,
  visibleGridRange,
  type PaintedGrid,
} from './geometry';

function element(left: number, top: number): HTMLElement {
  return {
    clientLeft: 0,
    clientTop: 0,
    scrollLeft: 0,
    scrollTop: 0,
    getBoundingClientRect: () => ({ left, top }),
  } as unknown as HTMLElement;
}

/** Two frozen rows and one frozen column over a body scrolled to row 10 and column 4. */
function frozenFrame(): DisplayList {
  return {
    width: 300,
    height: 100,
    commands: [],
    grid: {
      startRow: 0,
      startCol: 0,
      rowIndices: [0, 1, 10, 11, 12],
      colIndices: [0, 4, 5, 6],
      rowOffsets: [0, 20, 40, 60, 80, 100],
      colOffsets: [0, 64, 128, 192, 256],
    },
  };
}

function painted(overrides: Partial<PaintedGrid> = {}): PaintedGrid {
  return {
    frame: frozenFrame(),
    zoom: 2,
    version: 'v1',
    sheetId: 'sheet:0',
    viewport: { x: 192, y: 160, width: 300, height: 100 },
    ...overrides,
  };
}

describe('grid track geometry', () => {
  test('frozen extents count visible frozen tracks after hidden tracks are compacted', () => {
    const grid: GridMeta = {
      ...frozenFrame().grid!,
      rowIndices: [0, 4, 8, 10, 12],
      colIndices: [0, 6, 11, 13],
    };
    expect(frozenGridExtent(grid, 8, 10, { width: 300, height: 100 })).toEqual({
      width: 128,
      height: 40,
    });
  });

  test('frozen extents preserve dense tracks and missing-frame fallbacks', () => {
    const grid: GridMeta = {
      ...frozenFrame().grid!,
      rowIndices: undefined,
      colIndices: undefined,
    };
    expect(frozenGridExtent(grid, 2, 1)).toEqual({ width: 64, height: 40 });
    expect(frozenGridExtent(grid, 0, 0)).toEqual({ width: 0, height: 0 });
    expect(frozenGridExtent(undefined, 2, 1)).toEqual({ width: 0, height: 0 });
    expect(frozenGridExtent(undefined, 2, 1, { width: 300, height: 100 })).toEqual({
      width: 300,
      height: 100,
    });
  });

  test('frozen extents handle hidden prefixes without explicit track indices', () => {
    const grid: GridMeta = {
      ...frozenFrame().grid!,
      startRow: 2,
      startCol: 1,
      rowIndices: undefined,
      colIndices: undefined,
    };
    expect(frozenGridExtent(grid, 4, 2)).toEqual({ width: 64, height: 40 });
    expect(frozenGridExtent(grid, 2, 1)).toEqual({ width: 0, height: 0 });
  });

  test('merged-range queries reach the last compacted body tracks', () => {
    const grid: GridMeta = {
      ...frozenFrame().grid!,
      rowIndices: [0, 4, 8, 10, 12],
      colIndices: [0, 6, 11, 13],
    };
    expect(visibleGridRange(grid)).toEqual({ top: 0, left: 0, bottom: 12, right: 13 });
  });

  test('merged-range queries preserve dense ranges and skip empty grids', () => {
    const grid: GridMeta = {
      ...frozenFrame().grid!,
      startRow: 10,
      startCol: 4,
      rowIndices: undefined,
      colIndices: undefined,
    };
    expect(visibleGridRange(grid)).toEqual({ top: 10, left: 4, bottom: 14, right: 7 });
    expect(visibleGridRange({ ...grid, rowOffsets: [] })).toBeNull();
    expect(visibleGridRange({ ...grid, colOffsets: [0] })).toBeNull();
  });
});

describe('plugin layouts', () => {
  test('exist only while the painted frame shows the current version', () => {
    const frame = painted();
    const layout = pluginLayout(frame, 'v1');
    expect(layout).toMatchObject({
      version: 'v1',
      sheetId: 'sheet:0',
      zoom: 2,
      viewport: { x: 192, y: 160, width: 300, height: 100 },
    });
    expect(pluginLayout(frame, 'v1')?.id).toBe(layout!.id);
    expect(pluginLayout(painted(), 'v1')?.id).not.toBe(layout!.id);
    expect(pluginLayout(frame, 'v2')).toBeNull();
    expect(pluginLayout(null, 'v1')).toBeNull();
    expect(pluginLayout(frame, null)).toBeNull();
  });
});

describe('overlay coordinates', () => {
  test('retained geometry refuses a new acknowledged version or painted frame', () => {
    const frame = painted();
    let projection = { version: 'v1', painted: frame };
    const geometry = createPeerlessPluginGeometry(pluginLayout(frame, 'v1')!, frame,
      element(0, 0), element(0, 0), () => projection);
    const cell = { sheetId: 'sheet:0', row: 0, col: 0 };
    const range = { sheetId: 'sheet:0', range: { top: 0, left: 0, bottom: 1, right: 0 } };
    expect(geometry.getCellRect(cell)).not.toBeNull();
    projection = { version: 'v2', painted: frame };
    expect(geometry.getCellRect(cell)).toBeNull();
    expect(geometry.getRangeRect(range)).toBeNull();
    projection = { version: 'v1', painted: painted() };
    expect(geometry.getCellRect(cell)).toBeNull();
    expect(geometry.getRangeRect(range)).toBeNull();
  });

  test('zoom viewport-local grid rects, clip them to the frame and offset them by the canvas', () => {
    const grid = { frame: { width: 300, height: 100 } as DisplayList, zoom: 1.5 };
    expect(gridRectToOverlay(grid, { x: 64, y: 20, w: 64, h: 20 }, { x: 3, y: 4 })).toEqual({
      x: 3 + 96,
      y: 4 + 30,
      width: 96,
      height: 30,
    });
    expect(gridRectToOverlay(grid, { x: -10, y: 90, w: 30, h: 40 }, { x: 0, y: 0 })).toEqual({
      x: 0,
      y: 135,
      width: 30,
      height: 15,
    });
    expect(gridRectToOverlay(grid, { x: 300, y: 0, w: 10, h: 10 }, { x: 0, y: 0 })).toBeNull();
  });

  test('follow frozen tracks and the scrolled body, and refuse other sheets and old frames', () => {
    const frame = painted();
    const layout = pluginLayout(frame, 'v1')!;
    let current = true;
    const geometry = createPluginGeometry(
      layout,
      frame,
      element(110, 60),
      element(100, 50),
      () => current
    );
    const cell = (row: number, col: number) =>
      geometry.getCellRect({ sheetId: 'sheet:0', row, col });
    expect(cell(0, 0)).toEqual({ x: 10, y: 10, width: 128, height: 40 });
    expect(cell(10, 4)).toEqual({ x: 10 + 128, y: 10 + 80, width: 128, height: 40 });
    expect(cell(5, 2)).toBeNull();
    expect(cell(1, 6)).toEqual({ x: 10 + 384, y: 10 + 40, width: 128, height: 40 });
    expect(
      geometry.getRangeRect({
        sheetId: 'sheet:0',
        range: { top: 1, left: 0, bottom: 11, right: 4 },
      })
    ).toEqual({ x: 10, y: 10 + 40, width: 256, height: 120 });
    expect(
      geometry.getRangeRect({ sheetId: 'sheet:0', range: { top: 3, left: 1, bottom: 2, right: 1 } })
    ).toBeNull();
    expect(geometry.getCellRect({ sheetId: 'sheet:1', row: 0, col: 0 })).toBeNull();
    expect(geometry.getCellRect({ sheetId: 'sheet:0', row: -1, col: 0 })).toBeNull();
    expect(geometry.getCellRect({ sheetId: 'sheet:0', row: 0.5, col: 0 })).toBeNull();
    expect(geometry.getPositionAtPoint).toBeNull();

    current = false;
    expect(cell(0, 0)).toBeNull();
    expect(
      geometry.getRangeRect({ sheetId: 'sheet:0', range: { top: 0, left: 0, bottom: 1, right: 1 } })
    ).toBeNull();
  });
});
