import { cellRect, rangeRect } from '@betteroffice/xlsx';
import type { CellRange, DisplayList, Viewport } from '@betteroffice/xlsx';
import type { XlsxPluginGeometry, XlsxPluginLayout, XlsxPluginRect } from './types';

/** A grid frame the canvas painted, with what it shows. */
export interface PaintedGrid {
  frame: DisplayList;
  zoom: number;
  /** The workbook version the frame was built from. */
  version: string;
  sheetId: string;
  /** The requested viewport, in unzoomed sheet pixels. */
  viewport: Viewport;
}

const layoutIds = new WeakMap<PaintedGrid, string>();
let nextLayoutId = 0;

function layoutIdOf(painted: PaintedGrid): string {
  let id = layoutIds.get(painted);
  if (!id) {
    nextLayoutId += 1;
    id = `grid-layout-${nextLayoutId}`;
    layoutIds.set(painted, id);
  }
  return id;
}

/** The layout a painted frame shows, or null unless its pixels show `version`. */
export function pluginLayout(
  painted: PaintedGrid | null,
  version: string | null
): XlsxPluginLayout | null {
  if (!painted || version === null || painted.version !== version) return null;
  const { x, y, width, height } = painted.viewport;
  return {
    id: layoutIdOf(painted),
    version,
    sheetId: painted.sheetId,
    zoom: painted.zoom,
    viewport: { x, y, width, height },
  };
}

function isIndex(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0;
}

/**
 * Converts a viewport-local grid rectangle into overlay pixels: clipped to the painted frame,
 * zoomed, and offset by where the canvas sits in `layer`.
 */
export function gridRectToOverlay(
  painted: Pick<PaintedGrid, 'frame' | 'zoom'>,
  rect: { x: number; y: number; w: number; h: number },
  origin: { x: number; y: number }
): XlsxPluginRect | null {
  const left = Math.max(0, rect.x);
  const top = Math.max(0, rect.y);
  const right = Math.min(painted.frame.width, rect.x + rect.w);
  const bottom = Math.min(painted.frame.height, rect.y + rect.h);
  if (!(right > left && bottom > top)) return null;
  return {
    x: origin.x + left * painted.zoom,
    y: origin.y + top * painted.zoom,
    width: (right - left) * painted.zoom,
    height: (bottom - top) * painted.zoom,
  };
}

/** Where the canvas's top left corner sits in the overlay layer's coordinates. */
export function canvasOrigin(canvas: Element, layer: HTMLElement): { x: number; y: number } {
  const canvasRect = canvas.getBoundingClientRect();
  const layerRect = layer.getBoundingClientRect();
  return {
    x: canvasRect.left - layerRect.left - layer.clientLeft + layer.scrollLeft,
    y: canvasRect.top - layerRect.top - layer.clientTop + layer.scrollTop,
  };
}

/** Geometry over one painted frame; every method returns null once `current` turns false. */
export function createPluginGeometry(
  layout: XlsxPluginLayout,
  painted: PaintedGrid,
  canvas: Element,
  layer: HTMLElement,
  current: () => boolean
): XlsxPluginGeometry {
  const place = (rect: { x: number; y: number; w: number; h: number } | null) =>
    rect ? gridRectToOverlay(painted, rect, canvasOrigin(canvas, layer)) : null;
  return {
    layout,
    getCellRect(target) {
      if (!current() || target?.sheetId !== layout.sheetId) return null;
      if (!isIndex(target.row) || !isIndex(target.col)) return null;
      return place(cellRect(painted.frame.grid, target.row, target.col));
    },
    getRangeRect(target) {
      if (!current() || target?.sheetId !== layout.sheetId) return null;
      const range: CellRange | undefined = target.range;
      if (
        !range ||
        ![range.top, range.left, range.bottom, range.right].every(isIndex) ||
        range.top > range.bottom ||
        range.left > range.right
      ) {
        return null;
      }
      return place(rangeRect(painted.frame.grid, range));
    },
    getPositionAtPoint: null,
  };
}
