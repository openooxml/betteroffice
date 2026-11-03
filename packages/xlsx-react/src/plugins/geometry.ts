import { cellRect, rangeRect } from '@betteroffice/xlsx';
import type { CellRange, DisplayList, GridMeta, Proposal, Viewport } from '@betteroffice/xlsx';
import type { XlsxPointPosition } from '../XlsxEditor.types';
import type { XlsxAnchorGeometryResult, XlsxPluginGeometry, XlsxPluginLayout, XlsxPluginRect } from './types';

export interface ProposalGeometrySnapshot {
  proposals: readonly Proposal[];
  sheetIndex: number;
  previewVersion: number;
}

/** The editor's own hit test over the frame it shows, in client coordinates. */
export type EditorPointQuery = (clientX: number, clientY: number) => XlsxPointPosition | null;

export function frozenGridExtent(
  grid: GridMeta | undefined,
  frozenRows: number,
  frozenCols: number,
  fallback: Pick<Viewport, 'width' | 'height'> = { width: 0, height: 0 }
) {
  const edge = (
    offsets: number[] | undefined,
    indices: number[] | undefined,
    start: number,
    frozen: number,
    whole: number
  ) => {
    if (frozen === 0) return 0;
    const count = indices
      ? indices.filter((index) => index < frozen).length
      : Math.max(0, frozen - start);
    return offsets?.[count] ?? whole;
  };
  return {
    width: edge(grid?.colOffsets, grid?.colIndices, grid?.startCol ?? 0, frozenCols, fallback.width),
    height: edge(grid?.rowOffsets, grid?.rowIndices, grid?.startRow ?? 0, frozenRows, fallback.height),
  };
}

export function visibleGridRange(grid: GridMeta): CellRange | null {
  const rows = grid.rowOffsets.length - 1;
  const cols = grid.colOffsets.length - 1;
  if (rows <= 0 || cols <= 0) return null;
  return {
    top: grid.rowIndices?.[0] ?? grid.startRow,
    left: grid.colIndices?.[0] ?? grid.startCol,
    bottom: grid.rowIndices?.[rows - 1] ?? grid.startRow + rows - 1,
    right: grid.colIndices?.[cols - 1] ?? grid.startCol + cols - 1,
  };
}

/** A grid frame the canvas painted, with what it shows. */
export interface PaintedGrid {
  frame: DisplayList;
  zoom: number;
  /** The workbook version the frame was built from. */
  version: string;
  sheetId: string;
  /** The zero-based index of `sheetId`. */
  sheet: number;
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

/**
 * Geometry over one painted frame; every method returns null once `current` turns false.
 * Point queries go through `pointAt` and keep only answers on the frame's sheet.
 */
export function createPluginGeometry(
  layout: XlsxPluginLayout,
  painted: PaintedGrid,
  canvas: Element,
  layer: HTMLElement,
  current: () => boolean,
  proposals: () => ProposalGeometrySnapshot | null = () => ({ proposals: [], sheetIndex: -1, previewVersion: 0 }),
  pointAt?: EditorPointQuery
): XlsxPluginGeometry {
  const place = (rect: { x: number; y: number; w: number; h: number } | null) =>
    rect ? gridRectToOverlay(painted, rect, canvasOrigin(canvas, layer)) : null;
  return {
    layout,
    getAnchorGeometry(target) {
      const failure = (code: Extract<XlsxAnchorGeometryResult, { ok: false }>['failure']['code'],
        message: string): XlsxAnchorGeometryResult => ({ ok: false, failure: { code, message } });
      if (!current()) return failure('layout-unavailable', 'The layout is no longer painted');
      let snapshot: ProposalGeometrySnapshot | null;
      try { snapshot = proposals(); }
      catch { return failure('read-failed', 'The proposals could not be read'); }
      if (!snapshot) return failure('pending', 'The proposals are being refreshed');
      const { proposals: pending, sheetIndex, previewVersion } = snapshot;
      const proposal = pending.find((value) => value.id === target.id);
      if (!proposal) return failure('unknown-proposal', 'The proposal is not registered in this document');
      const rects = proposal.cells.flatMap((cell) => {
        if (cell.sheet !== sheetIndex) return [];
        const rect = place(cellRect(painted.frame.grid, cell.row, cell.col));
        return rect ? [rect] : [];
      });
      if (!rects.length) return failure('missing-target', 'The proposal has no visible cells on this sheet');
      let x = Infinity;
      let y = Infinity;
      let right = -Infinity;
      let bottom = -Infinity;
      for (const rect of rects) {
        x = Math.min(x, rect.x);
        y = Math.min(y, rect.y);
        right = Math.max(right, rect.x + rect.width);
        bottom = Math.max(bottom, rect.y + rect.height);
      }
      const origin = canvasOrigin(canvas, layer);
      return { ok: true, version: layout.version, previewVersion, layoutId: layout.id,
        rects, anchor: { x, y, width: right - x, height: bottom - y },
        pageRect: { ...origin, width: painted.frame.width * painted.zoom, height: painted.frame.height * painted.zoom } };
    },
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
    getPositionAtPoint: pointAt
      ? (clientX, clientY) => {
          if (!current()) return null;
          const position = pointAt(clientX, clientY);
          if (!position || position.sheet !== painted.sheet) return null;
          if (!isIndex(position.row) || !isIndex(position.col)) return null;
          return {
            sheet: position.sheet,
            sheetId: layout.sheetId,
            row: position.row,
            col: position.col,
            version: layout.version,
            layoutId: layout.id,
          };
        }
      : null,
  };
}
