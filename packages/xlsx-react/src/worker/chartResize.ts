import type { ChartResizeEdges, Rect } from '@betteroffice/xlsx';

export type ChartResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

/** What a selected chart lets the pointer do, resolved from its painted region. */
export interface SelectedChart {
  id: string;
  movable: boolean;
  resizable: boolean;
  placeholder: boolean;
}

/** The live pointer preview drawn over a selected chart. */
export type ChartDragPreview =
  | { kind: 'move'; x: number; y: number }
  | { kind: 'resize'; handle: ChartResizeHandle; dx: number; dy: number };

export interface ChartResizeHandleSpec {
  handle: ChartResizeHandle;
  /** where the handle sits within the frame, 0..1 on each axis */
  x: number;
  y: number;
  cursor: string;
}

/** The eight handles, in clockwise order from the top-left corner. */
export const CHART_RESIZE_HANDLES: readonly ChartResizeHandleSpec[] = [
  { handle: 'nw', x: 0, y: 0, cursor: 'nwse-resize' },
  { handle: 'n', x: 0.5, y: 0, cursor: 'ns-resize' },
  { handle: 'ne', x: 1, y: 0, cursor: 'nesw-resize' },
  { handle: 'e', x: 1, y: 0.5, cursor: 'ew-resize' },
  { handle: 'se', x: 1, y: 1, cursor: 'nwse-resize' },
  { handle: 's', x: 0.5, y: 1, cursor: 'ns-resize' },
  { handle: 'sw', x: 0, y: 1, cursor: 'nesw-resize' },
  { handle: 'w', x: 0, y: 0.5, cursor: 'ew-resize' },
];

/**
 * The edge movement a drag on `handle` asks the engine for. Each edge delta is
 * the content-pixel travel of that edge, so a south-east drag pushes only the
 * right and bottom edges.
 */
export function chartResizeEdges(
  handle: ChartResizeHandle,
  dx: number,
  dy: number
): ChartResizeEdges {
  const west = handle === 'nw' || handle === 'w' || handle === 'sw';
  const east = handle === 'ne' || handle === 'e' || handle === 'se';
  const north = handle === 'nw' || handle === 'n' || handle === 'ne';
  const south = handle === 'sw' || handle === 's' || handle === 'se';
  return { left: west ? dx : 0, top: north ? dy : 0, right: east ? dx : 0, bottom: south ? dy : 0 };
}

/** What the painted frame knows about the sheet around a chart. */
export interface ChartResizeViewport {
  /** content-pixel offset of the painted frame's origin (the scroll position). */
  x: number;
  y: number;
  frozenRows: number;
  frozenCols: number;
  /** frozen pane extent in content pixels, when the caller can read it. */
  frozenWidth?: number;
  frozenHeight?: number;
  /** worksheet extent in content pixels, when the caller can read it. */
  sheetWidth?: number;
  sheetHeight?: number;
}

/** A chart frame mapped back into sheet space, where the engine clamps. */
export interface ChartResizeSpace {
  originX: number;
  originY: number;
  width: number;
  height: number;
}

/**
 * Recover the sheet-space origin of a painted chart. A chart pinned inside a
 * frozen pane (the resolver pins it when its first cell is frozen) already has
 * sheet coordinates; every other chart is painted relative to the viewport and
 * needs the scroll origin added back. An unpinned chart's clip starts at or
 * past the frozen edge, a pinned one's inside the pane; without the pane extent
 * only a clip at 0 reads as pinned.
 */
export function chartResizeSpace(
  chart: { clip: Rect },
  viewport: ChartResizeViewport
): ChartResizeSpace {
  const inPane = (clip: number, extent: number | undefined) =>
    extent !== undefined && extent > 0 ? clip < extent - 0.5 : clip <= 0.5;
  const pinnedX = viewport.frozenCols > 0 && inPane(chart.clip.x, viewport.frozenWidth);
  const pinnedY = viewport.frozenRows > 0 && inPane(chart.clip.y, viewport.frozenHeight);
  return {
    originX: pinnedX ? 0 : viewport.x,
    originY: pinnedY ? 0 : viewport.y,
    width: viewport.sheetWidth ?? Infinity,
    height: viewport.sheetHeight ?? Infinity,
  };
}

/** One axis of the engine's `clamp_resized_edges`, including its min-size order. */
function clampResizedEdges(
  start: number,
  end: number,
  extent: number,
  startDelta: number,
  endDelta: number
): [number, number] {
  const min = 1;
  let nextStart = Math.min(Math.max(start + startDelta, 0), extent);
  let nextEnd = Math.min(Math.max(end + endDelta, 0), extent);
  if (nextEnd - nextStart < min) {
    if (startDelta !== 0 && endDelta === 0) nextStart = nextEnd - min;
    else nextEnd = nextStart + min;
    if (nextEnd > extent) {
      nextEnd = extent;
      nextStart = Math.max(nextEnd - min, 0);
    }
    if (nextStart < 0) {
      nextStart = 0;
      nextEnd = Math.min(min, extent);
    }
  }
  return [nextStart, nextEnd];
}

/**
 * The frame a live resize previews. `rect` is viewport-local; when `space` is
 * supplied the edges are clamped in sheet space, like the engine, and projected
 * back. Without it the result is the old viewport-local approximation, kept for
 * callers that only need the corners' relative movement.
 */
export function chartResizeRect(
  rect: Rect,
  handle: ChartResizeHandle,
  dx: number,
  dy: number,
  space?: ChartResizeSpace
): Rect {
  const edges = chartResizeEdges(handle, dx, dy);
  const originX = space?.originX ?? 0;
  const originY = space?.originY ?? 0;
  const sheetX = rect.x + originX;
  const sheetY = rect.y + originY;
  const [left, right] = clampResizedEdges(sheetX, sheetX + rect.w, space?.width ?? Infinity,
    edges.left, edges.right);
  const [top, bottom] = clampResizedEdges(sheetY, sheetY + rect.h, space?.height ?? Infinity,
    edges.top, edges.bottom);
  return { x: left - originX, y: top - originY, w: right - left, h: bottom - top };
}
