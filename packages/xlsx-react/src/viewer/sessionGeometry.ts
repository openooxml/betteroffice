import type { CellAddr, CellRange, DisplayList, MergedRange, Rect, SelectionLimits, SheetInfo, Viewport, WorkbookSheetView } from '@betteroffice/xlsx';
import { cellAtPoint, cellRect, chartRegionAtPoint } from '@betteroffice/xlsx';
import type { XlsxPointPosition } from '../XlsxEditor';

const ROW_H = 24;

export function positionAtPoint(
  frame: DisplayList, sheet: number, canvas: HTMLCanvasElement | null, clientX: number, clientY: number
): XlsxPointPosition | null {
  if (!canvas || !frame.grid || !Number.isFinite(clientX) || !Number.isFinite(clientY)) return null;
  const rect = canvas.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0 || clientX < rect.left || clientY < rect.top ||
      clientX >= rect.right || clientY >= rect.bottom) return null;
  const x = (clientX - rect.left) * frame.width / rect.width;
  const y = (clientY - rect.top) * frame.height / rect.height;
  if (chartRegionAtPoint(frame.charts, x, y)) return null;
  const cell = cellAtPoint(frame.grid, x, y);
  return cell ? { ...cell, sheet } : null;
}

function medianTrack(offsets: number[] | undefined, fallback: number): number {
  if (!offsets || offsets.length < 2) return fallback;
  const gaps: number[] = [];
  for (let i = 1; i < offsets.length; i++) gaps.push(offsets[i] - offsets[i - 1]);
  gaps.sort((a, b) => a - b);
  const mid = gaps[gaps.length >> 1];
  return mid > 0 ? mid : fallback;
}

export function deriveLimits(
  dl: DisplayList | null,
  viewportHeight: number
): SelectionLimits {
  const rowH = medianTrack(dl?.grid?.rowOffsets, ROW_H);
  const rowsPerPage = Math.max(1, Math.floor(viewportHeight / rowH));
  return { rows: 1_048_576, cols: 16_384, rowsPerPage };
}

export function extendScrollExtent<T extends Pick<SheetInfo, 'contentWidth' | 'contentHeight'>>(
  scroll: HTMLElement, info: T, rect: Rect, zoom: number
): T {
  const next = {
    ...info,
    contentWidth: Math.max(info.contentWidth, rect.x + rect.w),
    contentHeight: Math.max(info.contentHeight, rect.y + rect.h),
  };
  const content = scroll.firstElementChild as HTMLElement | null;
  if (content) {
    content.style.width = `${next.contentWidth * zoom}px`;
    content.style.height = `${next.contentHeight * zoom}px`;
  }
  return next;
}

export function cellInViewport(scroll: HTMLElement, rect: Rect, zoom: number): boolean {
  const left = scroll.scrollLeft / zoom;
  const top = scroll.scrollTop / zoom;
  return rect.x + rect.w > left && rect.y + rect.h > top &&
    rect.x < left + scroll.clientWidth / zoom && rect.y < top + scroll.clientHeight / zoom;
}

export function laidOutCellRect(
  frame: DisplayList, viewport: Viewport, view: WorkbookSheetView, cell: CellAddr
): Rect | null {
  const rect = cellRect(frame.grid, cell.row, cell.col);
  return rect ? {
    ...rect,
    x: rect.x + (cell.col < view.frozenCols ? 0 : viewport.x),
    y: rect.y + (cell.row < view.frozenRows ? 0 : viewport.y),
  } : null;
}

export function nearestCellScroll(
  rect: Rect, cell: CellAddr, view: WorkbookSheetView, viewport: Viewport, padding = 0
): { x: number; y: number } {
  const axis = (position: number, size: number, frozen: boolean, extent: number, scroll: number, body: number) => {
    if (frozen) return scroll;
    const start = position - extent;
    const available = Math.max(0, body - extent);
    if (start >= scroll + padding && start + size <= scroll + available - padding) return scroll;
    if (start < scroll + padding || size + padding * 2 >= available) return Math.max(0, start - padding);
    return Math.max(0, start + size - available + padding);
  };
  return {
    x: axis(rect.x, rect.w, cell.col < view.frozenCols, view.frozenWidth, viewport.x, viewport.width),
    y: axis(rect.y, rect.h, cell.row < view.frozenRows, view.frozenHeight, viewport.y, viewport.height),
  };
}

export function scaledRect(rect: { x: number; y: number; w: number; h: number }, zoom: number) {
  return {
    x: rect.x * zoom,
    y: rect.y * zoom,
    w: rect.w * zoom,
    h: rect.h * zoom,
  };
}

function intersects(left: CellRange, right: CellRange): boolean {
  return (
    left.left <= right.right &&
    left.right >= right.left &&
    left.top <= right.bottom &&
    left.bottom >= right.top
  );
}

export function expandRangeToMergedCells(
  range: CellRange,
  mergedRanges: readonly MergedRange[]
): CellRange {
  const expanded = { ...range };
  let changed = true;
  while (changed) {
    changed = false;
    for (const merged of mergedRanges) {
      const mergedRange = {
        top: Math.min(merged.start.row, merged.end.row),
        left: Math.min(merged.start.col, merged.end.col),
        bottom: Math.max(merged.start.row, merged.end.row),
        right: Math.max(merged.start.col, merged.end.col),
      };
      if (!intersects(expanded, mergedRange)) continue;
      const top = Math.min(expanded.top, mergedRange.top);
      const left = Math.min(expanded.left, mergedRange.left);
      const bottom = Math.max(expanded.bottom, mergedRange.bottom);
      const right = Math.max(expanded.right, mergedRange.right);
      if (
        top === expanded.top &&
        left === expanded.left &&
        bottom === expanded.bottom &&
        right === expanded.right
      ) {
        continue;
      }
      Object.assign(expanded, { top, left, bottom, right });
      changed = true;
    }
  }
  return expanded;
}
