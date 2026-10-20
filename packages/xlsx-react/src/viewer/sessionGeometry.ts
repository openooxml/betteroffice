import type { CellRange, DisplayList, MergedRange, SelectionLimits, SheetInfo } from '@betteroffice/xlsx';
import { cellAtPoint, chartRegionAtPoint } from '@betteroffice/xlsx';
import type { XlsxPointPosition } from '../XlsxEditor';

const COL_W = 96;
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

function trackCount(count: number, spans: [number, number][] | undefined, limit: number): number {
  let last = count - 1;
  for (const [start, end] of spans ?? []) {
    if (start > last) break;
    last += end - start + 1;
  }
  return Math.min(last + 1, limit);
}

export function deriveLimits(
  dl: DisplayList | null,
  info: Pick<SheetInfo, 'contentWidth' | 'contentHeight'>,
  viewportHeight: number
): SelectionLimits {
  const rowH = medianTrack(dl?.grid?.rowOffsets, ROW_H);
  const colW = medianTrack(dl?.grid?.colOffsets, COL_W);
  const rows = trackCount(Math.max(1, Math.round(info.contentHeight / rowH)) + 1, dl?.grid?.hiddenRowSpans, 1_048_576);
  const cols = trackCount(Math.max(1, Math.round(info.contentWidth / colW)) + 1, dl?.grid?.hiddenColSpans, 16_384);
  const rowsPerPage = Math.max(1, Math.floor(viewportHeight / rowH));
  return { rows, cols, rowsPerPage };
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
