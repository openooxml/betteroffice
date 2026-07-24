import {
  CANVAS_PAGE_GAP_PX,
  CANVAS_PAGES_PADDING_PX,
  type DisplayList,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';

export const REMOTE_PRESENCE_MAX_PAGES = 8;
export const REMOTE_PRESENCE_MAX_RECTS = 64;

const REMOTE_PRESENCE_PAGE_BUFFER = 1;

export interface RemotePresencePageMetrics {
  tops: readonly number[];
  bottoms: readonly number[];
  centers: readonly number[];
}

export interface RemotePresencePageWindow {
  start: number;
  end: number;
}

export interface RemotePresenceSelectionRange {
  from: number;
  to: number;
}

function lowerBound(values: readonly number[], target: number): number {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (values[middle] < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

function upperBound(values: readonly number[], target: number): number {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (values[middle] <= target) low = middle + 1;
    else high = middle;
  }
  return low;
}

export function buildRemotePresencePageMetrics(
  displayList: DisplayList,
  zoom: number
): RemotePresencePageMetrics {
  const safeZoom = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const tops: number[] = [];
  const bottoms: number[] = [];
  const centers: number[] = [];
  let top = CANVAS_PAGES_PADDING_PX;
  for (const page of displayList.pages) {
    const bottom = top + Math.max(0, page.height * safeZoom);
    tops.push(top);
    bottoms.push(bottom);
    centers.push((top + bottom) / 2);
    top = bottom + CANVAS_PAGE_GAP_PX;
  }
  return { tops, bottoms, centers };
}

export function remotePresencePageWindow(
  metrics: RemotePresencePageMetrics,
  columnTop: number,
  viewportTop: number,
  viewportBottom: number
): RemotePresencePageWindow | null {
  if (
    metrics.tops.length === 0 ||
    !Number.isFinite(columnTop) ||
    !Number.isFinite(viewportTop) ||
    !Number.isFinite(viewportBottom)
  ) {
    return null;
  }
  const top = Math.min(viewportTop, viewportBottom) - columnTop;
  const bottom = Math.max(viewportTop, viewportBottom) - columnTop;
  const firstVisible = lowerBound(metrics.bottoms, top);
  const lastVisible = upperBound(metrics.tops, bottom) - 1;
  if (firstVisible >= metrics.tops.length || lastVisible < firstVisible) return null;

  const start = Math.max(0, firstVisible - REMOTE_PRESENCE_PAGE_BUFFER);
  const end = Math.min(metrics.tops.length - 1, lastVisible + REMOTE_PRESENCE_PAGE_BUFFER);
  if (end - start + 1 <= REMOTE_PRESENCE_MAX_PAGES) return { start, end };

  const viewportCenter = (top + bottom) / 2;
  const nextCenter = lowerBound(metrics.centers, viewportCenter);
  const previousCenter = Math.max(0, nextCenter - 1);
  const center =
    nextCenter >= metrics.centers.length ||
    Math.abs(metrics.centers[previousCenter] - viewportCenter) <=
      Math.abs(metrics.centers[nextCenter] - viewportCenter)
      ? previousCenter
      : nextCenter;
  const centeredStart = Math.min(
    Math.max(0, center - Math.floor((REMOTE_PRESENCE_MAX_PAGES - 1) / 2)),
    metrics.tops.length - REMOTE_PRESENCE_MAX_PAGES
  );
  return {
    start: centeredStart,
    end: centeredStart + REMOTE_PRESENCE_MAX_PAGES - 1,
  };
}

export function pageInRemotePresenceWindow(
  pageIndex: number,
  pageWindow: RemotePresencePageWindow
): boolean {
  return pageIndex >= pageWindow.start && pageIndex <= pageWindow.end;
}

export function remotePresencePagePositionRange(
  displayList: DisplayList,
  pageWindow: RemotePresencePageWindow
): RemotePresenceSelectionRange | null {
  let pageFrom = Number.POSITIVE_INFINITY;
  let pageTo = Number.NEGATIVE_INFINITY;
  for (let pageIndex = pageWindow.start; pageIndex <= pageWindow.end; pageIndex += 1) {
    const page = displayList.pages[pageIndex];
    if (!page) continue;
    for (const primitive of page.primitives) {
      if (
        primitive.kind !== 'text' &&
        primitive.kind !== 'glyphRun' &&
        primitive.kind !== 'image'
      ) {
        continue;
      }
      const primitiveFrom = primitive.docStart;
      const primitiveTo = primitive.docEnd;
      if (
        typeof primitiveFrom !== 'number' ||
        typeof primitiveTo !== 'number' ||
        !Number.isSafeInteger(primitiveFrom) ||
        !Number.isSafeInteger(primitiveTo)
      ) {
        continue;
      }
      pageFrom = Math.min(pageFrom, primitiveFrom);
      pageTo = Math.max(
        pageTo,
        primitiveFrom === primitiveTo && primitiveTo < Number.MAX_SAFE_INTEGER
          ? primitiveTo + 1
          : primitiveTo
      );
    }
  }
  if (!Number.isFinite(pageFrom) || !Number.isFinite(pageTo)) return null;
  return pageFrom < pageTo ? { from: pageFrom, to: pageTo } : null;
}

export function clampRemoteSelectionRange(
  pageRange: RemotePresenceSelectionRange,
  from: number,
  to: number
): RemotePresenceSelectionRange | null {
  const selectionFrom = Math.min(from, to);
  const selectionTo = Math.max(from, to);
  if (selectionFrom === selectionTo) return null;
  const clampedFrom = Math.max(selectionFrom, pageRange.from);
  const clampedTo = Math.min(selectionTo, pageRange.to);
  return clampedFrom < clampedTo ? { from: clampedFrom, to: clampedTo } : null;
}

function validRect(rect: DisplayListRect): boolean {
  return (
    Number.isSafeInteger(rect.pageIndex) &&
    Number.isFinite(rect.x) &&
    Number.isFinite(rect.y) &&
    Number.isFinite(rect.width) &&
    Number.isFinite(rect.height) &&
    rect.width > 0 &&
    rect.height > 0
  );
}

function unionRects(
  rects: readonly DisplayListRect[],
  start: number,
  end: number
): DisplayListRect {
  const first = rects[start];
  let left = first.x;
  let top = first.y;
  let right = first.x + first.width;
  let bottom = first.y + first.height;
  for (let index = start + 1; index < end; index += 1) {
    const rect = rects[index];
    left = Math.min(left, rect.x);
    top = Math.min(top, rect.y);
    right = Math.max(right, rect.x + rect.width);
    bottom = Math.max(bottom, rect.y + rect.height);
  }
  return {
    pageIndex: first.pageIndex,
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  };
}

function mergeLineRects(rects: readonly DisplayListRect[]): DisplayListRect[] {
  const merged: DisplayListRect[] = [];
  for (const rect of rects) {
    const previous = merged.at(-1);
    if (
      previous &&
      previous.pageIndex === rect.pageIndex &&
      Math.abs(previous.y - rect.y) <= 1 &&
      Math.abs(previous.height - rect.height) <= 1 &&
      rect.x <= previous.x + previous.width + 2
    ) {
      const left = Math.min(previous.x, rect.x);
      const top = Math.min(previous.y, rect.y);
      const right = Math.max(previous.x + previous.width, rect.x + rect.width);
      const bottom = Math.max(previous.y + previous.height, rect.y + rect.height);
      previous.x = left;
      previous.y = top;
      previous.width = right - left;
      previous.height = bottom - top;
    } else {
      merged.push({ ...rect });
    }
  }
  return merged;
}

export function simplifyRemoteSelectionRects(
  rects: readonly DisplayListRect[],
  pageWindow: RemotePresencePageWindow
): DisplayListRect[] {
  const ordered = rects
    .filter((rect) => pageInRemotePresenceWindow(rect.pageIndex, pageWindow) && validRect(rect))
    .slice()
    .sort(
      (left, right) => left.pageIndex - right.pageIndex || left.y - right.y || left.x - right.x
    );
  const merged = mergeLineRects(ordered);
  if (merged.length <= REMOTE_PRESENCE_MAX_RECTS) return merged;

  const byPage = new Map<number, DisplayListRect[]>();
  for (const rect of merged) {
    const pageRects = byPage.get(rect.pageIndex);
    if (pageRects) pageRects.push(rect);
    else byPage.set(rect.pageIndex, [rect]);
  }
  const pages = [...byPage.entries()];
  const allocations = pages.map(() => 1);
  let remaining = REMOTE_PRESENCE_MAX_RECTS - pages.length;
  while (remaining > 0) {
    let choice = -1;
    let score = 0;
    for (let index = 0; index < pages.length; index += 1) {
      const count = pages[index][1].length;
      if (allocations[index] >= count) continue;
      const candidate = count / allocations[index];
      if (candidate > score) {
        score = candidate;
        choice = index;
      }
    }
    if (choice < 0) break;
    allocations[choice] += 1;
    remaining -= 1;
  }

  const simplified: DisplayListRect[] = [];
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
    const pageRects = pages[pageIndex][1];
    const allocation = allocations[pageIndex];
    for (let index = 0; index < allocation; index += 1) {
      const start = Math.floor((index * pageRects.length) / allocation);
      const end = Math.floor(((index + 1) * pageRects.length) / allocation);
      simplified.push(unionRects(pageRects, start, end));
    }
  }
  return simplified;
}
