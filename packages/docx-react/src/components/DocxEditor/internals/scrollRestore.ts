import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
  type DisplayListRect,
  type DisplayListVisualLine,
} from '@betteroffice/docx/layout/render';
import type { YrsStickyPosition } from '@betteroffice/docx/yrs';
import {
  computeViewportAnchoredScrollTop,
  type ViewportAnchorSnapshot,
} from './viewportAnchoring';

export interface DisplayListScrollAnchor {
  pmPos: number;
  clientOffset: number | null;
  ratio: number;
  scrollTopSnapshot: number;
}

interface ParagraphViewportTarget {
  kind: 'paragraph';
  paraId: string;
  position: YrsStickyPosition;
}

interface PageViewportTarget {
  kind: 'page';
  pageIndex: number;
  pageY: number;
}

export interface DisplayListViewportAnchor extends ViewportAnchorSnapshot {
  target: ParagraphViewportTarget | PageViewportTarget | null;
}

interface PageProjection {
  top: number;
  scaleY: number;
}

export type CaptureViewportPosition = (
  displayPosition: number,
  paraId: string
) => YrsStickyPosition | null;

export type ResolveViewportPosition = (
  position: YrsStickyPosition,
  paraId: string
) => number | null;

function pageProjection(
  queries: DisplayListQueries,
  host: HTMLElement,
  pageIndex: number,
  cache?: Map<number, PageProjection | null>
): PageProjection | null {
  if (cache?.has(pageIndex)) return cache.get(pageIndex) ?? null;
  const pageRect = resolveDisplayPageClientRect(host, queries, pageIndex);
  const pageSize = queries.pageSize(pageIndex);
  const projection =
    pageRect && pageSize && pageSize.height > 0
      ? { top: pageRect.top, scaleY: pageRect.height / pageSize.height }
      : null;
  cache?.set(pageIndex, projection);
  return projection;
}

function projectedRectClientY(
  queries: DisplayListQueries,
  host: HTMLElement,
  rect: DisplayListRect,
  cache?: Map<number, PageProjection | null>
): { top: number; bottom: number } | null {
  const projection = pageProjection(queries, host, rect.pageIndex, cache);
  if (!projection) return null;
  const top = projection.top + rect.y * projection.scaleY;
  return { top, bottom: top + rect.height * projection.scaleY };
}

function projectedAnchorClientY(
  queries: DisplayListQueries,
  host: HTMLElement,
  pmPos: number
): number | null {
  const anchor = queries.anchorRect(pmPos);
  return anchor ? (projectedRectClientY(queries, host, anchor)?.top ?? null) : null;
}

function visualLineOrder(left: DisplayListVisualLine, right: DisplayListVisualLine): number {
  return left.pageIndex - right.pageIndex || left.y - right.y || left.x - right.x;
}

function paragraphTargetLine(
  queries: DisplayListQueries,
  target: ParagraphViewportTarget,
  resolvePosition: ResolveViewportPosition
): DisplayListVisualLine | null {
  const lines = queries
    .visualLines()
    .filter((line) => line.paraId === target.paraId)
    .sort(visualLineOrder);
  if (lines.length === 0) return null;
  const position = resolvePosition(target.position, target.paraId);
  if (position == null) return null;
  const paragraphStart = Math.min(...lines.map((line) => line.from));
  const paragraphEnd = Math.max(...lines.map((line) => line.to));
  if (position < paragraphStart || position > paragraphEnd) return null;
  return (
    lines.find((line) => line.from === position) ??
    lines.find((line) => line.from < position && position <= line.to) ??
    lines.reduce((closest, line) =>
      Math.abs(line.from - position) < Math.abs(closest.from - position) ? line : closest
    )
  );
}

function viewportTargetClientY(
  anchor: DisplayListViewportAnchor,
  queries: DisplayListQueries,
  host: HTMLElement,
  resolvePosition: ResolveViewportPosition
): number | null {
  const target = anchor.target;
  if (!target) return null;
  if (target.kind === 'paragraph') {
    const line = paragraphTargetLine(queries, target, resolvePosition);
    return line ? (projectedRectClientY(queries, host, line)?.top ?? null) : null;
  }
  const pageRect = resolveDisplayPageClientRect(host, queries, target.pageIndex);
  const pageSize = queries.pageSize(target.pageIndex);
  if (!pageRect || !pageSize || pageSize.height <= 0) return null;
  return pageRect.top + target.pageY * (pageRect.height / pageSize.height);
}

function visibleParagraphAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  viewport: DOMRect,
  lines: readonly DisplayListVisualLine[],
  capturePosition: CaptureViewportPosition
): { target: ParagraphViewportTarget; clientY: number } | null {
  let selected: { line: DisplayListVisualLine; clientY: number } | null = null;
  const projectionCache = new Map<number, PageProjection | null>();
  for (const line of lines) {
    if (!line.paraId) continue;
    const projected = projectedRectClientY(queries, host, line, projectionCache);
    if (!projected || projected.bottom < viewport.top || projected.top > viewport.bottom) continue;
    if (!selected || projected.top < selected.clientY) {
      selected = { line, clientY: projected.top };
    }
  }
  if (!selected?.line.paraId) return null;
  const position = capturePosition(selected.line.from, selected.line.paraId);
  if (!position) return null;
  return {
    target: {
      kind: 'paragraph',
      paraId: selected.line.paraId,
      position,
    },
    clientY: selected.clientY,
  };
}

function visiblePageAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  viewport: DOMRect
): { target: PageViewportTarget; clientY: number } | null {
  for (let pageIndex = 0; pageIndex < queries.pageCount(); pageIndex += 1) {
    const pageRect = resolveDisplayPageClientRect(host, queries, pageIndex);
    const pageSize = queries.pageSize(pageIndex);
    if (
      !pageRect ||
      !pageSize ||
      pageSize.height <= 0 ||
      pageRect.bottom < viewport.top ||
      pageRect.top > viewport.bottom
    ) {
      continue;
    }
    const scaleY = pageRect.height / pageSize.height;
    const pageY = Math.min(
      pageSize.height,
      Math.max(0, (Math.max(viewport.top, pageRect.top) - pageRect.top) / scaleY)
    );
    return {
      target: { kind: 'page', pageIndex, pageY },
      clientY: pageRect.top + pageY * scaleY,
    };
  }
  return null;
}

export function captureDisplayListScrollAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement,
  pmPos: number
): DisplayListScrollAnchor {
  if (!scrollParent.style.overflowAnchor) {
    scrollParent.style.setProperty('overflow-anchor', 'none');
  }
  const maxScroll = Math.max(1, scrollParent.scrollHeight - scrollParent.clientHeight);
  const clientY = projectedAnchorClientY(queries, host, pmPos);
  const scrollerTop = scrollParent.getBoundingClientRect().top;
  return {
    pmPos,
    clientOffset: clientY == null ? null : clientY - scrollerTop,
    ratio: scrollParent.scrollTop / maxScroll,
    scrollTopSnapshot: scrollParent.scrollTop,
  };
}

export function captureDisplayListViewportAnchor(
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement,
  capturePosition: CaptureViewportPosition
): DisplayListViewportAnchor {
  if (!scrollParent.style.overflowAnchor) {
    scrollParent.style.setProperty('overflow-anchor', 'none');
  }
  const viewport = scrollParent.getBoundingClientRect();
  const lines = queries.visualLines();
  const resolved =
    visibleParagraphAnchor(queries, host, viewport, lines, capturePosition) ??
    visiblePageAnchor(queries, host, viewport);
  return {
    target: resolved?.target ?? null,
    viewportOffset: resolved ? resolved.clientY - viewport.top : 0,
    scrollTopSnapshot: scrollParent.scrollTop,
  };
}

export function restoreDisplayListScrollAnchor(
  anchor: DisplayListScrollAnchor,
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement
): void {
  const clientY = projectedAnchorClientY(queries, host, anchor.pmPos);
  if (clientY != null && anchor.clientOffset != null) {
    const currentOffset = clientY - scrollParent.getBoundingClientRect().top;
    scrollParent.scrollTop += anchor.clientOffset - currentOffset;
    return;
  }
  const maxScroll = Math.max(1, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = anchor.ratio * maxScroll;
}

export function restoreDisplayListViewportAnchor(
  anchor: DisplayListViewportAnchor,
  queries: DisplayListQueries,
  host: HTMLElement,
  scrollParent: HTMLElement,
  resolvePosition: ResolveViewportPosition
): void {
  const clientY = viewportTargetClientY(anchor, queries, host, resolvePosition);
  const scrollerTop = scrollParent.getBoundingClientRect().top;
  const nextTargetTop =
    clientY == null ? null : scrollParent.scrollTop + clientY - scrollerTop;
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = computeViewportAnchoredScrollTop(anchor, nextTargetTop, maxScroll);
}

export function restoreScrollSnapshot(
  anchor: Pick<DisplayListScrollAnchor, 'scrollTopSnapshot'>,
  scrollParent: HTMLElement
): void {
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = Math.min(Math.max(0, anchor.scrollTopSnapshot), maxScroll);
}
