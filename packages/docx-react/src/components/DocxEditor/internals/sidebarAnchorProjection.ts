import {
  displayPageCanvases,
  effectiveZoom,
  resolveDisplayPageClientRect,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';

export const SIDEBAR_ANCHOR_EMIT_MS = 150;
export const SIDEBAR_ANCHOR_STALE_MS = 400;

export function sidebarAnchorProjectY(
  host: HTMLElement,
  target: HTMLElement,
  queries: DisplayListQueries,
  zoom: number
): (rect: DisplayListRect) => number | null {
  const targetRect = target.getBoundingClientRect();
  const targetZoom = effectiveZoom(target);
  const canvasByPage = new Map<number, HTMLCanvasElement>();
  for (const canvas of displayPageCanvases(host)) {
    const pageIndex = Number(canvas.dataset.pageIndex);
    if (Number.isFinite(pageIndex)) canvasByPage.set(pageIndex, canvas);
  }
  return (rect) => {
    const pageRect =
      canvasByPage.get(rect.pageIndex)?.getBoundingClientRect() ??
      resolveDisplayPageClientRect(host, queries, rect.pageIndex);
    const pageSize = queries.pageSize(rect.pageIndex);
    if (!pageRect || !pageSize || pageSize.height <= 0) return null;
    return (
      (pageRect.top - targetRect.top + rect.y * (pageRect.height / pageSize.height)) /
      (zoom * targetZoom)
    );
  };
}
