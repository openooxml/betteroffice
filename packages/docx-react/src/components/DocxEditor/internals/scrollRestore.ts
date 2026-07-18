import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';

export interface DisplayListScrollAnchor {
  pmPos: number;
  clientOffset: number | null;
  ratio: number;
  scrollTopSnapshot: number;
}

function projectedAnchorClientY(
  queries: DisplayListQueries,
  host: HTMLElement,
  pmPos: number
): number | null {
  const anchor = queries.anchorRect(pmPos);
  if (!anchor) return null;
  const pageRect = resolveDisplayPageClientRect(host, queries, anchor.pageIndex);
  const pageSize = queries.pageSize(anchor.pageIndex);
  if (!pageRect || !pageSize || pageSize.height <= 0) return null;
  return pageRect.top + anchor.y * (pageRect.height / pageSize.height);
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

export function restoreScrollSnapshot(
  anchor: DisplayListScrollAnchor,
  scrollParent: HTMLElement
): void {
  const maxScroll = Math.max(0, scrollParent.scrollHeight - scrollParent.clientHeight);
  scrollParent.scrollTop = Math.min(Math.max(0, anchor.scrollTopSnapshot), maxScroll);
}
