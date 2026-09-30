import { useCallback, useEffect, useRef, useState } from 'react';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { displayPageCanvases } from '@betteroffice/docx/layout/render';
import type { PagedEditorRef } from '../PagedEditor';
import { documentPageCount } from './documentPageCount';

interface ScrollPageInfo {
  currentPage: number;
  totalPages: number;
  visible: boolean;
}

/**
 * The 1-based page under the middle of the visible part of `scroller`. Reads
 * the live page canvases where there are any, so zoom, page gaps and content
 * above the pages count; otherwise sums the layout's page heights.
 */
export function pageAtViewportMiddle(scroller: HTMLElement, layout: Layout): number {
  const host = scroller.querySelector<HTMLElement>('.canvas-pages');
  const canvases = host ? displayPageCanvases(host) : [];
  const rect = scroller.getBoundingClientRect();
  if (canvases.length > 0 && rect.height > 0) {
    const top = Math.max(rect.top, 0);
    const bottom = Math.min(rect.bottom, window.innerHeight);
    const middle = bottom > top ? (top + bottom) / 2 : (rect.top + rect.bottom) / 2;
    let low = 0;
    let high = canvases.length - 1;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (canvases[mid]!.getBoundingClientRect().bottom > middle) high = mid;
      else low = mid + 1;
    }
    const pageIndex = Number(canvases[low]!.dataset.pageIndex);
    if (Number.isInteger(pageIndex)) return pageIndex + 1;
  }

  const pageGap = 24; // DEFAULT_PAGE_GAP from PagedEditor
  const paddingTop = 24; // top padding in paged-editor__pages
  const viewportCenter = scroller.scrollTop + scroller.clientHeight / 2;
  let accumulatedY = paddingTop;
  let currentPage = 1;
  for (let i = 0; i < layout.pages.length; i++) {
    const pageHeight = layout.pages[i].size.h;
    const pageEnd = accumulatedY + pageHeight;
    if (viewportCenter < pageEnd) {
      currentPage = i + 1;
      break;
    }
    accumulatedY = pageEnd + pageGap;
    currentPage = i + 2;
  }
  return currentPage;
}

/**
 * Drives the floating page indicator (the "3 of 12" pill that fades in
 * on scroll), then hides it after 600ms of no scrolling, and reads the
 * current page on demand. Re-attaches when the scroll container first
 * mounts, which is after loading completes (the loading state renders a
 * different subtree).
 */
export function useScrollPageInfo({
  scrollContainerRef,
  pagedEditorRef,
}: {
  scrollContainerRef: React.RefObject<HTMLDivElement | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
}) {
  const [scrollPageInfo, setScrollPageInfo] = useState<ScrollPageInfo>({
    currentPage: 1,
    totalPages: 0,
    visible: false,
  });
  const scrollFadeTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  /** The page the scroll position shows now among the laid-out pages, or null without any. */
  const readCurrentPage = useCallback((): number | null => {
    const scroller = scrollContainerRef.current;
    const layout = pagedEditorRef.current?.getLayout();
    if (!scroller || !layout || layout.pages.length === 0) return null;
    return Math.min(pageAtViewportMiddle(scroller, layout), layout.pages.length);
  }, [scrollContainerRef, pagedEditorRef]);

  const scrollContainerEl = scrollContainerRef.current;
  useEffect(() => {
    if (!scrollContainerEl) return;

    const handleScroll = () => {
      const totalPages = documentPageCount(pagedEditorRef.current?.getLayout());
      const currentPage = totalPages === 0 ? null : readCurrentPage();
      if (currentPage === null) return;

      // bail out on unchanged values: this fires per scroll event, and a new
      // object every time re-renders the whole editor tree every frame
      setScrollPageInfo((previous) =>
        previous.currentPage === currentPage &&
        previous.totalPages === totalPages &&
        previous.visible
          ? previous
          : { currentPage, totalPages, visible: true }
      );

      if (scrollFadeTimerRef.current) {
        clearTimeout(scrollFadeTimerRef.current);
      }
      scrollFadeTimerRef.current = setTimeout(() => {
        setScrollPageInfo((prev) => ({ ...prev, visible: false }));
      }, 600);
    };

    scrollContainerEl.addEventListener('scroll', handleScroll, { passive: true });
    return () => {
      scrollContainerEl.removeEventListener('scroll', handleScroll);
      if (scrollFadeTimerRef.current) {
        clearTimeout(scrollFadeTimerRef.current);
      }
    };
  }, [scrollContainerEl, pagedEditorRef, readCurrentPage]);

  return { scrollPageInfo, setScrollPageInfo, readCurrentPage };
}
