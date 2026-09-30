/** Display-list-backed scroll/ref API for PagedEditor. */

import { useCallback, useEffect, useRef } from 'react';
import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { ParagraphHighlightOptions, ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';

import type { YrsInputRef } from '../YrsInput';
import { runAfterFrames } from '../internals/scrollUtils';
import { scrollViewport } from '../internals/viewportBand';

export interface UsePagedScrollApiOptions {
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  yrsSession: YrsSession | null;
  yrsLocToDisplayPosition: (loc: YrsLoc) => number | null;
  getScrollContainer: () => HTMLDivElement | null;
  displayListQueries?: DisplayListQueries | null;
  /** The current layout: a page past a partial one's last waits for the full layout. */
  layout?: Layout | null;
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  onNavigationIntent?: () => void;
  /** Counts navigation intents; a scroll waiting for the full layout drops on a newer one. */
  navigationEpoch?: () => number;
  requestCanvasParagraphFlash?: (req: {
    from: number;
    to: number;
    options?: ParagraphHighlightOptions;
  }) => void;
}

/** How revealing a display position went. */
export type RevealPositionOutcome = 'scrolled' | 'layout-unavailable' | 'unsupported';

export interface UsePagedScrollApiReturn {
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
  /**
   * Scrolls a position into view without touching focus or selection. Aborting `signal` stops
   * following the position while its page is still being built.
   */
  revealPositionImpl: (position: number, signal?: AbortSignal) => RevealPositionOutcome;
  scrollToPageImpl: (pageNumber: number) => void;
  scrollToParaIdImpl: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
}

const SMOOTH_SCROLL_VIEWPORTS = 2;
const REFINE_WINDOW_MS = 3000;
const USER_SCROLL_EVENTS = ['wheel', 'touchstart', 'pointerdown', 'keydown'] as const;

interface PendingRefine {
  position: number;
  pageIndex: number;
  until: number;
  stop: AbortController;
  /** The document version `position` belongs to, when a reveal set it. */
  version?: string;
}

function isUnbuiltPage(queries: DisplayListQueries, pageIndex: number): boolean {
  return queries.displayList?.pages[pageIndex]?.unbuilt === true;
}

export function usePagedScrollApi(opts: UsePagedScrollApiOptions): UsePagedScrollApiReturn {
  const {
    pagesContainerRef,
    yrsInputRef,
    yrsSession,
    yrsLocToDisplayPosition,
    getScrollContainer,
    displayListQueries = null,
    layout = null,
    canvasHostRef,
    onNavigationIntent,
    navigationEpoch,
    requestCanvasParagraphFlash,
  } = opts;
  const scrollAbortRef = useRef<AbortController | null>(null);
  // A position on a page that is not built yet scrolls to a best guess. As
  // pages get built, the scroll follows the position until it lands on a built
  // page, the attempt runs out, or the user scrolls or navigates on their own.
  const pendingRefineRef = useRef<PendingRefine | null>(null);
  const clearPendingRefine = useCallback(() => {
    pendingRefineRef.current?.stop.abort();
    pendingRefineRef.current = null;
  }, []);

  useEffect(
    () => () => {
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = null;
      clearPendingRefine();
    },
    [clearPendingRefine]
  );

  const scrollRectIntoView = useCallback(
    (rect: DisplayListRect, smooth: boolean): boolean => {
      const queries = displayListQueries;
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      if (!queries || !host) return false;
      const pageRect = resolveDisplayPageClientRect(host, queries, rect.pageIndex);
      const pageSize = queries.pageSize(rect.pageIndex);
      if (!pageRect || !pageSize) return false;
      const scroller = getScrollContainer() ?? findVerticalScrollParentOrRoot(host);
      const viewport = scrollViewport(scroller);
      const scaleY = pageSize.height > 0 ? pageRect.height / pageSize.height : 1;
      const clientY = pageRect.top + (rect.y + rect.height / 2) * scaleY;
      const top =
        scroller.scrollTop +
        clientY / viewport.zoom -
        viewport.top / viewport.zoom -
        viewport.height / 2;
      const near = Math.abs(top - scroller.scrollTop) <= viewport.height * SMOOTH_SCROLL_VIEWPORTS;
      // 'auto' would follow a CSS `scroll-behavior: smooth` and animate anyway
      scroller.scrollTo({ top, behavior: smooth ? (near ? 'smooth' : 'instant') : 'auto' });
      return true;
    },
    [canvasHostRef, displayListQueries, getScrollContainer, pagesContainerRef]
  );

  const scrollAnchorIntoView = useCallback(
    (queries: DisplayListQueries, rect: DisplayListRect, position: number, smooth: boolean) => {
      clearPendingRefine();
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      if (host && isUnbuiltPage(queries, rect.pageIndex)) {
        const stop = new AbortController();
        const scroller = getScrollContainer() ?? findVerticalScrollParentOrRoot(host);
        const listening = { passive: true, signal: stop.signal };
        for (const type of USER_SCROLL_EVENTS) {
          scroller.addEventListener(type, clearPendingRefine, listening);
        }
        const until = performance.now() + REFINE_WINDOW_MS;
        pendingRefineRef.current = { position, pageIndex: rect.pageIndex, until, stop };
      }
      return scrollRectIntoView(rect, smooth);
    },
    [canvasHostRef, clearPendingRefine, getScrollContainer, pagesContainerRef, scrollRectIntoView]
  );

  useEffect(() => {
    const pending = pendingRefineRef.current;
    if (!pending || !displayListQueries) return;
    // an edit moved the positions: the old one now names other text
    if (pending.version !== undefined && pending.version !== yrsSession?.version()) {
      clearPendingRefine();
      return;
    }
    const rect =
      performance.now() <= pending.until ? displayListQueries.anchorRect(pending.position) : null;
    if (!rect) {
      clearPendingRefine();
      return;
    }
    if (isUnbuiltPage(displayListQueries, rect.pageIndex)) {
      if (rect.pageIndex === pending.pageIndex) return;
      pending.pageIndex = rect.pageIndex;
    } else {
      clearPendingRefine();
    }
    scrollRectIntoView(rect, false);
  }, [clearPendingRefine, displayListQueries, scrollRectIntoView, yrsSession]);

  const scrollToPositionImpl = useCallback(
    (pmPos: number, forParaIdScroll = false) => {
      if (!Number.isInteger(pmPos) || pmPos < 0 || !displayListQueries) return;
      onNavigationIntent?.();
      clearPendingRefine();
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = new AbortController();
      const rect = displayListQueries.anchorRect(pmPos);
      if (rect) scrollAnchorIntoView(displayListQueries, rect, pmPos, !forParaIdScroll);
    },
    [clearPendingRefine, displayListQueries, onNavigationIntent, scrollAnchorIntoView]
  );

  const revealPositionImpl = useCallback(
    (position: number, signal?: AbortSignal): RevealPositionOutcome => {
      if (!Number.isInteger(position) || position < 0) return 'unsupported';
      if (!displayListQueries) return 'layout-unavailable';
      clearPendingRefine();
      const rect = displayListQueries.anchorRect(position);
      if (!rect) return 'unsupported';
      onNavigationIntent?.();
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = new AbortController();
      const scrolled = scrollAnchorIntoView(displayListQueries, rect, position, true);
      const pending = pendingRefineRef.current;
      if (pending) pending.version = yrsSession?.version();
      if (pending && signal) {
        // The caller ends the follow through `signal`, however long the page takes to build.
        pending.until = Number.POSITIVE_INFINITY;
        const stop = () => {
          if (pendingRefineRef.current === pending) clearPendingRefine();
        };
        if (signal.aborted) stop();
        else signal.addEventListener('abort', stop, { once: true, signal: pending.stop.signal });
      }
      return scrolled ? 'scrolled' : 'layout-unavailable';
    },
    [clearPendingRefine, displayListQueries, onNavigationIntent, scrollAnchorIntoView, yrsSession]
  );

  const pendingPageRef = useRef<{
    page: number;
    epoch: number | undefined;
    session: YrsSession | null;
  } | null>(null);
  const scrollToPageImpl = useCallback(
    (pageNumber: number): void => {
      pendingPageRef.current = null;
      clearPendingRefine();
      if (!Number.isInteger(pageNumber) || pageNumber < 1 || !displayListQueries) return;
      if (pageNumber > displayListQueries.pageCount()) {
        if (layout?.partial) {
          pendingPageRef.current = {
            page: pageNumber,
            epoch: navigationEpoch?.(),
            session: yrsSession,
          };
        }
        return;
      }
      onNavigationIntent?.();
      const bounds = displayListQueries.pageBounds(pageNumber - 1);
      if (bounds) scrollRectIntoView(bounds, true);
    },
    [
      clearPendingRefine,
      displayListQueries,
      layout,
      navigationEpoch,
      onNavigationIntent,
      scrollRectIntoView,
      yrsSession,
    ]
  );

  useEffect(() => {
    const pending = pendingPageRef.current;
    if (!pending || !displayListQueries) return;
    if (pending.session !== yrsSession || pending.epoch !== navigationEpoch?.()) {
      pendingPageRef.current = null;
      return;
    }
    const pages = displayListQueries.pageCount();
    const complete = layout !== null && !layout.partial && pages === layout.pages.length;
    if (pending.page <= pages || complete) scrollToPageImpl(pending.page);
  }, [displayListQueries, layout, navigationEpoch, scrollToPageImpl, yrsSession]);

  const scrollToParaIdImpl = useCallback(
    (paraId: string, options?: ScrollToParaIdOptions): boolean => {
      if (!yrsSession) return false;
      let story: string | null = null;
      for (const storyId of yrsSession.storyIds()) {
        if (yrsSession.paragraphs(storyId).some((paragraph) => paragraph.paraId === paraId)) {
          story = storyId;
          break;
        }
      }
      if (!story) return false;
      const span = yrsSession.locateParagraph(story, paraId);
      const startLoc = { story, paraId, offset: 0 };
      const endLoc = { story, paraId, offset: Math.max(0, span.end - span.start) };
      const startPos = yrsLocToDisplayPosition(startLoc);
      if (startPos == null || startPos < 0) return false;
      scrollToPositionImpl(startPos, true);
      yrsSession.setSelection(startLoc);
      if (options?.highlight && requestCanvasParagraphFlash) {
        const endPos = yrsLocToDisplayPosition(endLoc) ?? startPos + 1;
        requestCanvasParagraphFlash({
          from: startPos,
          to: Math.max(startPos + 1, endPos),
          options: options.highlight,
        });
      }
      const signal = scrollAbortRef.current?.signal;
      if (!signal) return true;
      runAfterFrames(() => yrsInputRef.current?.focus(), signal);
      return true;
    },
    [requestCanvasParagraphFlash, scrollToPositionImpl, yrsInputRef, yrsLocToDisplayPosition, yrsSession]
  );

  return { scrollToPositionImpl, revealPositionImpl, scrollToPageImpl, scrollToParaIdImpl };
}
