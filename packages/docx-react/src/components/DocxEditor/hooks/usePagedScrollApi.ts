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
import { layoutScrollCompensation } from '../internals/scrollRestore';
import { runAfterFrames } from '../internals/scrollUtils';
import { scrollViewport } from '../internals/viewportBand';
import type { DisplayPageNavigation } from './useDisplayList';

export interface UsePagedScrollApiOptions {
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  yrsSession: YrsSession | null;
  yrsLocToDisplayPosition: (loc: YrsLoc) => number | null;
  getScrollContainer: () => HTMLDivElement | null;
  displayListQueries?: DisplayListQueries | null;
  /** Builds a navigation's target page now and reports the frame that brings it. */
  pageNavigation?: DisplayPageNavigation | null;
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
const SCROLL_EPSILON = 1;
const USER_SCROLL_EVENTS = ['wheel', 'touchstart', 'pointerdown', 'keydown'] as const;

interface PendingRefine {
  position: number;
  pageIndex: number;
  until: number;
  stop: AbortController;
  scroller: HTMLElement;
  scrollTop: number;
  compensationSequence: number;
  smoothTarget?: number;
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
    pageNavigation,
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
    const pending = pendingRefineRef.current;
    if (!pending) return;
    pending.stop.abort();
    pendingRefineRef.current = null;
    pageNavigation?.buildPages([]);
  }, [pageNavigation]);
  const checkPendingScroll = useCallback(() => {
    const pending = pendingRefineRef.current;
    if (!pending) return;
    const top = pending.scroller.scrollTop;
    if (Math.abs(top - pending.scrollTop) <= SCROLL_EPSILON) return;
    const compensation = layoutScrollCompensation(pending.scroller);
    const target = pending.smoothTarget;
    if (
      target !== undefined &&
      top >= Math.min(pending.scrollTop, target) - SCROLL_EPSILON &&
      top <= Math.max(pending.scrollTop, target) + SCROLL_EPSILON
    ) {
      pending.scrollTop = top;
      pending.compensationSequence = compensation?.sequence ?? pending.compensationSequence;
      if (Math.abs(top - target) <= SCROLL_EPSILON) pending.smoothTarget = undefined;
      return;
    }
    const maxScrollTop = Math.max(0, pending.scroller.scrollHeight - pending.scroller.clientHeight);
    if (
      (compensation &&
        compensation.sequence > pending.compensationSequence &&
        (Math.abs(compensation.from - pending.scrollTop) <= SCROLL_EPSILON ||
          Math.abs(compensation.scrollTopSnapshot - pending.scrollTop) <= SCROLL_EPSILON ||
          Math.abs(compensation.from - Math.min(pending.scrollTop, maxScrollTop)) <= SCROLL_EPSILON) &&
        Math.abs(compensation.to - top) <= SCROLL_EPSILON) ||
      (pending.scrollTop > maxScrollTop + SCROLL_EPSILON &&
        Math.abs(top - maxScrollTop) <= SCROLL_EPSILON)
    ) {
      pending.scrollTop = top;
      pending.compensationSequence = compensation?.sequence ?? pending.compensationSequence;
      return;
    }
    clearPendingRefine();
  }, [clearPendingRefine]);

  useEffect(
    () => () => {
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = null;
      clearPendingRefine();
    },
    [clearPendingRefine]
  );

  const scrollRectIntoView = useCallback(
    (rect: DisplayListRect, smooth: boolean, queries = displayListQueries): boolean => {
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
      const behavior = smooth ? (near ? 'smooth' : 'instant') : 'auto';
      const pending = pendingRefineRef.current;
      if (pending) {
        pending.smoothTarget =
          behavior === 'smooth' ||
          (behavior === 'auto' && getComputedStyle(scroller).scrollBehavior === 'smooth')
            ? top
            : undefined;
      }
      scroller.scrollTo({ top, behavior });
      if (pending) {
        pending.scrollTop = scroller.scrollTop;
        pending.compensationSequence = layoutScrollCompensation(scroller)?.sequence ?? 0;
      }
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
        scroller.addEventListener('scroll', checkPendingScroll, listening);
        scroller.addEventListener(
          'scrollend',
          () => {
            const pending = pendingRefineRef.current;
            if (pending) pending.smoothTarget = undefined;
          },
          listening
        );
        scroller.ownerDocument.addEventListener('keydown', clearPendingRefine, {
          ...listening,
          capture: true,
        });
        const until = performance.now() + REFINE_WINDOW_MS;
        pendingRefineRef.current = {
          position,
          pageIndex: rect.pageIndex,
          until,
          stop,
          scroller,
          scrollTop: scroller.scrollTop,
          compensationSequence: layoutScrollCompensation(scroller)?.sequence ?? 0,
        };
      }
      const scrolled = scrollRectIntoView(rect, smooth);
      if (pendingRefineRef.current) pageNavigation?.buildPages([rect.pageIndex]);
      return scrolled;
    },
    [
      canvasHostRef,
      checkPendingScroll,
      clearPendingRefine,
      getScrollContainer,
      pageNavigation,
      pagesContainerRef,
      scrollRectIntoView,
    ]
  );

  const refinePending = useCallback(
    (queries: DisplayListQueries) => {
      checkPendingScroll();
      const pending = pendingRefineRef.current;
      if (!pending) return;
      // an edit moved the positions: the old one now names other text
      if (pending.version !== undefined && pending.version !== yrsSession?.version()) {
        clearPendingRefine();
        return;
      }
      const rect = performance.now() <= pending.until ? queries.anchorRect(pending.position) : null;
      if (!rect) {
        clearPendingRefine();
        return;
      }
      if (isUnbuiltPage(queries, rect.pageIndex)) {
        if (rect.pageIndex === pending.pageIndex) return;
        pending.pageIndex = rect.pageIndex;
        pageNavigation?.buildPages([rect.pageIndex]);
      } else {
        clearPendingRefine();
      }
      scrollRectIntoView(rect, false, queries);
    },
    [checkPendingScroll, clearPendingRefine, pageNavigation, scrollRectIntoView, yrsSession]
  );

  useEffect(() => {
    if (displayListQueries) refinePending(displayListQueries);
  }, [displayListQueries, refinePending]);
  useEffect(() => pageNavigation?.subscribeFrames(refinePending), [pageNavigation, refinePending]);

  // A position the layout does not place yet (still paginating) is scrolled to once it does.
  const pendingPositionRef = useRef<{
    position: number;
    forParaIdScroll: boolean;
    session: YrsSession | null;
    version: string | undefined;
    stop: AbortController;
  } | null>(null);
  const clearPendingPosition = useCallback(() => {
    pendingPositionRef.current?.stop.abort();
    pendingPositionRef.current = null;
  }, []);
  useEffect(() => clearPendingPosition, [clearPendingPosition]);

  const scrollToPositionImpl = useCallback(
    (pmPos: number, forParaIdScroll = false) => {
      clearPendingPosition();
      if (!Number.isInteger(pmPos) || pmPos < 0 || !displayListQueries) return;
      onNavigationIntent?.();
      clearPendingRefine();
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = new AbortController();
      const rect = displayListQueries.anchorRect(pmPos);
      if (rect) {
        scrollAnchorIntoView(displayListQueries, rect, pmPos, !forParaIdScroll);
        return;
      }
      const host = canvasHostRef?.current ?? pagesContainerRef.current;
      const scroller = getScrollContainer() ?? (host ? findVerticalScrollParentOrRoot(host) : null);
      if (!scroller) return;
      const stop = new AbortController();
      const listening = { passive: true, signal: stop.signal };
      for (const type of USER_SCROLL_EVENTS) {
        scroller.addEventListener(type, clearPendingPosition, listening);
      }
      // the editor's input sits outside the scroll container
      scroller.ownerDocument.addEventListener('keydown', clearPendingPosition, {
        ...listening,
        capture: true,
      });
      pendingPositionRef.current = {
        position: pmPos,
        forParaIdScroll,
        session: yrsSession,
        version: yrsSession?.version(),
        stop,
      };
    },
    [
      canvasHostRef,
      clearPendingPosition,
      clearPendingRefine,
      displayListQueries,
      getScrollContainer,
      onNavigationIntent,
      pagesContainerRef,
      scrollAnchorIntoView,
      yrsSession,
    ]
  );

  useEffect(() => {
    const pending = pendingPositionRef.current;
    if (!pending || !displayListQueries) return;
    if (pending.session !== yrsSession || pending.version !== yrsSession?.version()) {
      clearPendingPosition();
      return;
    }
    const rect = displayListQueries.anchorRect(pending.position);
    if (!rect) return;
    clearPendingPosition();
    onNavigationIntent?.();
    scrollAnchorIntoView(displayListQueries, rect, pending.position, !pending.forParaIdScroll);
    if (pendingRefineRef.current) pendingRefineRef.current.version = pending.version;
  }, [
    clearPendingPosition,
    displayListQueries,
    onNavigationIntent,
    scrollAnchorIntoView,
    yrsSession,
  ]);

  const revealPositionImpl = useCallback(
    (position: number, signal?: AbortSignal): RevealPositionOutcome => {
      clearPendingPosition();
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
    [
      clearPendingPosition,
      clearPendingRefine,
      displayListQueries,
      onNavigationIntent,
      scrollAnchorIntoView,
      yrsSession,
    ]
  );

  const pendingPageRef = useRef<{
    page: number;
    epoch: number | undefined;
    session: YrsSession | null;
  } | null>(null);
  const scrollToPageImpl = useCallback(
    (pageNumber: number): void => {
      pendingPageRef.current = null;
      clearPendingPosition();
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
      if (bounds) {
        scrollRectIntoView(bounds, true);
        if (isUnbuiltPage(displayListQueries, pageNumber - 1)) {
          pageNavigation?.buildPages([pageNumber - 1]);
        }
      }
    },
    [
      clearPendingPosition,
      clearPendingRefine,
      displayListQueries,
      layout,
      navigationEpoch,
      onNavigationIntent,
      pageNavigation,
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
