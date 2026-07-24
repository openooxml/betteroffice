/** Display-list-backed scroll/ref API for PagedEditor. */

import { useCallback, useEffect, useRef } from 'react';
import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import type { ParagraphHighlightOptions, ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';

import type { YrsInputRef } from '../YrsInput';
import { runAfterFrames } from '../internals/scrollUtils';

export interface UsePagedScrollApiOptions {
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  yrsSession: YrsSession | null;
  yrsLocToDisplayPosition: (loc: YrsLoc) => number | null;
  getScrollContainer: () => HTMLDivElement | null;
  displayListQueries?: DisplayListQueries | null;
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  onNavigationIntent?: () => void;
  requestCanvasParagraphFlash?: (req: {
    from: number;
    to: number;
    options?: ParagraphHighlightOptions;
  }) => void;
}

export interface UsePagedScrollApiReturn {
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
  scrollToPageImpl: (pageNumber: number) => void;
  scrollToParaIdImpl: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
}

export function usePagedScrollApi(opts: UsePagedScrollApiOptions): UsePagedScrollApiReturn {
  const {
    pagesContainerRef,
    yrsInputRef,
    yrsSession,
    yrsLocToDisplayPosition,
    getScrollContainer,
    displayListQueries = null,
    canvasHostRef,
    onNavigationIntent,
    requestCanvasParagraphFlash,
  } = opts;
  const scrollAbortRef = useRef<AbortController | null>(null);

  useEffect(
    () => () => {
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = null;
    },
    []
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
      const scrollerRect = scroller.getBoundingClientRect();
      const scaleY = pageSize.height > 0 ? pageRect.height / pageSize.height : 1;
      const clientY = pageRect.top + (rect.y + rect.height / 2) * scaleY;
      scroller.scrollTo({
        top: scroller.scrollTop + clientY - scrollerRect.top - scroller.clientHeight / 2,
        behavior: smooth ? 'smooth' : 'auto',
      });
      return true;
    },
    [canvasHostRef, displayListQueries, getScrollContainer, pagesContainerRef]
  );

  const scrollToPositionImpl = useCallback(
    (pmPos: number, forParaIdScroll = false) => {
      if (!Number.isInteger(pmPos) || pmPos < 0 || !displayListQueries) return;
      onNavigationIntent?.();
      scrollAbortRef.current?.abort();
      scrollAbortRef.current = new AbortController();
      const rect = displayListQueries.anchorRect(pmPos);
      if (rect) scrollRectIntoView(rect, !forParaIdScroll);
    },
    [displayListQueries, onNavigationIntent, scrollRectIntoView]
  );

  const scrollToPageImpl = useCallback(
    (pageNumber: number): void => {
      if (
        !Number.isInteger(pageNumber) ||
        pageNumber < 1 ||
        !displayListQueries ||
        pageNumber > displayListQueries.pageCount()
      ) {
        return;
      }
      onNavigationIntent?.();
      const bounds = displayListQueries.pageBounds(pageNumber - 1);
      if (bounds) scrollRectIntoView(bounds, true);
    },
    [displayListQueries, onNavigationIntent, scrollRectIntoView]
  );

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
      runAfterFrames(() => {
        yrsSession.setSelection(startLoc);
        yrsInputRef.current?.focus();
      }, signal);
      return true;
    },
    [requestCanvasParagraphFlash, scrollToPositionImpl, yrsInputRef, yrsLocToDisplayPosition, yrsSession]
  );

  return { scrollToPositionImpl, scrollToPageImpl, scrollToParaIdImpl };
}
