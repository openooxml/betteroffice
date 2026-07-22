/** Display-list-backed selection state for PagedEditor overlays. */

import { useCallback, useEffect, useRef, useState } from 'react';
import type { CaretPosition, SelectionRect } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import {
  sameYrsSelection,
  type YrsResidentCaretSnapshot,
  type YrsSelection,
} from '@betteroffice/docx/yrs';

import type { YrsDisplaySelection } from '../YrsInput';
import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';

export interface UseSelectionOverlayOptions {
  layout: Layout | null;
  containerRef: React.RefObject<HTMLDivElement | null>;
  syncCoordinator: LayoutSelectionGate;
  displayListQueries?: DisplayListQueries | null;
  displayListFrameEpoch?: number | null;
  residentCaret?: YrsResidentCaretSnapshot | null;
  residentCaretAuthoritative?: boolean;
  getYrsDisplaySelection: () => YrsDisplaySelection | null;
  /** Live sticky selection, read at call time to validate the worker caret. */
  getYrsStickySelection?: () => YrsSelection | null;
}

export interface UseSelectionOverlayReturn {
  selectionRects: SelectionRect[];
  caretPosition: CaretPosition | null;
  setSelectionRects: React.Dispatch<React.SetStateAction<SelectionRect[]>>;
  setCaretPosition: React.Dispatch<React.SetStateAction<CaretPosition | null>>;
  updateSelectionOverlay: () => void;
}

export function useSelectionOverlay(opts: UseSelectionOverlayOptions): UseSelectionOverlayReturn {
  const {
    layout,
    containerRef,
    syncCoordinator,
    displayListQueries,
    displayListFrameEpoch = null,
    residentCaret = null,
    residentCaretAuthoritative = false,
    getYrsDisplaySelection,
    getYrsStickySelection,
  } = opts;

  const [selectionRects, setQueriedSelectionRects] = useState<SelectionRect[]>([]);
  const [queriedCaretPosition, setQueriedCaretPosition] = useState<CaretPosition | null>(null);
  const [sampledFrameEpoch, setSampledFrameEpoch] = useState<number | null>(null);
  const sampledFrameEpochRef = useRef<number | null>(null);
  const setSelectionRects: React.Dispatch<React.SetStateAction<SelectionRect[]>> = useCallback(
    (next) => {
      sampledFrameEpochRef.current = displayListFrameEpoch;
      setSampledFrameEpoch(displayListFrameEpoch);
      setQueriedSelectionRects(next);
    },
    [displayListFrameEpoch]
  );
  const setCaretPosition: React.Dispatch<React.SetStateAction<CaretPosition | null>> = useCallback(
    (next) => {
      sampledFrameEpochRef.current = displayListFrameEpoch;
      setSampledFrameEpoch(displayListFrameEpoch);
      setQueriedCaretPosition(next);
    },
    [displayListFrameEpoch]
  );

  const updateSelectionOverlay = useCallback(
    () => {
      const yrsSelection = getYrsDisplaySelection();
      if (!yrsSelection) return;
      const anchor = yrsSelection.anchor;
      const head = yrsSelection.head;
      const from = Math.min(anchor, head);
      const to = Math.max(anchor, head);
      // Worker-computed caret for this exact frame and selection: no facade
      // query is needed. When the epoch arbitration below already renders it,
      // skip state churn entirely; otherwise publish it as the queried rect.
      if (
        from === to &&
        residentCaretAuthoritative &&
        residentCaret?.caretRect &&
        residentCaret.frameEpoch === displayListFrameEpoch &&
        residentCaret.selection &&
        sameYrsSelection(residentCaret.selection, getYrsStickySelection?.() ?? null)
      ) {
        const rect = residentCaret.caretRect;
        const sampled = sampledFrameEpochRef.current;
        if (sampled === null || sampled < residentCaret.frameEpoch) return;
        setCaretPosition((current) =>
          current?.x === rect.x &&
          current?.y === rect.y &&
          current?.height === rect.height &&
          current?.pageIndex === rect.pageIndex
            ? current
            : { x: rect.x, y: rect.y, height: rect.height, pageIndex: rect.pageIndex }
        );
        setSelectionRects((current) => (current.length === 0 ? current : []));
        return;
      }
      if (!displayListQueries) {
        if (residentCaretAuthoritative) return;
        setCaretPosition((current) => (current === null ? current : null));
        setSelectionRects((current) => (current.length === 0 ? current : []));
        return;
      }
      if (!displayListQueries.isReady()) return;
      if (from === to) {
        const rect = displayListQueries.caretRect(head);
        const next = rect
          ? { x: rect.x, y: rect.y, height: rect.height, pageIndex: rect.pageIndex }
          : null;
        setCaretPosition((current) =>
          current?.x === next?.x &&
          current?.y === next?.y &&
          current?.height === next?.height &&
          current?.pageIndex === next?.pageIndex
            ? current
            : next
        );
        setSelectionRects((current) => (current.length === 0 ? current : []));
      } else {
        const next = displayListQueries.rangeRects(from, to).map((rect) => ({
            x: rect.x,
            y: rect.y,
            width: rect.width,
            height: rect.height,
            pageIndex: rect.pageIndex,
          }));
        setSelectionRects((current) =>
          current.length === next.length &&
          current.every(
            (rect, index) =>
              rect.x === next[index]?.x &&
              rect.y === next[index]?.y &&
              rect.width === next[index]?.width &&
              rect.height === next[index]?.height &&
              rect.pageIndex === next[index]?.pageIndex
          )
            ? current
            : next
        );
        setCaretPosition((current) => (current === null ? current : null));
      }
    },
    [
      containerRef,
      displayListFrameEpoch,
      displayListQueries,
      getYrsDisplaySelection,
      getYrsStickySelection,
      residentCaret,
      residentCaretAuthoritative,
      setCaretPosition,
      setSelectionRects,
    ]
  );

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    const observer = new ResizeObserver(() => {
      if (syncCoordinator.isSafeToRender()) updateSelectionOverlay();
    });
    observer.observe(container);
    return () => observer.disconnect();
  }, [containerRef, syncCoordinator, updateSelectionOverlay]);

  useEffect(() => {
    if (layout) updateSelectionOverlay();
  }, [layout, updateSelectionOverlay]);

  const authoritativeRect = residentCaretAuthoritative ? residentCaret?.caretRect : null;
  const authoritativeNewer = Boolean(
    authoritativeRect &&
      residentCaret &&
      (sampledFrameEpoch === null || sampledFrameEpoch < residentCaret.frameEpoch)
  );
  const caretPosition =
    authoritativeRect && authoritativeNewer
      ? {
          x: authoritativeRect.x,
          y: authoritativeRect.y,
          height: authoritativeRect.height,
          pageIndex: authoritativeRect.pageIndex,
        }
      : queriedCaretPosition;

  return {
    selectionRects: authoritativeNewer ? [] : selectionRects,
    caretPosition,
    setSelectionRects,
    setCaretPosition,
    updateSelectionOverlay,
  };
}
