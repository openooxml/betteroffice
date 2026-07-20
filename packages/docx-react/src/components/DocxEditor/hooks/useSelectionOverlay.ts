/** Display-list-backed selection state for PagedEditor overlays. */

import { useCallback, useEffect, useState } from 'react';
import type { CaretPosition, SelectionRect } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsResidentCaretSnapshot } from '@betteroffice/docx/yrs';

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
  } = opts;

  const [selectionRects, setQueriedSelectionRects] = useState<SelectionRect[]>([]);
  const [queriedCaretPosition, setQueriedCaretPosition] = useState<CaretPosition | null>(null);
  const [sampledFrameEpoch, setSampledFrameEpoch] = useState<number | null>(null);
  const setSelectionRects: React.Dispatch<React.SetStateAction<SelectionRect[]>> = useCallback(
    (next) => {
      setSampledFrameEpoch(displayListFrameEpoch);
      setQueriedSelectionRects(next);
    },
    [displayListFrameEpoch]
  );
  const setCaretPosition: React.Dispatch<React.SetStateAction<CaretPosition | null>> = useCallback(
    (next) => {
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
      displayListQueries,
      getYrsDisplaySelection,
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
