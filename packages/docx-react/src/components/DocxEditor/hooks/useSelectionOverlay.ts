/** Display-list-backed selection state for PagedEditor overlays. */

import { useCallback, useEffect, useState } from 'react';
import type { CaretPosition, SelectionRect } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';

import type { YrsDisplaySelection } from '../YrsInput';
import type { LayoutSelectionGate } from '../internals/LayoutSelectionGate';

export interface UseSelectionOverlayOptions {
  layout: Layout | null;
  containerRef: React.RefObject<HTMLDivElement | null>;
  syncCoordinator: LayoutSelectionGate;
  displayListQueries?: DisplayListQueries | null;
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
    getYrsDisplaySelection,
  } = opts;

  const [selectionRects, setSelectionRects] = useState<SelectionRect[]>([]);
  const [caretPosition, setCaretPosition] = useState<CaretPosition | null>(null);

  const updateSelectionOverlay = useCallback(
    () => {
      const yrsSelection = getYrsDisplaySelection();
      if (!yrsSelection) return;
      const anchor = yrsSelection.anchor;
      const head = yrsSelection.head;
      const from = Math.min(anchor, head);
      const to = Math.max(anchor, head);
      if (!displayListQueries) {
        setCaretPosition((current) => (current === null ? current : null));
        setSelectionRects((current) => (current.length === 0 ? current : []));
        return;
      }
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

  return {
    selectionRects,
    caretPosition,
    setSelectionRects,
    setCaretPosition,
    updateSelectionOverlay,
  };
}
