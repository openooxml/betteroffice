/**
 * Canvas-mode selection + caret overlay.
 *
 * The caret / selection rects come from the display list. This component
 * portals a SelectionOverlay onto the visible canvas pages
 * (portal target = `editorContentRef`, the positioned ancestor that shares the
 * `.canvas-pages` host's top-left, exactly like the comment overlays fixed in
 * #92). Reusing SelectionOverlay keeps the caret-blink animation, the
 * `--doc-caret` token, and the selection-highlight color byte-identical to the
 * existing overlay implementation.
 *
 * Worker-presented caret geometry renders synchronously inside page-sized
 * wrappers. DOM-canvas fallback geometry keeps its live-canvas projection.
 *
 * Renders nothing on the default DOM-painter path: PagedEditor only mounts this
 * (in place of its inline overlay) once `overlayTarget` resolves, which
 * `useCanvasOverlayTarget` returns non-null only while the canvas paints.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import type { SelectionRect, CaretPosition } from '@betteroffice/docx/layout';
import {
  CANVAS_PAGE_GAP_PX,
  CANVAS_PAGES_PADDING_PX,
  type DisplayList,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import { SIDEBAR_DOCUMENT_SHIFT } from '../../sidebar/constants';
import { SelectionOverlay } from './SelectionOverlay';

export interface CanvasSelectionOverlayProps {
  /** Selection rects in page-local px + pageIndex (from the display list). */
  selectionRects: SelectionRect[];
  /** Caret position in page-local px + pageIndex, or null for a range. */
  caretPosition: CaretPosition | null;
  /** Whether the (hidden body) ProseMirror is focused — drives the blink. */
  isFocused: boolean;
  /** Hide the caret / selection in read-only mode. */
  readOnly?: boolean;
  /**
   * Portal target — `editorContentRef.current`, the positioned ancestor that
   * shares the `.canvas-pages` host's top-left.
   */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host used by the DOM-canvas projection path. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  displayList: DisplayList;
  displayListQueries?: DisplayListQueries | null;
  /** The worker has already presented the frame owning these rects. */
  directProjection: boolean;
  /** Sidebar open state. */
  sidebarOpen: boolean;
  /** Current page zoom. */
  zoom: number;
}

export function CanvasSelectionOverlay({
  selectionRects,
  caretPosition,
  isFocused,
  readOnly = false,
  overlayTarget,
  canvasHostRef,
  displayList,
  displayListQueries = null,
  directProjection,
  sidebarOpen,
  zoom,
}: CanvasSelectionOverlayProps) {
  if (directProjection) {
    return (
      <DirectCanvasSelectionOverlay
        selectionRects={selectionRects}
        caretPosition={caretPosition}
        isFocused={isFocused}
        readOnly={readOnly}
        overlayTarget={overlayTarget}
        displayList={displayList}
        sidebarOpen={sidebarOpen}
        zoom={zoom}
      />
    );
  }
  if (!displayListQueries) return null;
  return (
    <ProjectedCanvasSelectionOverlay
      selectionRects={selectionRects}
      caretPosition={caretPosition}
      isFocused={isFocused}
      readOnly={readOnly}
      overlayTarget={overlayTarget}
      canvasHostRef={canvasHostRef}
      displayListQueries={displayListQueries}
      sidebarOpen={sidebarOpen}
      zoom={zoom}
    />
  );
}

function DirectCanvasSelectionOverlay({
  selectionRects,
  caretPosition,
  isFocused,
  readOnly,
  overlayTarget,
  displayList,
  sidebarOpen,
  zoom,
}: Omit<CanvasSelectionOverlayProps, 'canvasHostRef' | 'displayListQueries' | 'directProjection'>) {
  const rectsByPage = new Map<number, SelectionRect[]>();
  for (const rect of selectionRects) {
    const pageRects = rectsByPage.get(rect.pageIndex) ?? [];
    pageRects.push({
      ...rect,
      x: rect.x * zoom,
      y: rect.y * zoom,
      width: rect.width * zoom,
      height: rect.height * zoom,
    });
    rectsByPage.set(rect.pageIndex, pageRects);
  }

  let pageTop = CANVAS_PAGES_PADDING_PX;
  const overlays = displayList.pages.flatMap((page) => {
    const top = pageTop;
    pageTop += page.height * zoom + CANVAS_PAGE_GAP_PX;
    const rects = rectsByPage.get(page.pageIndex) ?? [];
    const caret =
      caretPosition?.pageIndex === page.pageIndex
        ? {
            ...caretPosition,
            x: caretPosition.x * zoom,
            y: caretPosition.y * zoom,
            height: caretPosition.height * zoom,
          }
        : null;
    if (rects.length === 0 && !caret) return [];
    return [
      <div
        key={page.pageIndex}
        style={{
          position: 'absolute',
          top,
          left: '50%',
          width: page.width * zoom,
          height: page.height * zoom,
          transform: `translateX(-50%)${sidebarOpen ? ` translateX(-${SIDEBAR_DOCUMENT_SHIFT}px)` : ''}`,
          transition: 'transform 0.2s ease',
          pointerEvents: 'none',
        }}
      >
        <SelectionOverlay
          selectionRects={rects}
          caretPosition={caret}
          isFocused={isFocused}
          readOnly={readOnly}
        />
      </div>,
    ];
  });

  return createPortal(overlays, overlayTarget);
}

function ProjectedCanvasSelectionOverlay({
  selectionRects,
  caretPosition,
  isFocused,
  readOnly,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
}: Omit<CanvasSelectionOverlayProps, 'displayList' | 'directProjection'> & {
  displayListQueries: DisplayListQueries;
}) {
  const [converted, setConverted] = useState<{
    rects: SelectionRect[];
    caret: CaretPosition | null;
  }>({ rects: [], caret: null });

  useLayoutEffect(() => {
    const host = canvasHostRef.current;
    if (!host) {
      setConverted({ rects: [], caret: null });
      return;
    }

    const recompute = () => {
      const targetRect = overlayTarget.getBoundingClientRect();
      const project = (pageIndex: number, x: number, y: number) => {
        const canvasEl = host.querySelector<HTMLCanvasElement>(
          `canvas[data-page-index="${pageIndex}"]`
        );
        const size = displayListQueries.pageSize(pageIndex);
        if (!canvasEl || !size) return null;
        const canvasRect = canvasEl.getBoundingClientRect();
        const scaleX = size.width > 0 ? canvasRect.width / size.width : 1;
        const scaleY = size.height > 0 ? canvasRect.height / size.height : 1;
        return {
          left: canvasRect.left - targetRect.left + x * scaleX,
          top: canvasRect.top - targetRect.top + y * scaleY,
          scaleX,
          scaleY,
        };
      };

      const nextRects: SelectionRect[] = [];
      for (const rect of selectionRects) {
        const projected = project(rect.pageIndex, rect.x, rect.y);
        if (!projected) continue;
        nextRects.push({
          x: projected.left,
          y: projected.top,
          width: rect.width * projected.scaleX,
          height: rect.height * projected.scaleY,
          pageIndex: rect.pageIndex,
        });
      }

      let nextCaret: CaretPosition | null = null;
      if (caretPosition) {
        const projected = project(caretPosition.pageIndex, caretPosition.x, caretPosition.y);
        if (projected) {
          nextCaret = {
            x: projected.left,
            y: projected.top,
            height: caretPosition.height * projected.scaleY,
            pageIndex: caretPosition.pageIndex,
          };
        }
      }

      setConverted({ rects: nextRects, caret: nextCaret });
    };

    recompute();
    const resizeObserver = new ResizeObserver(recompute);
    resizeObserver.observe(host);
    resizeObserver.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);
    return () => {
      resizeObserver.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
    };
  }, [
    selectionRects,
    caretPosition,
    overlayTarget,
    canvasHostRef,
    displayListQueries,
    sidebarOpen,
    zoom,
  ]);

  return createPortal(
    <SelectionOverlay
      selectionRects={converted.rects}
      caretPosition={converted.caret}
      isFocused={isFocused}
      readOnly={readOnly}
    />,
    overlayTarget
  );
}

export default CanvasSelectionOverlay;
