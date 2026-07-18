/**
 * Canvas-mode image selection + resize/drag overlay.
 *
 * This component sources the selected image's rect from the display list's
 * `ImagePrimitive` (found by the selection's PM
 * position) and projects it onto the visible canvas pages through the live
 * per-page `<canvas>` rect — the same conversion `CanvasSelectionOverlay` uses
 * for the caret. It then portals the border + handles + drag affordance onto
 * the canvas (portal target = `editorContentRef`, which shares the
 * `.canvas-pages` host's top-left).
 *
 * Resize math is the shared core `calculateResizedImageDimensions`, and the
 * resize / drag COMMITS route through the exact same callbacks
 * (`useImageInteractions` → core `imageCommit`) the DOM overlay uses. Renders
 * nothing off the canvas path (the parent only mounts it once
 * `overlayTarget` + `displayListQueries` resolve, i.e. while the canvas paints).
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

import React, { useCallback, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import {
  calculateResizedImageDimensions,
  type ImageResizeHandle,
} from '@betteroffice/docx/docx';
import { findImagePrimitiveByDocPos, type DisplayListQueries } from '@betteroffice/docx/layout/render';
import { canvasPageScale } from '../internals/canvasProjection';

const HANDLE_SIZE = 10;
const HANDLE_HALF = HANDLE_SIZE / 2;
const BORDER_WIDTH = 2;
const ACCENT_COLOR = '#2563eb'; // Blue-600 — matches the DOM ImageSelectionOverlay

const HANDLE_CURSORS: Record<ImageResizeHandle, string> = {
  nw: 'nw-resize',
  ne: 'ne-resize',
  se: 'se-resize',
  sw: 'sw-resize',
  n: 'ns-resize',
  s: 'ns-resize',
  e: 'ew-resize',
  w: 'ew-resize',
};

// x/y are fractions of the box: 0 = start edge, 0.5 = midpoint, 1 = end edge.
const HANDLES: ReadonlyArray<{ pos: ImageResizeHandle; x: number; y: number }> = [
  { pos: 'nw', x: 0, y: 0 },
  { pos: 'ne', x: 1, y: 0 },
  { pos: 'se', x: 1, y: 1 },
  { pos: 'sw', x: 0, y: 1 },
  { pos: 'n', x: 0.5, y: 0 },
  { pos: 's', x: 0.5, y: 1 },
  { pos: 'e', x: 1, y: 0.5 },
  { pos: 'w', x: 0, y: 0.5 },
];

export interface CanvasImageSelectionOverlayProps {
  /** PM position of the selected image, or null when no image is selected. */
  pmPos: number | null;
  /** Whether the (hidden body) ProseMirror is focused — gates the overlay. */
  isFocused: boolean;
  readOnly?: boolean;
  /** Portal target — `editorContentRef.current`, the positioned canvas ancestor. */
  overlayTarget: HTMLElement;
  /** `.canvas-pages` host — live per-page `<canvas>` rects are read from here. */
  canvasHostRef: React.RefObject<HTMLDivElement | null>;
  /** Display-list queries — carries the display list + page sizes. */
  displayListQueries: DisplayListQueries;
  /** Sidebar open — re-project after its `translateX` transition settles. */
  sidebarOpen: boolean;
  /** Zoom — re-project when page geometry scales. */
  zoom: number;
  /** Commit callbacks (shared with the DOM path via useImageInteractions). */
  onResize: (pmPos: number, newWidth: number, newHeight: number) => void;
  onResizeStart?: () => void;
  onResizeEnd?: () => void;
  onDragMove?: (pmPos: number, clientX: number, clientY: number) => void;
  onDragStart?: () => void;
  onDragEnd?: () => void;
  onContextMenu?: (e: React.MouseEvent) => void;
}

interface Projected {
  pageIndex: number;
  /** document-px (page-local) image box from the display-list primitive */
  doc: { x: number; y: number; w: number; h: number };
  /** overlay-target-px box (rendered geometry) */
  rect: { left: number; top: number; width: number; height: number };
  scaleX: number;
  scaleY: number;
}

export function CanvasImageSelectionOverlay({
  pmPos,
  isFocused,
  readOnly = false,
  overlayTarget,
  canvasHostRef,
  displayListQueries,
  sidebarOpen,
  zoom,
  onResize,
  onResizeStart,
  onResizeEnd,
  onDragMove,
  onDragStart,
  onDragEnd,
  onContextMenu,
}: CanvasImageSelectionOverlayProps): React.ReactPortal | null {
  const [base, setBase] = useState<Projected | null>(null);
  // While a gesture is live, `preview` overrides the projected rect (resize
  // live-preview); re-projection from listeners is frozen so it can't clobber it.
  const [preview, setPreview] = useState<{
    left: number;
    top: number;
    width: number;
    height: number;
  } | null>(null);
  const [dims, setDims] = useState<{ w: number; h: number } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const gestureActiveRef = useRef(false);

  // Callback + geometry refs so imperative window handlers see latest values.
  const onResizeRef = useRef(onResize);
  const onResizeStartRef = useRef(onResizeStart);
  const onResizeEndRef = useRef(onResizeEnd);
  const onDragMoveRef = useRef(onDragMove);
  const onDragStartRef = useRef(onDragStart);
  const onDragEndRef = useRef(onDragEnd);
  onResizeRef.current = onResize;
  onResizeStartRef.current = onResizeStart;
  onResizeEndRef.current = onResizeEnd;
  onDragMoveRef.current = onDragMove;
  onDragStartRef.current = onDragStart;
  onDragEndRef.current = onDragEnd;
  const pmPosRef = useRef(pmPos);
  pmPosRef.current = pmPos;

  const project = useCallback((): Projected | null => {
    const host = canvasHostRef.current;
    if (host == null || pmPos == null) return null;
    const located = findImagePrimitiveByDocPos(displayListQueries.displayList, pmPos);
    if (!located) return null;
    const { primitive, pageIndex } = located;
    const scale = canvasPageScale(host, displayListQueries, pageIndex);
    if (!scale) return null;
    const targetRect = overlayTarget.getBoundingClientRect();
    const { canvasRect, scaleX, scaleY } = scale;
    return {
      pageIndex,
      doc: { x: primitive.x, y: primitive.y, w: primitive.w, h: primitive.h },
      rect: {
        left: canvasRect.left - targetRect.left + primitive.x * scaleX,
        top: canvasRect.top - targetRect.top + primitive.y * scaleY,
        width: primitive.w * scaleX,
        height: primitive.h * scaleY,
      },
      scaleX,
      scaleY,
    };
  }, [canvasHostRef, displayListQueries, overlayTarget, pmPos]);

  useLayoutEffect(() => {
    if (gestureActiveRef.current) return; // don't fight an in-flight preview
    setBase(project());
    setPreview(null);
    const host = canvasHostRef.current;
    if (!host) return;
    const recompute = () => {
      if (gestureActiveRef.current) return;
      setBase(project());
    };
    const ro = new ResizeObserver(recompute);
    ro.observe(host);
    ro.observe(overlayTarget);
    window.addEventListener('resize', recompute);
    host.addEventListener('transitionend', recompute);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', recompute);
      host.removeEventListener('transitionend', recompute);
    };
  }, [project, canvasHostRef, overlayTarget, sidebarOpen, zoom]);

  const handleResizeStart = useCallback(
    (handle: ImageResizeHandle, e: React.MouseEvent) => {
      if (!base) return;
      e.preventDefault();
      e.stopPropagation();

      const { doc, rect, scaleX, scaleY } = base;
      const startX = e.clientX;
      const startY = e.clientY;
      const startWidthDoc = doc.w;
      const startHeightDoc = doc.h;
      let finalWidth = Math.round(startWidthDoc);
      let finalHeight = Math.round(startHeightDoc);

      gestureActiveRef.current = true;
      setDims({ w: finalWidth, h: finalHeight });
      onResizeStartRef.current?.();

      const onMove = (moveEvent: MouseEvent) => {
        // client px → document px through the live canvas scale (== zoom).
        const deltaX = scaleX > 0 ? (moveEvent.clientX - startX) / scaleX : 0;
        const deltaY = scaleY > 0 ? (moveEvent.clientY - startY) / scaleY : 0;
        const lockAspect = !moveEvent.shiftKey;
        const d = calculateResizedImageDimensions(
          handle,
          deltaX,
          deltaY,
          startWidthDoc,
          startHeightDoc,
          lockAspect
        );
        finalWidth = Math.round(d.width);
        finalHeight = Math.round(d.height);
        setDims({ w: finalWidth, h: finalHeight });
        // w/n handles anchor the opposite edge, so the box origin shifts.
        const newDocX = handle.includes('w') ? doc.x + (doc.w - d.width) : doc.x;
        const newDocY = handle.includes('n') ? doc.y + (doc.h - d.height) : doc.y;
        setPreview({
          left: rect.left + (newDocX - doc.x) * scaleX,
          top: rect.top + (newDocY - doc.y) * scaleY,
          width: d.width * scaleX,
          height: d.height * scaleY,
        });
      };
      const onUp = () => {
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
        gestureActiveRef.current = false;
        setPreview(null);
        setDims(null);
        const pos = pmPosRef.current;
        if (pos != null) onResizeRef.current?.(pos, finalWidth, finalHeight);
        onResizeEndRef.current?.();
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    },
    [base]
  );

  const handleBodyMouseDown = useCallback(
    (e: React.MouseEvent) => {
      if (!base) return;
      e.preventDefault();
      e.stopPropagation();
      const DRAG_THRESHOLD = 4;
      const startX = e.clientX;
      const startY = e.clientY;
      const ghostW = base.rect.width;
      const ghostH = base.rect.height;
      let dragStarted = false;
      let ghostEl: HTMLElement | null = null;

      const onMove = (moveEvent: MouseEvent) => {
        const dx = moveEvent.clientX - startX;
        const dy = moveEvent.clientY - startY;
        if (!dragStarted && Math.sqrt(dx * dx + dy * dy) < DRAG_THRESHOLD) return;
        if (!dragStarted) {
          dragStarted = true;
          gestureActiveRef.current = true;
          setIsDragging(true);
          onDragStartRef.current?.();
          ghostEl = document.createElement('div');
          ghostEl.style.cssText =
            'position: fixed; pointer-events: none; z-index: 10000; ' +
            'opacity: 0.5; border: 2px dashed #2563eb; border-radius: 4px; ' +
            'background: rgba(37, 99, 235, 0.1);';
          ghostEl.style.width = `${ghostW}px`;
          ghostEl.style.height = `${ghostH}px`;
          document.body.appendChild(ghostEl);
        }
        if (ghostEl) {
          ghostEl.style.left = `${moveEvent.clientX - ghostW / 2}px`;
          ghostEl.style.top = `${moveEvent.clientY - ghostH / 2}px`;
        }
      };
      const onUp = (upEvent: MouseEvent) => {
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
        ghostEl?.remove();
        ghostEl = null;
        gestureActiveRef.current = false;
        setIsDragging(false);
        if (dragStarted) {
          const pos = pmPosRef.current;
          if (pos != null) onDragMoveRef.current?.(pos, upEvent.clientX, upEvent.clientY);
          onDragEndRef.current?.();
        }
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    },
    [base]
  );

  if (!base || !isFocused) return null;
  const r = preview ?? base.rect;

  const overlayStyles: CSSProperties = {
    position: 'absolute',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    pointerEvents: 'none',
    zIndex: 15,
    overflow: 'visible',
  };
  const handleBase: CSSProperties = {
    position: 'absolute',
    width: HANDLE_SIZE,
    height: HANDLE_SIZE,
    backgroundColor: '#ffffff',
    border: `1.5px solid ${ACCENT_COLOR}`,
    borderRadius: '50%',
    boxShadow: '0 1px 2.5px rgba(0, 0, 0, 0.35)',
    boxSizing: 'border-box',
    pointerEvents: readOnly ? 'none' : 'auto',
    zIndex: 16,
  };

  return createPortal(
    <div style={overlayStyles} className="image-selection-overlay canvas-image-selection-overlay">
      <div
        style={{
          position: 'absolute',
          left: r.left - BORDER_WIDTH,
          top: r.top - BORDER_WIDTH,
          width: r.width + BORDER_WIDTH * 2,
          height: r.height + BORDER_WIDTH * 2,
          border: `${BORDER_WIDTH}px solid ${ACCENT_COLOR}`,
          pointerEvents: 'none',
          boxSizing: 'border-box',
        }}
      />
      <div
        style={{
          position: 'absolute',
          left: r.left,
          top: r.top,
          width: r.width,
          height: r.height,
          cursor: readOnly ? 'default' : isDragging ? 'grabbing' : 'grab',
          pointerEvents: readOnly ? 'none' : 'auto',
          zIndex: 15,
        }}
        onMouseDown={readOnly ? undefined : handleBodyMouseDown}
        onContextMenu={onContextMenu}
      />
      {HANDLES.map(({ pos, x, y }) => (
        <div
          key={pos}
          data-handle={pos}
          style={{
            ...handleBase,
            left: r.left + r.width * x - HANDLE_HALF,
            top: r.top + r.height * y - HANDLE_HALF,
            cursor: HANDLE_CURSORS[pos],
          }}
          onMouseDown={readOnly ? undefined : (e) => handleResizeStart(pos, e)}
        />
      ))}
      {dims && (
        <div
          style={{
            position: 'absolute',
            left: r.left + r.width / 2,
            top: r.top + r.height + 12,
            transform: 'translateX(-50%)',
            backgroundColor: 'rgba(0, 0, 0, 0.75)',
            color: 'var(--doc-on-primary)',
            fontSize: 11,
            fontFamily: 'system-ui, sans-serif',
            padding: '2px 8px',
            borderRadius: 3,
            whiteSpace: 'nowrap',
            pointerEvents: 'none',
            zIndex: 20,
          }}
        >
          {dims.w} × {dims.h}
        </div>
      )}
    </div>,
    overlayTarget
  );
}

export default CanvasImageSelectionOverlay;
