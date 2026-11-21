import { useEffect, useRef } from 'react';
import type { RefObject } from 'react';
import { MAX_ZOOM, MIN_ZOOM } from './commands/evaluate';

const MAX_WHEEL_DELTA = 50;
const WHEEL_ZOOM_SENSITIVITY = 0.005;

export function pinchZoomScale(
  scale: number,
  deltaY: number,
  deltaMode: number,
  viewportHeight: number
): number {
  const unit = deltaMode === WheelEvent.DOM_DELTA_LINE
    ? 16
    : deltaMode === WheelEvent.DOM_DELTA_PAGE
      ? viewportHeight
      : 1;
  const delta = Math.max(-MAX_WHEEL_DELTA, Math.min(MAX_WHEEL_DELTA, deltaY * unit));
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, scale * Math.exp(-delta * WHEEL_ZOOM_SENSITIVITY)));
}

/**
 * Converts the browser's Ctrl+wheel trackpad pinch gesture into editor zoom while
 * keeping the point below the gesture stationary in the scroll viewport.
 */
export function usePptxPinchZoom(
  viewportRef: RefObject<HTMLElement | null>,
  contentRef: RefObject<HTMLElement | null>,
  scale: number,
  setScale: (scale: number) => void,
  presentation: unknown
) {
  const scaleRef = useRef(scale);
  const setScaleRef = useRef(setScale);
  const commitAnimationRef = useRef<number | null>(null);
  const anchorAnimationRef = useRef<number | null>(null);
  const pendingAnchorRef = useRef<{
    width: number;
    height: number;
    baseWidth: number;
    baseHeight: number;
    anchorX: number;
    anchorY: number;
    clientX: number;
    clientY: number;
  } | null>(null);
  scaleRef.current = scale;
  setScaleRef.current = setScale;

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const view = viewport.ownerDocument.defaultView;
    const applyAnchor = () => {
      const pending = pendingAnchorRef.current;
      const content = contentRef.current;
      if (!pending || !content) return;
      const after = content.getBoundingClientRect();
      if (
        Math.abs(after.width - pending.width) >= 0.5 ||
        Math.abs(after.height - pending.height) >= 0.5
      ) {
        return;
      }
      pendingAnchorRef.current = null;
      viewport.scrollLeft += after.left + pending.anchorX * after.width - pending.clientX;
      viewport.scrollTop += after.top + pending.anchorY * after.height - pending.clientY;
      observer?.disconnect();
    };
    const observer =
      typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(applyAnchor);
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey || event.metaKey || event.altKey || !event.cancelable) return;
      const content = contentRef.current;
      if (!content) return;
      event.preventDefault();
      const next = pinchZoomScale(
        scaleRef.current,
        event.deltaY,
        event.deltaMode,
        viewport.clientHeight
      );
      if (next === scaleRef.current) return;
      const before = content.getBoundingClientRect();
      const currentScale = scaleRef.current;
      const pending = pendingAnchorRef.current;
      const baseWidth = pending?.baseWidth ??
        (currentScale > 0 ? before.width / currentScale : before.width);
      const baseHeight = pending?.baseHeight ??
        (currentScale > 0 ? before.height / currentScale : before.height);
      const anchorX = pending?.anchorX ??
        (before.width > 0 ? (event.clientX - before.left) / before.width : 0.5);
      const anchorY = pending?.anchorY ??
        (before.height > 0 ? (event.clientY - before.top) / before.height : 0.5);
      const nextWidth = baseWidth * next;
      const nextHeight = baseHeight * next;
      pendingAnchorRef.current = {
        width: nextWidth,
        height: nextHeight,
        baseWidth,
        baseHeight,
        anchorX,
        anchorY,
        clientX: event.clientX,
        clientY: event.clientY,
      };
      observer?.disconnect();
      observer?.observe(content);
      scaleRef.current = next;
      if (!view) {
        setScaleRef.current(next);
        return;
      }
      if (commitAnimationRef.current === null) {
        commitAnimationRef.current = view.requestAnimationFrame(() => {
          commitAnimationRef.current = null;
          setScaleRef.current(scaleRef.current);
          if (anchorAnimationRef.current !== null) {
            view.cancelAnimationFrame(anchorAnimationRef.current);
          }
          anchorAnimationRef.current = view.requestAnimationFrame(() => {
            anchorAnimationRef.current = null;
            applyAnchor();
          });
        });
      }
    };
    viewport.addEventListener('wheel', onWheel, { passive: false });
    return () => {
      viewport.removeEventListener('wheel', onWheel);
      observer?.disconnect();
      pendingAnchorRef.current = null;
      if (view && commitAnimationRef.current !== null) {
        view.cancelAnimationFrame(commitAnimationRef.current);
        commitAnimationRef.current = null;
      }
      if (view && anchorAnimationRef.current !== null) {
        view.cancelAnimationFrame(anchorAnimationRef.current);
        anchorAnimationRef.current = null;
      }
    };
  }, [contentRef, presentation, viewportRef]);
}
