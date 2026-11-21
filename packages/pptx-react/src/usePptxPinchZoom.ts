import { useEffect, useLayoutEffect, useRef } from 'react';
import type { RefObject } from 'react';
import { MAX_ZOOM, MIN_ZOOM } from './commands/evaluate';

const MAX_WHEEL_DELTA = 50;
const WHEEL_ZOOM_SENSITIVITY = 0.005;
const ANCHOR_IDLE_MS = 300;

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
  scale: number | null,
  setScale: (scale: number) => void,
  presentation: unknown
) {
  const scaleRef = useRef(scale);
  const committedScaleRef = useRef(scale);
  const presentationRef = useRef(presentation);
  const setScaleRef = useRef(setScale);
  const cancelPendingRef = useRef<(() => void) | null>(null);
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
    timeStamp: number;
  } | null>(null);
  setScaleRef.current = setScale;

  useLayoutEffect(() => {
    if (Object.is(presentationRef.current, presentation) && scale === committedScaleRef.current) return;
    presentationRef.current = presentation;
    committedScaleRef.current = scale;
    scaleRef.current = scale;
    cancelPendingRef.current?.();
  }, [presentation, scale]);

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
      clearAnchor();
      viewport.scrollLeft += after.left + pending.anchorX * after.width - pending.clientX;
      viewport.scrollTop += after.top + pending.anchorY * after.height - pending.clientY;
    };
    const observer =
      typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(applyAnchor);
    const clearAnchor = () => {
      pendingAnchorRef.current = null;
      observer?.disconnect();
    };
    const commitScale = () => {
      const currentScale = scaleRef.current;
      if (currentScale === null) return;
      committedScaleRef.current = currentScale;
      setScaleRef.current(currentScale);
    };
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey || event.metaKey || event.altKey || !event.cancelable) return;
      const content = contentRef.current;
      const currentScale = scaleRef.current;
      if (!content || currentScale === null) return;
      event.preventDefault();
      if (pendingAnchorRef.current && event.timeStamp - pendingAnchorRef.current.timeStamp > ANCHOR_IDLE_MS) {
        clearAnchor();
      }
      const next = pinchZoomScale(
        currentScale,
        event.deltaY,
        event.deltaMode,
        viewport.clientHeight
      );
      if (next === currentScale) return;
      const before = content.getBoundingClientRect();
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
        timeStamp: event.timeStamp,
      };
      observer?.disconnect();
      observer?.observe(content);
      scaleRef.current = next;
      if (!view) {
        commitScale();
        return;
      }
      if (commitAnimationRef.current === null) {
        commitAnimationRef.current = view.requestAnimationFrame(() => {
          commitAnimationRef.current = null;
          commitScale();
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
    const cancelPending = () => {
      clearAnchor();
      if (view && commitAnimationRef.current !== null) {
        view.cancelAnimationFrame(commitAnimationRef.current);
        commitAnimationRef.current = null;
      }
      if (view && anchorAnimationRef.current !== null) {
        view.cancelAnimationFrame(anchorAnimationRef.current);
        anchorAnimationRef.current = null;
      }
    };
    cancelPendingRef.current = cancelPending;
    viewport.addEventListener('wheel', onWheel, { passive: false });
    return () => {
      viewport.removeEventListener('wheel', onWheel);
      cancelPending();
      cancelPendingRef.current = null;
    };
  }, [contentRef, presentation, viewportRef]);
}
