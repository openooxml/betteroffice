import { useCallback, useEffect, useRef, type PointerEvent, type RefObject } from 'react';
import {
  commandMotion, navigateConfirmedSelection, setPeerlessGestureActive, peerlessInputReference,
  peerlessKnownState, type PeerlessViewStore,
} from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { peerlessPositionAtPoint } from './createPeerlessEditorApi';

let nextPointerGroup = 0;

export function usePeerlessTextSelection(store: PeerlessViewStore | null, {
  canvas, presentation, focus,
}: { canvas?: RefObject<HTMLCanvasElement | null>; presentation?: PeerlessPresentation | null;
  focus?: () => void } = {}) {
  const gesture = useRef<{ pointerId: number; group: string; element: HTMLElement;
    granularity: 'character' | 'word' | 'paragraph'; x: number; y: number;
    lastX: number; lastY: number; dragging: boolean; inFlight?: boolean;
    queued?: PointerEvent<HTMLElement> } | null>(null);
  const recent = useRef<{ time: number; x: number; y: number; target: string; count: number } | null>(null);
  useEffect(() => () => {
    const current = gesture.current;
    if (store) setPeerlessGestureActive(store, 'text', false);
    gesture.current = null;
    if (current) current.element.releasePointerCapture?.(current.pointerId);
    recent.current = null;
  }, [store]);

  const move = useCallback((event: KeyboardEvent): boolean => {
    if (!store || store.getSnapshot().selection?.kind === 'object') return false;
    const motion = commandMotion(event.key, event);
    if (!motion) return false;
    const textDependent = motion === 'wordLeft' || motion === 'wordRight' ||
      motion === 'storyStart' || motion === 'storyEnd';
    if (!textDependent && peerlessKnownState(store)) {
      const selection = navigateConfirmedSelection(store, event.key, event.shiftKey, event);
      if (!selection) return false;
      event.preventDefault();
      void store.dispatch({ kind: 'setSelection', payload: { selection,
        inputTime: event.timeStamp, undoGroup: 'selection' } }).catch(() => {});
      return true;
    }
    const at = peerlessInputReference(store, true);
    if (!at) return false;
    event.preventDefault();
    void store.dispatch({ kind: 'moveSelection', payload: { at, motion, extend: event.shiftKey,
      inputTime: event.timeStamp, undoGroup: 'selection' } }).catch(() => {});
    return true;
  }, [store]);

  const select = useCallback(function select(event: PointerEvent<HTMLElement>, drag: boolean, final = false): boolean {
    if (!store || !presentation?.selectionEnabled || !canvas?.current || store.getSnapshot().composition) return false;
    const confirmed = presentation.getVisualSnapshot()?.view.confirmed;
    const bounds = canvas.current.getBoundingClientRect();
    if (!confirmed || !bounds.width || !bounds.height) return false;
    const x = (event.clientX - bounds.left) * confirmed.displayList.width / bounds.width;
    const y = (event.clientY - bounds.top) * confirmed.displayList.height / bounds.height;
    if (!drag && (x < 0 || y < 0 || x > confirmed.displayList.width || y > confirmed.displayList.height)) return false;
    const at = peerlessInputReference(store, true);
    let extend = drag;
    if (!drag) {
      const hit = peerlessPositionAtPoint(presentation, canvas.current, event.clientX, event.clientY);
      if (!hit) {
        recent.current = null;
        event.preventDefault();
        void store.dispatch({ kind: 'setSelection', payload: { selection: null,
          inputTime: event.timeStamp, undoGroup: 'selection' } }).catch(() => {});
        event.currentTarget.focus();
        return true;
      }
      if (hit.kind !== 'text') return false;
      const selected = at?.source === 'confirmed' ? at.selection : null;
      extend = event.shiftKey && (at?.source === 'afterSeq' || selected?.slideId === confirmed.frame.slideId &&
        selected.shapeId === hit.shapeId && selected.storyId === hit.storyId && selected.kind !== 'shape' && selected.kind !== 'object');
      const target = JSON.stringify([confirmed.frame.slideId, hit.shapeId, hit.storyId]);
      const previous = recent.current;
      const repeated = previous?.target === target && event.timeStamp - previous.time <= 600 &&
        Math.hypot(event.clientX - previous.x, event.clientY - previous.y) <= 6;
      const count = event.shiftKey ? 1 : event.detail > 0 ? Math.min(event.detail, 3)
        : repeated ? Math.min(previous!.count + 1, 3) : 1;
      recent.current = { target, count, time: event.timeStamp, x: event.clientX, y: event.clientY };
      gesture.current = { pointerId: event.pointerId, group: `pointer-${++nextPointerGroup}`,
        granularity: count >= 3 ? 'paragraph' : count === 2 ? 'word' : 'character',
        element: event.currentTarget, x: event.clientX, y: event.clientY,
        lastX: event.clientX, lastY: event.clientY, dragging: false };
      setPeerlessGestureActive(store, 'text', true);
      event.currentTarget.setPointerCapture?.(event.pointerId);
    }
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return false;
    event.preventDefault();
    if (drag && current.lastX === event.clientX && current.lastY === event.clientY) {
      current.queued = undefined;
      return true;
    }
    if (drag && current.inFlight && !final) {
      current.queued = event;
      return true;
    }
    current.queued = undefined;
    current.lastX = event.clientX;
    current.lastY = event.clientY;
    const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = confirmed.frame;
    if (drag) current.inFlight = true;
    const operation = store.dispatch({ kind: 'selectAtPoint', payload: {
      ...(at ? { at } : {}), frame: { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation },
      x, y, extend, granularity: current.granularity, drag,
      inputTime: event.timeStamp, undoGroup: current.group,
    } });
    const settled = () => {
      if (!drag || gesture.current !== current) return;
      current.inFlight = false;
      const queued = current.queued;
      current.queued = undefined;
      if (queued) select(queued, true);
    };
    void operation.then(settled, settled);
    focus?.();
    return true;
  }, [store, presentation, canvas, focus]);

  const flush = useCallback(() => {
    const queued = gesture.current?.queued;
    if (queued) select(queued, true, true);
  }, [select]);
  useEffect(() => {
    const stage = canvas?.current?.closest('[role="application"]');
    if (!stage) return;
    const events = ['keydown', 'beforeinput', 'paste', 'compositionstart'];
    for (const event of events) stage.addEventListener(event, flush, true);
    return () => { for (const event of events) stage.removeEventListener(event, flush, true); };
  }, [canvas, store, flush]);

  const onPointerDown = useCallback((event: PointerEvent<HTMLElement>) => {
    if (event.button !== 0 || event.isPrimary === false) return;
    flush();
    select(event, false);
  }, [select, flush]);
  const onPointerMove = useCallback((event: PointerEvent<HTMLElement>) => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return;
    if (!current.dragging && Math.hypot(event.clientX - current.x, event.clientY - current.y) <= 3) return;
    current.dragging = true;
    recent.current = null;
    select(event, true);
  }, [select]);
  const onPointerUp = useCallback((event: PointerEvent<HTMLElement>) => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return;
    if (current.dragging) select(event, true, true);
    gesture.current = null;
    if (store) setPeerlessGestureActive(store, 'text', false);
    current.element.releasePointerCapture?.(event.pointerId);
  }, [select, store]);
  const onPointerCancel = useCallback((event: PointerEvent<HTMLElement>) => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return;
    gesture.current = null;
    if (store) setPeerlessGestureActive(store, 'text', false);
    current.element.releasePointerCapture?.(event.pointerId);
    recent.current = null;
  }, [store]);
  return { move, onPointerDown, onPointerMove, onPointerUp, onPointerCancel,
    onLostPointerCapture: onPointerCancel };
}
