import { useCallback, useEffect, useRef, type KeyboardEvent, type PointerEvent, type RefObject } from 'react';
import type { ShapeRect } from '@betteroffice/pptx';
import type {
  PresentationCommandFrameTag, PresentationCommandObjectSelection, PresentationCommandObjectTarget,
} from '@betteroffice/pptx/worker';
import { MIN_SHAPE_SIZE_EMU, passedDragThreshold, RESIZE_HANDLES, resizeCommitDelta, resizeCursor, resizedBounds,
  slidePoint, type ResizeHandle, type SlidePoint } from '../interactions';
import { hitConfirmedGeometry, peerlessGestureActive, peerlessInputReference, setPeerlessGestureActive, type PeerlessViewStore } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

interface ShapeGesture {
  pointerId: number;
  element: HTMLElement;
  group: string;
  frame: PresentationCommandFrameTag;
  start: SlidePoint;
  client: SlidePoint;
  delta: SlidePoint;
  shapeId: string;
  dragThreshold: number;
  handle?: ResizeHandle;
  object: Promise<PresentationCommandObjectSelection | null>;
  selected: PresentationCommandObjectSelection | null;
  dragging: boolean;
  cancelled: boolean;
  abort: AbortController;
}

export interface PeerlessShapePreview {
  token: string;
  bounds: ShapeRect;
  retainCanvas?: { slideId: string; release(): void };
  request?: {
    gesture: string;
    frame: PresentationCommandFrameTag;
    target: PresentationCommandObjectTarget;
    rect: ShapeRect;
  };
}

export interface PeerlessShapePreviewStore {
  getSnapshot(): PeerlessShapePreview | null;
  subscribe(listener: () => void): () => void;
}

let nextShapeGroup = 0;

export function peerlessGestureRect(object: PresentationCommandObjectSelection, delta: SlidePoint,
  handle?: ResizeHandle): ShapeRect | null {
  if (!object.rect || !object.canMove || handle && !object.canResize) return null;
  const emu = { x: Math.round(delta.x * object.units.x), y: Math.round(delta.y * object.units.y) };
  if (emu.x === 0 && emu.y === 0) return { ...object.rect };
  return handle ? resizedBounds(object.rect, handle, emu, MIN_SHAPE_SIZE_EMU, object.aspectRatioLocked)
    : { ...object.rect, x: object.rect.x + emu.x, y: object.rect.y + emu.y };
}

function previewBounds(object: PresentationCommandObjectSelection, rect: ShapeRect, resize: boolean): ShapeRect {
  return resize ? { x: rect.x / object.units.x, y: rect.y / object.units.y,
    width: rect.width / object.units.x, height: rect.height / object.units.y }
    : { ...object.bounds, x: object.bounds.x + (rect.x - object.rect!.x) / object.units.x,
      y: object.bounds.y + (rect.y - object.rect!.y) / object.units.y };
}

export function usePeerlessShapeGestures(store: PeerlessViewStore | null, {
  canvas, presentation,
}: { canvas: RefObject<HTMLCanvasElement | null>; presentation: PeerlessPresentation | null }) {
  const gesture = useRef<ShapeGesture | null>(null);
  const outstanding = useRef(new Set<ShapeGesture>());
  const recent = useRef<{ slideId: string; shapeId: string; x: number; y: number; time: number } | null>(null);
  const previewValue = useRef<PeerlessShapePreview | null>(null);
  const previewListeners = useRef(new Set<() => void>());
  const preview = useRef<PeerlessShapePreviewStore>({
    getSnapshot: () => previewValue.current,
    subscribe: (listener) => {
      previewListeners.current.add(listener);
      return () => { previewListeners.current.delete(listener); };
    },
  });
  const setPreview = useCallback((value: PeerlessShapePreview | null) => {
    if (previewValue.current === value) return;
    previewValue.current = value;
    for (const listener of previewListeners.current) listener();
  }, []);

  const cancelActive = useCallback(() => {
    const current = gesture.current;
    if (store) setPeerlessGestureActive(store, 'shape', false);
    gesture.current = null;
    recent.current = null;
    if (current) {
      current.cancelled = true;
      current.abort.abort();
      outstanding.current.delete(current);
      current.element.releasePointerCapture?.(current.pointerId);
    }
    setPreview(null);
  }, [store, setPreview]);

  const cancel = useCallback(() => {
    cancelActive();
    for (const pending of outstanding.current) {
      pending.cancelled = true;
      pending.abort.abort();
    }
    outstanding.current.clear();
  }, [cancelActive]);

  useEffect(() => {
    setPreview(null);
    if (!store) return;
    const changed = () => {
      const current = gesture.current;
      const preview = previewValue.current;
      const slideId = current?.frame.slideId ?? preview?.retainCanvas?.slideId;
      if (slideId === undefined) return;
      const token = current ? current.selected?.token : preview?.token;
      const view = store.getSnapshot();
      const frame = view.confirmed?.frame;
      if (view.editState.readOnly || !frame || frame.slideId !== slideId ||
        current && frame.slideRevision !== current.frame.slideRevision ||
        token !== undefined && (view.selection?.kind !== 'object' || view.selection.token !== token)) cancel();
    };
    const unsubscribe = [store.subscribePresentation(changed), store.subscribeEditState(changed)];
    if (presentation) unsubscribe.push(presentation.subscribeVisual(() => {
      if (previewValue.current?.retainCanvas && !presentation.pluginVisual()) setPreview(null);
    }));
    window.addEventListener('blur', cancel);
    return () => {
      window.removeEventListener('blur', cancel);
      for (const release of unsubscribe) release();
      const current = gesture.current;
      setPeerlessGestureActive(store, 'shape', false);
      gesture.current = null;
      for (const pending of outstanding.current) {
        pending.cancelled = true;
        pending.abort.abort();
      }
      outstanding.current.clear();
      recent.current = null;
      current?.element.releasePointerCapture?.(current.pointerId);
    };
  }, [store, presentation, cancel, setPreview]);

  const position = useCallback((event: PointerEvent<HTMLElement>): SlidePoint | null => {
    const confirmed = presentation?.getVisualSnapshot()?.view.confirmed;
    return canvas.current && confirmed
      ? slidePoint(canvas.current.getBoundingClientRect(), confirmed.displayList, event.clientX, event.clientY) : null;
  }, [canvas, presentation]);

  const paint = useCallback((current: ShapeGesture) => {
    if (gesture.current !== current || current.cancelled || !current.dragging || !current.selected ||
      store?.getSnapshot().editState.readOnly) return;
    const rect = peerlessGestureRect(current.selected, current.delta, current.handle);
    if (rect) setPreview({ token: current.selected.token, bounds: previewBounds(current.selected, rect, Boolean(current.handle)),
      ...(current.handle && current.selected.storyId ? { request: { gesture: current.group, frame: current.frame,
        target: { slideId: current.selected.slideId, shapeId: current.selected.shapeId, token: current.selected.token },
        rect } } : {}) });
  }, [store, setPreview]);

  const onPointerDown = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    if (!store || !presentation?.selectionEnabled || !canvas.current || event.button !== 0 || event.isPrimary === false) return false;
    const handle = (event.target as HTMLElement).dataset?.pptxResize as ResizeHandle | undefined;
    const view = store.getSnapshot();
    if (view.composition) {
      if (handle) event.preventDefault();
      return Boolean(handle);
    }
    const visual = presentation.getVisualSnapshot();
    const confirmed = visual?.view.confirmed;
    const start = position(event);
    if (!confirmed || !start) return false;
    if (handle && !RESIZE_HANDLES.includes(handle)) return false;
    const selected = view.selection;
    const scale = canvas.current.getBoundingClientRect().width / confirmed.displayList.width;
    if (handle && (!selected || selected.kind === 'object' && !selected.canResize ||
      view.editState.readOnly)) return true;
    const hit = handle ? null : hitConfirmedGeometry(store, start, scale);
    if (!handle && hit?.kind !== 'shape') return false;
    const shapeId = handle ? selected!.shapeId : hit!.shapeId;
    const previous = recent.current;
    const repeated = !handle && previous?.slideId === confirmed.frame.slideId && previous.shapeId === shapeId &&
      event.timeStamp >= previous.time && event.timeStamp - previous.time <= 600 &&
      Math.hypot(event.clientX - previous.x, event.clientY - previous.y) <= 6;
    cancelActive();
    const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = confirmed.frame;
    const frame = { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation };
    const group = `shape-${++nextShapeGroup}`;
    const operation = handle && selected?.kind === 'object' ? Promise.resolve(selected)
      : store.dispatch({ kind: 'selectObject', payload: handle
        ? { frame, shapeId, inputTime: event.timeStamp, undoGroup: group }
        : { frame, x: start.x, y: start.y, scale, inputTime: event.timeStamp, undoGroup: group } }).then((ack) =>
        ack.outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled') ||
        ack.selection?.kind !== 'object' ? null : ack.selection, () => null);
    const current: ShapeGesture = { pointerId: event.pointerId, element: event.currentTarget, frame, group,
      start, client: { x: event.clientX, y: event.clientY }, delta: { x: 0, y: 0 },
      shapeId, dragThreshold: repeated ? 8 : 4,
      handle, object: operation, selected: handle && selected?.kind === 'object' ? selected : null,
      dragging: false, cancelled: false, abort: new AbortController() };
    gesture.current = current;
    setPeerlessGestureActive(store, 'shape', true);
    outstanding.current.add(current);
    void operation.then((object) => {
      current.selected = object;
      paint(current);
    });
    event.preventDefault();
    event.currentTarget.focus();
    event.currentTarget.setPointerCapture?.(event.pointerId);
    return true;
  }, [store, presentation, canvas, cancelActive, position, paint]);

  const cursorAt = useCallback((event: PointerEvent<HTMLElement>): string => {
    const handle = (event.target as HTMLElement).dataset?.pptxResize as ResizeHandle | undefined;
    if (handle && RESIZE_HANDLES.includes(handle)) return resizeCursor(handle);
    const confirmed = presentation?.getVisualSnapshot()?.view.confirmed;
    const point = position(event);
    if (!store || !presentation?.selectionEnabled || !canvas.current || !confirmed || !point) return 'default';
    const hit = hitConfirmedGeometry(store, point, canvas.current.getBoundingClientRect().width / confirmed.displayList.width);
    return hit?.kind === 'text' ? 'text' : hit?.kind === 'shape' ? 'move' : 'default';
  }, [store, presentation, canvas, position]);

  const hoverCursor = useCallback((event: PointerEvent<HTMLElement>): string | null => {
    if (event.buttons !== 0 || (store && peerlessGestureActive(store))) return null;
    return cursorAt(event);
  }, [store, cursorAt]);

  const onPointerMove = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return false;
    const point = position(event);
    if (!point) return true;
    current.dragging ||= passedDragThreshold(current.client.x, current.client.y, event.clientX, event.clientY, current.dragThreshold);
    current.delta = { x: point.x - current.start.x, y: point.y - current.start.y };
    paint(current);
    return true;
  }, [position, paint]);

  const onPointerUp = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    const current = gesture.current;
    if (!store || !current || current.pointerId !== event.pointerId) return false;
    const point = position(event);
    const delta = resizeCommitDelta(current.start, current.delta, point);
    const dragging = current.dragging || passedDragThreshold(current.client.x, current.client.y,
      event.clientX, event.clientY, current.dragThreshold);
    const inputTime = event.timeStamp;
    gesture.current = null;
    setPeerlessGestureActive(store, 'shape', false);
    recent.current = !current.handle && !dragging ? { slideId: current.frame.slideId, shapeId: current.shapeId,
      x: event.clientX, y: event.clientY, time: inputTime } : null;
    current.element.releasePointerCapture?.(current.pointerId);
    if (!dragging || delta.x === 0 && delta.y === 0) {
      outstanding.current.delete(current);
      setPreview(null);
      return true;
    }
    const commit = (object: PresentationCommandObjectSelection) => {
      outstanding.current.delete(current);
      if (current.cancelled || !presentation?.current || store.getSnapshot().editState.readOnly) {
        setPreview(null);
        return;
      }
      const rect = peerlessGestureRect(object, delta, current.handle);
      if (!rect || Object.keys(rect).every((key) =>
        rect[key as keyof ShapeRect] === object.rect![key as keyof ShapeRect])) {
        setPreview(null);
        return;
      }
      const payload = { frame: current.frame,
        target: { slideId: object.slideId, shapeId: object.shapeId, token: object.token },
        inputTime, undoGroup: current.group };
      const preview = previewValue.current;
      const retained = preview?.request ? { token: preview.token, bounds: preview.bounds,
        retainCanvas: { slideId: object.slideId, release: () => setPreview(null) } } : null;
      setPreview(retained);
      const operation = store.dispatch(current.handle ? { kind: 'setShapeRect', payload: { ...payload, rect } }
        : { kind: 'moveShape', payload: { ...payload, x: rect.x, y: rect.y } });
      void operation.then(async (ack) => {
        const outcome = ack.outcomes.find((entry) => entry.seq === ack.throughSeq);
        if (!outcome?.committed || !retained || previewValue.current !== retained) return;
        await presentation.whenPaintedThrough(outcome.seq);
      }).catch(() => {}).finally(() => {
        if (previewValue.current === retained) setPreview(null);
      });
    };
    if (!current.selected) {
      if (current.handle) {
        if (current.cancelled || !presentation?.current || store.getSnapshot().editState.readOnly) {
          outstanding.current.delete(current);
          setPreview(null);
          return true;
        }
        const at = peerlessInputReference(store, true);
        if (at?.source === 'afterSeq') {
          const preview = previewValue.current;
          const retained = preview?.request ? { token: preview.token, bounds: preview.bounds,
            retainCanvas: { slideId: current.frame.slideId, release: () => setPreview(null) } } : null;
          setPreview(retained);
          const resize = store.dispatch({ kind: 'setShapeRect', payload: { at, frame: current.frame,
            shapeId: current.shapeId, handle: current.handle, delta,
            inputTime, undoGroup: current.group } }, {
            reportRefusal: false, signal: current.abort.signal, waitForPredecessor: true,
          });
          void resize.then(() => { outstanding.current.delete(current); },
            () => { outstanding.current.delete(current); });
          void resize.then(async (ack) => {
            const outcome = ack.outcomes.find((entry) => entry.seq === ack.throughSeq);
            if (current.abort.signal.aborted || !outcome?.committed || !retained || previewValue.current !== retained) return;
            await presentation.whenPaintedThrough(outcome.seq);
          }).catch(() => {}).finally(() => {
            if (previewValue.current === retained) setPreview(null);
          });
        } else {
          outstanding.current.delete(current);
          setPreview(null);
        }
        return true;
      }
      outstanding.current.delete(current);
      setPreview(null);
      if (current.cancelled || !presentation?.current || store.getSnapshot().editState.readOnly) return true;
      const at = peerlessInputReference(store, true);
      if (at?.source === 'afterSeq') {
        void store.dispatch({ kind: 'moveShape', payload: { at, frame: current.frame,
          shapeId: current.shapeId, delta, inputTime, undoGroup: current.group } }).catch(() => {});
      }
      return true;
    }
    commit(current.selected);
    return true;
  }, [store, presentation, position, setPreview]);

  const onPointerCancel = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    if (gesture.current?.pointerId !== event.pointerId) return false;
    cancelActive();
    return true;
  }, [cancelActive]);

  const onKeyDown = useCallback((event: KeyboardEvent<HTMLElement>) => {
    if (event.key !== 'Escape' || !outstanding.current.size) return;
    event.preventDefault();
    event.stopPropagation();
    cancel();
  }, [cancel]);

  return { preview: preview.current, cursorAt, hoverCursor, onPointerDown, onPointerMove, onPointerUp, onPointerCancel, onKeyDown,
    onLostPointerCapture: onPointerCancel };
}
