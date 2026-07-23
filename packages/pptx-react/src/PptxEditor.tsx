import {
  initWasm,
  openPresentation,
  paintSlide,
  PRESENCE_LABEL_DURATION_MS,
  sizeCanvasForSlide,
} from '@betteroffice/pptx';
import type {
  CanvasImageResolver,
  CollaborationReplica,
  DeckSnapshot,
  PptxPresence,
  PptxPresencePeer,
  PptxFontFace,
  PresentationHandle,
  SlideDisplayList,
  StorySnapshot,
  TextBoxPrimitive,
  TextStyle,
  TextStylePatch,
} from '@betteroffice/pptx';
import type { Translations } from '@betteroffice/pptx-i18n';
import { LocaleProvider, useTranslation } from './i18n';
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type {
  CSSProperties,
  ChangeEvent,
  KeyboardEvent,
  MouseEvent as ReactMouseEvent,
  PointerEvent,
} from 'react';
import {
  canMoveShape,
  findShape,
  findTopLevelShape,
  frameBoundsForShape,
  indexShapes,
  movedShapePosition,
  passedDragThreshold,
  slidePoint,
  textPositionAtPoint,
} from './interactions';
import type { FrameBounds, SlidePoint } from './interactions';
import {
  groupPresenceBySlide,
  groupShapePresence,
  limitPresence,
  type BoundedPresence,
} from './presence-rendering';

export interface PptxTextSelection {
  shapeId: string;
  storyId: string;
  anchor: number;
  focus: number;
}

export interface PptxEditorApi {
  handle: PresentationHandle;
  refresh: () => void;
}

export interface PptxEditorCollaborationOptions {
  clientId: number;
  initialUpdate?: Uint8Array;
  onReplica?: (replica: CollaborationReplica | null) => void;
  presence?: PptxPresence;
}

export interface PptxEditorProps {
  file?: Uint8Array;
  /** Font faces; equivalent inline arrays do not reopen the presentation. */
  fonts: ReadonlyArray<PptxFontFace>;
  clientId?: number;
  collaboration?: PptxEditorCollaborationOptions;
  i18n?: Translations;
  className?: string;
  onReady?: (api: PptxEditorApi) => void;
  onChange?: (snapshot: DeckSnapshot) => void;
  onError?: (error: Error) => void;
}

interface EditorModel {
  snapshot: DeckSnapshot;
  slideIndex: number;
  frame: SlideDisplayList | null;
  thumbnails: Map<string, SlideDisplayList>;
}

interface PptxShapeSelection {
  slideId: string;
  shapeId: string;
}

type PointerGesture =
  | {
      kind: 'text';
      pointerId: number;
      slideId: string;
      shapeId: string;
      storyId: string;
      anchor: number;
    }
  | {
      kind: 'shape';
      pointerId: number;
      slideId: string;
      shapeId: string;
      startClientX: number;
      startClientY: number;
      start: SlidePoint;
      last: SlidePoint;
      dragThreshold: number;
      repeatedClick: boolean;
      dragging: boolean;
    };

interface ShapeDragPreview {
  shapeId: string;
  delta: SlidePoint;
}

interface RemoteShapePresence {
  peer: PptxPresencePeer;
  peerCount: number;
  shapeId: string;
  bounds: FrameBounds;
}

interface RecentShapeClick {
  slideId: string;
  shapeId: string;
  clientX: number;
  clientY: number;
  timeStamp: number;
}

const initialStyle: Required<Pick<TextStyle, 'bold' | 'italic' | 'fontSizePt' | 'color'>> = {
  bold: false,
  italic: false,
  fontSizePt: 24,
  color: '#111827',
};

type PptxToolbarIconName = 'undo' | 'redo' | 'addSlide' | 'deleteSlide' | 'textBox';

function PptxToolbarIcon({ name }: { name: PptxToolbarIconName }) {
  return (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
      style={{ flex: '0 0 auto' }}
    >
      {name === 'undo' && <path d="m9 7-5 5 5 5M5 12h9a6 6 0 0 1 6 6" />}
      {name === 'redo' && <path d="m15 7 5 5-5 5m4-5h-9a6 6 0 0 0-6 6" />}
      {name === 'addSlide' && (
        <>
          <rect x="3" y="5" width="14" height="14" rx="2" />
          <path d="M7 9h6M7 13h4M20 10v6m-3-3h6" />
        </>
      )}
      {name === 'deleteSlide' && (
        <>
          <path d="M5 7h14M9 7V4h6v3m2 0-1 13H8L7 7" />
          <path d="M10 11v5m4-5v5" />
        </>
      )}
      {name === 'textBox' && (
        <>
          <rect x="3" y="4" width="18" height="16" rx="2" />
          <path d="M8 8h8m-4 0v8m-3 0h6" />
        </>
      )}
    </svg>
  );
}

export function PptxEditor({
  i18n,
  ...props
}: PptxEditorProps) {
  return (
    <LocaleProvider i18n={i18n}>
      <PptxEditorContent {...props} />
    </LocaleProvider>
  );
}

function PptxEditorContent({
  file,
  fonts,
  clientId,
  collaboration,
  className,
  onReady,
  onChange,
  onError,
}: Omit<PptxEditorProps, 'i18n'>) {
  const { t } = useTranslation();
  const decodeImageError = t('errors.decodeSlideImage');
  const collaborationClientId = collaboration?.clientId ?? clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const collaborationOnReplica = collaboration?.onReplica;
  const collaborationPresence = collaboration?.presence;
  const handleRef = useRef<PresentationHandle | null>(null);
  const modelRef = useRef<EditorModel | null>(null);
  const onReadyRef = useRef(onReady);
  const onChangeRef = useRef(onChange);
  const onErrorRef = useRef(onError);
  const stageRef = useRef<HTMLDivElement>(null);
  const canvasHostRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement>(null);
  const pointerGestureRef = useRef<PointerGesture | null>(null);
  const recentShapeClickRef = useRef<RecentShapeClick | null>(null);
  const imageCacheRef = useRef(new Map<string, Promise<CanvasImageSource | null>>());
  const stableFonts = useStableFontFaces(fonts);
  const [model, setModel] = useState<EditorModel | null>(null);
  const [selection, setSelection] = useState<PptxTextSelection | null>(null);
  const [shapeSelection, setShapeSelection] = useState<PptxShapeSelection | null>(null);
  const [dragPreview, setDragPreview] = useState<ShapeDragPreview | null>(null);
  const [textStyle, setTextStyle] = useState(initialStyle);
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);
  const [remotePeers, setRemotePeers] = useState<readonly PptxPresencePeer[]>([]);

  onReadyRef.current = onReady;
  onChangeRef.current = onChange;
  onErrorRef.current = onError;
  modelRef.current = model;

  const reportError = useCallback((value: unknown) => {
    const next = value instanceof Error ? value : new Error(String(value));
    setError(next.message);
    onErrorRef.current?.(next);
  }, []);

  const refreshAt = useCallback(
    (requestedIndex?: number, notify = false, refreshAll = false): EditorModel | null => {
      const handle = handleRef.current;
      if (!handle) return null;
      try {
        const snapshot = handle.snapshot();
        const index = clampSlideIndex(
          requestedIndex ?? modelRef.current?.slideIndex ?? 0,
          snapshot.slides.length
        );
        const thumbnails = new Map<string, SlideDisplayList>();
        for (let slideIndex = 0; slideIndex < snapshot.slides.length; slideIndex += 1) {
          const slide = snapshot.slides[slideIndex];
          const cached = modelRef.current?.thumbnails.get(slide.id);
          if (slideIndex !== index && cached && !refreshAll) thumbnails.set(slide.id, cached);
          else if (slideIndex !== index) thumbnails.set(slide.id, handle.layoutSlide(slideIndex));
        }
        const frame = snapshot.slides.length > 0 ? handle.layoutSlide(index) : null;
        if (frame) thumbnails.set(snapshot.slides[index].id, frame);
        const next = { snapshot, slideIndex: index, frame, thumbnails };
        const activeSlide = snapshot.slides[index];
        setSelection((current) =>
          current && activeSlide && findShape(activeSlide.shapes, current.shapeId) ? current : null
        );
        setShapeSelection((current) =>
          current &&
          activeSlide?.id === current.slideId &&
          findShape(activeSlide.shapes, current.shapeId)
            ? current
            : null
        );
        const gesture = pointerGestureRef.current;
        if (
          gesture &&
          (!activeSlide ||
            activeSlide.id !== gesture.slideId ||
            !findShape(activeSlide.shapes, gesture.shapeId))
        ) {
          pointerGestureRef.current = null;
          recentShapeClickRef.current = null;
          setDragPreview(null);
        }
        modelRef.current = next;
        setModel(next);
        setError(null);
        if (notify) onChangeRef.current?.(snapshot);
        return next;
      } catch (value) {
        reportError(value);
        return null;
      }
    },
    [reportError]
  );

  const refresh = useCallback(() => {
    refreshAt(undefined, false, true);
  }, [refreshAt]);

  useEffect(() => {
    let disposed = false;
    let handle: PresentationHandle | null = null;
    let browserFaces: FontFace[] = [];
    let unsubscribeUpdates = () => {};
    handleRef.current?.dispose();
    handleRef.current = null;
    setCollaborationReplica(null);
    modelRef.current = null;
    setModel(null);
    setSelection(null);
    setShapeSelection(null);
    setDragPreview(null);
    pointerGestureRef.current = null;
    recentShapeClickRef.current = null;
    setError(null);
    imageCacheRef.current.clear();
    if (!file) return;
    setLoading(true);
    void Promise.all([initWasm(), installBrowserFonts(stableFonts)]).then(
      ([, installed]) => {
        browserFaces = installed;
        if (disposed) {
          removeBrowserFonts(browserFaces);
          return;
        }
        try {
          handle = openPresentation(file, {
            clientId: collaborationClientId,
            fonts: stableFonts,
            initialUpdate: collaborationInitialUpdate,
          });
          handleRef.current = handle;
          unsubscribeUpdates = handle.onUpdate((_update, origin) => {
            if (origin === 'remote') refreshAt(undefined, true, true);
          });
          refreshAt(0);
          setLoading(false);
          setCollaborationReplica(handle);
          onReadyRef.current?.({ handle, refresh });
        } catch (value) {
          setLoading(false);
          reportError(value);
        }
      },
      (value: unknown) => {
        if (disposed) return;
        setLoading(false);
        reportError(value);
      }
    );
    return () => {
      disposed = true;
      unsubscribeUpdates();
      handle?.dispose();
      if (handleRef.current === handle) handleRef.current = null;
      removeBrowserFonts(browserFaces);
    };
  }, [
    collaborationClientId,
    collaborationInitialUpdate,
    file,
    stableFonts,
    refresh,
    refreshAt,
    reportError,
  ]);

  useEffect(() => {
    if (!collaborationOnReplica || !collaborationReplica) return;
    collaborationOnReplica(collaborationReplica);
    return () => collaborationOnReplica(null);
  }, [collaborationOnReplica, collaborationReplica]);

  useEffect(() => {
    if (!collaborationPresence) {
      setRemotePeers([]);
      return;
    }
    setRemotePeers(collaborationPresence.peers);
    return collaborationPresence.onPresence(setRemotePeers);
  }, [collaborationPresence]);

  useEffect(() => {
    const host = canvasHostRef.current;
    if (!host) return;
    const update = () => setViewport({ width: host.clientWidth, height: host.clientHeight });
    update();
    const observer = new ResizeObserver(update);
    observer.observe(host);
    return () => observer.disconnect();
  }, []);

  const scale = useMemo(() => {
    if (!model?.frame || viewport.width <= 0 || viewport.height <= 0) return 1;
    return Math.min(
      (viewport.width - 40) / model.frame.width,
      (viewport.height - 40) / model.frame.height,
      1
    );
  }, [model?.frame, viewport]);

  useEffect(() => {
    const canvas = canvasRef.current;
    const frame = model?.frame;
    if (!canvas || !frame) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(canvas, frame, dpr, scale);
    let cancelled = false;
    void paintSlide(ctx, frame, dpr, scale, {
      resolveImage: (assetId) =>
        resolveImage(assetId, handleRef, imageCacheRef, decodeImageError),
    }).catch((value: unknown) => {
      if (!cancelled) reportError(value);
    });
    return () => {
      cancelled = true;
    };
  }, [decodeImageError, model?.frame, reportError, scale]);

  useEffect(() => {
    const canvas = overlayCanvasRef.current;
    const frame = model?.frame;
    if (!canvas || !frame) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(canvas, frame, dpr, scale);
    paintSelection(ctx, frame, selection, dpr, scale);
  }, [model?.frame, scale, selection]);

  const selectedShape = useMemo(() => {
    if (!model?.frame || !shapeSelection) return null;
    const slide = model.snapshot.slides[model.slideIndex];
    if (!slide || slide.id !== shapeSelection.slideId) return null;
    return findShape(slide.shapes, shapeSelection.shapeId);
  }, [model, shapeSelection]);

  const selectedShapeBounds = useMemo<FrameBounds | null>(
    () =>
      model?.frame && selectedShape
        ? frameBoundsForShape(model.snapshot, model.frame, selectedShape)
        : null,
    [model, selectedShape]
  );

  const activeSlide = model?.snapshot.slides[model.slideIndex];
  const currentSlideId = activeSlide?.id;
  const localPresenceShapeId = useMemo(() => {
    if (!activeSlide) return undefined;
    if (shapeSelection?.slideId === activeSlide.id) return shapeSelection.shapeId;
    if (!selection) return undefined;
    return findTopLevelShape(activeSlide, selection.shapeId)?.id;
  }, [activeSlide, selection, shapeSelection]);

  useEffect(() => {
    if (!collaborationPresence) return;
    collaborationPresence.setCursor(
      currentSlideId
        ? {
            slideId: currentSlideId,
            ...(localPresenceShapeId ? { shapeId: localPresenceShapeId } : {}),
          }
        : null
    );
  }, [collaborationPresence, currentSlideId, localPresenceShapeId]);

  useEffect(
    () => () => {
      collaborationPresence?.setCursor(null);
    },
    [collaborationPresence]
  );

  const toolbarPresence = useMemo(() => limitPresence(remotePeers), [remotePeers]);

  const remoteShapePresence = useMemo<BoundedPresence<RemoteShapePresence>>(() => {
    if (!model?.frame || !activeSlide) return { visible: [], overflow: 0 };
    const shapeIndex = indexShapes(activeSlide.shapes);
    const grouped = groupShapePresence(remotePeers, activeSlide.id, shapeIndex);
    const visible: RemoteShapePresence[] = [];
    let overflow = grouped.overflow;
    for (const group of grouped.visible) {
      const shape = shapeIndex.get(group.shapeId);
      if (!shape) continue;
      const bounds = frameBoundsForShape(model.snapshot, model.frame, shape);
      if (bounds) {
        visible.push({
          peer: group.peer,
          peerCount: group.count,
          shapeId: group.shapeId,
          bounds,
        });
      } else {
        overflow += group.count;
      }
    }
    return { visible, overflow };
  }, [activeSlide, model, remotePeers]);

  const remotePeersBySlide = useMemo(() => {
    const slideIds = new Set(model?.snapshot.slides.map((slide) => slide.id));
    return groupPresenceBySlide(remotePeers, slideIds);
  }, [model, remotePeers]);

  const selectSlide = (index: number) => {
    setSelection(null);
    setShapeSelection(null);
    setDragPreview(null);
    pointerGestureRef.current = null;
    recentShapeClickRef.current = null;
    refreshAt(index);
    stageRef.current?.focus();
  };

  const pointerDown = (event: PointerEvent<HTMLCanvasElement>) => {
    if (!event.isPrimary || event.button !== 0) return;
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current?.frame) return;
    const point = slidePoint(
      event.currentTarget.getBoundingClientRect(),
      current.frame,
      event.clientX,
      event.clientY
    );
    if (!point) return;
    try {
      handle.layoutSlide(current.slideIndex);
      const hit = handle.hitTest(point.x, point.y);
      const slide = current.snapshot.slides[current.slideIndex];
      if (
        slide &&
        selection &&
        hit?.kind === 'text' &&
        hit.shapeId === selection.shapeId &&
        hit.storyId === selection.storyId
      ) {
        handle.story(hit.storyId);
        const anchor = event.shiftKey ? selection.anchor : hit.position;
        setSelection({
          shapeId: hit.shapeId,
          storyId: hit.storyId,
          anchor,
          focus: hit.position,
        });
        setShapeSelection(null);
        setDragPreview(null);
        recentShapeClickRef.current = null;
        pointerGestureRef.current = {
          kind: 'text',
          pointerId: event.pointerId,
          slideId: slide.id,
          shapeId: hit.shapeId,
          storyId: hit.storyId,
          anchor,
        };
        event.currentTarget.setPointerCapture(event.pointerId);
      } else if (slide && hit) {
        const shape = findTopLevelShape(slide, hit.shapeId);
        if (shape) {
          setSelection(null);
          setShapeSelection({ slideId: slide.id, shapeId: shape.id });
          setDragPreview(null);
          const recentClick = recentShapeClickRef.current;
          const repeatedClick =
            recentClick?.slideId === slide.id &&
            recentClick.shapeId === shape.id &&
            event.timeStamp - recentClick.timeStamp <= 600 &&
            Math.hypot(
              event.clientX - recentClick.clientX,
              event.clientY - recentClick.clientY
            ) <= 6;
          recentShapeClickRef.current = null;
          if (canMoveShape(shape)) {
            pointerGestureRef.current = {
              kind: 'shape',
              pointerId: event.pointerId,
              slideId: slide.id,
              shapeId: shape.id,
              startClientX: event.clientX,
              startClientY: event.clientY,
              start: point,
              last: point,
              dragThreshold: repeatedClick ? 8 : 4,
              repeatedClick,
              dragging: false,
            };
            event.currentTarget.setPointerCapture(event.pointerId);
          } else {
            pointerGestureRef.current = null;
          }
        } else {
          setSelection(null);
          setShapeSelection(null);
          recentShapeClickRef.current = null;
          pointerGestureRef.current = null;
        }
      } else {
        setSelection(null);
        setShapeSelection(null);
        setDragPreview(null);
        recentShapeClickRef.current = null;
        pointerGestureRef.current = null;
      }
      stageRef.current?.focus();
      event.preventDefault();
    } catch (value) {
      reportError(value);
    }
  };

  const pointerDoubleClick = (event: ReactMouseEvent<HTMLCanvasElement>) => {
    if (event.button !== 0) return;
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current?.frame) return;
    const point = slidePoint(
      event.currentTarget.getBoundingClientRect(),
      current.frame,
      event.clientX,
      event.clientY
    );
    if (!point) return;
    try {
      handle.layoutSlide(current.slideIndex);
      const hit = handle.hitTest(point.x, point.y);
      if (hit?.kind === 'text') {
        handle.story(hit.storyId);
        setSelection({
          shapeId: hit.shapeId,
          storyId: hit.storyId,
          anchor: hit.position,
          focus: hit.position,
        });
        setShapeSelection(null);
        setDragPreview(null);
        pointerGestureRef.current = null;
        recentShapeClickRef.current = null;
        stageRef.current?.focus();
        event.preventDefault();
      }
    } catch (value) {
      reportError(value);
    }
  };

  const updatePointerGesture = (event: PointerEvent<HTMLCanvasElement>): boolean => {
    const gesture = pointerGestureRef.current;
    const current = modelRef.current;
    if (!gesture || gesture.pointerId !== event.pointerId || !current?.frame) {
      return false;
    }
    const point = slidePoint(
      event.currentTarget.getBoundingClientRect(),
      current.frame,
      event.clientX,
      event.clientY
    );
    if (!point || current.snapshot.slides[current.slideIndex]?.id !== gesture.slideId) return false;
    if (gesture.kind === 'text') {
      const focus = textPositionAtPoint(
        current.frame,
        gesture.shapeId,
        gesture.storyId,
        point
      );
      if (focus !== null) {
        setSelection({
          shapeId: gesture.shapeId,
          storyId: gesture.storyId,
          anchor: gesture.anchor,
          focus,
        });
      }
    } else {
      gesture.last = point;
      if (
        !gesture.dragging &&
        passedDragThreshold(
          gesture.startClientX,
          gesture.startClientY,
          event.clientX,
          event.clientY,
          gesture.dragThreshold
        )
      ) {
        gesture.dragging = true;
      }
      if (gesture.dragging) {
        setDragPreview({
          shapeId: gesture.shapeId,
          delta: { x: point.x - gesture.start.x, y: point.y - gesture.start.y },
        });
      }
    }
    return true;
  };

  const pointerMove = (event: PointerEvent<HTMLCanvasElement>) => {
    if (updatePointerGesture(event)) event.preventDefault();
  };

  const pointerUp = (event: PointerEvent<HTMLCanvasElement>) => {
    updatePointerGesture(event);
    const gesture = pointerGestureRef.current;
    if (!gesture || gesture.pointerId !== event.pointerId) return;
    pointerGestureRef.current = null;
    setDragPreview(null);
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
    if (gesture.kind !== 'shape') return;
    if (!gesture.dragging) {
      recentShapeClickRef.current = gesture.repeatedClick
        ? null
        : {
            slideId: gesture.slideId,
            shapeId: gesture.shapeId,
            clientX: event.clientX,
            clientY: event.clientY,
            timeStamp: event.timeStamp,
          };
      return;
    }
    recentShapeClickRef.current = null;
    const handle = handleRef.current;
    const current = modelRef.current;
    const slide = current?.snapshot.slides[current.slideIndex];
    if (!handle || !current?.frame || slide?.id !== gesture.slideId) return;
    const shape = findShape(slide.shapes, gesture.shapeId);
    if (!shape) return;
    try {
      const position = movedShapePosition(current.snapshot, current.frame, shape, {
        x: gesture.last.x - gesture.start.x,
        y: gesture.last.y - gesture.start.y,
      });
      if (position.x !== shape.x || position.y !== shape.y) {
        handle.moveShape(slide.id, shape.id, position.x, position.y);
        refreshAt(undefined, true);
      }
      setShapeSelection({ slideId: slide.id, shapeId: shape.id });
      event.preventDefault();
    } catch (value) {
      reportError(value);
    }
  };

  const cancelPointerGesture = (event: PointerEvent<HTMLCanvasElement>) => {
    if (pointerGestureRef.current?.pointerId !== event.pointerId) return;
    pointerGestureRef.current = null;
    setDragPreview(null);
    recentShapeClickRef.current = null;
  };

  const commit = (nextSelection: PptxTextSelection | null) => {
    setSelection(nextSelection);
    setShapeSelection(null);
    recentShapeClickRef.current = null;
    refreshAt(undefined, true);
  };

  const keyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    if ((event.metaKey || event.ctrlKey) && (event.key === 'b' || event.key === 'B')) {
      event.preventDefault();
      applyFormatting({ bold: !textStyle.bold });
      return;
    }
    if ((event.metaKey || event.ctrlKey) && (event.key === 'i' || event.key === 'I')) {
      event.preventDefault();
      applyFormatting({ italic: !textStyle.italic });
      return;
    }
    const start = Math.min(selection.anchor, selection.focus);
    const end = Math.max(selection.anchor, selection.focus);
    try {
      if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
        event.preventDefault();
        const story = handle.story(selection.storyId);
        const delta = event.key === 'ArrowLeft' ? -1 : 1;
        const focus = Math.max(0, Math.min(story.length, selection.focus + delta));
        setSelection({
          ...selection,
          anchor: event.shiftKey ? selection.anchor : focus,
          focus,
        });
        return;
      }
      if (event.key === 'Backspace') {
        event.preventDefault();
        if (start !== end) {
          handle.deleteText(selection.storyId, start, end);
          commit({ ...selection, anchor: start, focus: start });
          return;
        }
        const previous = previousTextIndex(handle.story(selection.storyId), start);
        if (previous < start) {
          handle.deleteText(selection.storyId, previous, start);
          commit({ ...selection, anchor: previous, focus: previous });
        }
        return;
      }
      if (event.key === 'Delete') {
        event.preventDefault();
        if (start !== end) {
          handle.deleteText(selection.storyId, start, end);
          commit({ ...selection, anchor: start, focus: start });
          return;
        }
        const next = nextTextIndex(handle.story(selection.storyId), end);
        if (next > end) {
          handle.deleteText(selection.storyId, end, next);
          commit({ ...selection, anchor: end, focus: end });
        }
        return;
      }
      if (event.key === 'Enter') {
        event.preventDefault();
        if (start !== end) handle.deleteText(selection.storyId, start, end);
        handle.insertParagraphBreak(selection.storyId, start);
        commit({ ...selection, anchor: start + 1, focus: start + 1 });
        return;
      }
      if (
        event.key.length > 0 &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        Array.from(event.key).length === 1
      ) {
        event.preventDefault();
        if (start !== end) handle.deleteText(selection.storyId, start, end);
        handle.insertText(selection.storyId, start, event.key, textStyle);
        const next = start + event.key.length;
        commit({ ...selection, anchor: next, focus: next });
      }
    } catch (value) {
      reportError(value);
    }
  };

  const applyFormatting = (patch: TextStylePatch) => {
    setTextStyle((current) => ({ ...current, ...patch }));
    const handle = handleRef.current;
    if (!handle || !selection || selection.anchor === selection.focus) return;
    try {
      handle.formatText(
        selection.storyId,
        Math.min(selection.anchor, selection.focus),
        Math.max(selection.anchor, selection.focus),
        patch
      );
      refreshAt(undefined, true);
    } catch (value) {
      reportError(value);
    }
  };

  const addSlide = () => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current) return;
    try {
      const index = current.slideIndex + 1;
      const layout = current.snapshot.slides[current.slideIndex]?.layoutPartPath ?? undefined;
      handle.insertSlide(index, layout);
      setSelection(null);
      setShapeSelection(null);
      setDragPreview(null);
      pointerGestureRef.current = null;
      recentShapeClickRef.current = null;
      refreshAt(index, true, true);
    } catch (value) {
      reportError(value);
    }
  };

  const deleteSlide = () => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current || current.snapshot.slides.length <= 1) return;
    try {
      handle.deleteSlide(current.snapshot.slides[current.slideIndex].id);
      setSelection(null);
      setShapeSelection(null);
      setDragPreview(null);
      pointerGestureRef.current = null;
      recentShapeClickRef.current = null;
      refreshAt(Math.max(0, current.slideIndex - 1), true, true);
    } catch (value) {
      reportError(value);
    }
  };

  const addTextBox = () => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current) return;
    const slide = current.snapshot.slides[current.slideIndex];
    if (!slide) return;
    try {
      const receipt = handle.addTextBox(slide.id, {
        name: t('objects.defaultTextBoxName'),
        rect: {
          x: Math.round(current.snapshot.widthEmu * 0.15),
          y: Math.round(current.snapshot.heightEmu * 0.25),
          width: Math.round(current.snapshot.widthEmu * 0.7),
          height: Math.round(current.snapshot.heightEmu * 0.3),
        },
        text: '',
        style: textStyle,
      });
      const next = refreshAt(undefined, true);
      setShapeSelection(null);
      setDragPreview(null);
      pointerGestureRef.current = null;
      recentShapeClickRef.current = null;
      const shape = next?.snapshot.slides[next.slideIndex]?.shapes.find(
        (candidate) => candidate.id === receipt.shapeId
      );
      const story = shape?.textStories[0];
      if (story) {
        setSelection({ shapeId: shape.id, storyId: story.id, anchor: 0, focus: 0 });
        stageRef.current?.focus();
      }
    } catch (value) {
      reportError(value);
    }
  };

  const history = (direction: 'undo' | 'redo') => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      if (direction === 'undo') handle.undo();
      else handle.redo();
      setSelection(null);
      setShapeSelection(null);
      setDragPreview(null);
      pointerGestureRef.current = null;
      recentShapeClickRef.current = null;
      refreshAt(undefined, true);
    } catch (value) {
      reportError(value);
    }
  };

  const slideCount = model?.snapshot.slides.length ?? 0;
  const currentSlide = model?.slideIndex ?? 0;
  const shapeDragDelta =
    dragPreview && dragPreview.shapeId === shapeSelection?.shapeId ? dragPreview.delta : null;
  const selectedShapeMovable = selectedShape ? canMoveShape(selectedShape) : false;

  return (
    <div className={className} style={styles.root}>
      <div style={styles.toolbarShell}>
        <div style={styles.toolbar} role="group" aria-label={t('toolbar.label')}>
          <div style={styles.toolbarGroup} role="group" aria-label={t('toolbar.historyLabel')}>
            <button
              type="button"
              style={toolbarButton(handleRef.current?.canUndo() ?? false)}
              disabled={!handleRef.current?.canUndo()}
              aria-label={t('toolbar.undo')}
              title={t('toolbar.undo')}
              onClick={() => history('undo')}
            >
              <PptxToolbarIcon name="undo" />
            </button>
            <button
              type="button"
              style={toolbarButton(handleRef.current?.canRedo() ?? false)}
              disabled={!handleRef.current?.canRedo()}
              aria-label={t('toolbar.redo')}
              title={t('toolbar.redo')}
              onClick={() => history('redo')}
            >
              <PptxToolbarIcon name="redo" />
            </button>
          </div>
          <span style={styles.divider} aria-hidden="true" />
          <div
            style={styles.toolbarGroup}
            role="group"
            aria-label={t('toolbar.textFormattingLabel')}
          >
            <button
              type="button"
              aria-pressed={textStyle.bold}
              aria-label={t('toolbar.bold')}
              title={t('toolbar.bold')}
              style={formatButton(textStyle.bold)}
              onClick={() => applyFormatting({ bold: !textStyle.bold })}
            >
              B
            </button>
            <button
              type="button"
              aria-pressed={textStyle.italic}
              aria-label={t('toolbar.italic')}
              title={t('toolbar.italic')}
              style={{ ...formatButton(textStyle.italic), fontStyle: 'italic' }}
              onClick={() => applyFormatting({ italic: !textStyle.italic })}
            >
              I
            </button>
            <input
              aria-label={t('toolbar.fontSize')}
              title={t('toolbar.fontSize')}
              type="number"
              min={6}
              max={400}
              value={textStyle.fontSizePt}
              style={styles.numberInput}
              onChange={(event: ChangeEvent<HTMLInputElement>) => {
                const value = Number(event.target.value);
                if (Number.isFinite(value)) applyFormatting({ fontSizePt: value });
              }}
            />
            <input
              aria-label={t('toolbar.textColor')}
              title={t('toolbar.textColor')}
              type="color"
              value={textStyle.color}
              style={styles.colorInput}
              onChange={(event: ChangeEvent<HTMLInputElement>) =>
                applyFormatting({ color: event.target.value })
              }
            />
          </div>
          <span style={styles.divider} aria-hidden="true" />
          <div style={styles.toolbarGroup} role="group" aria-label={t('toolbar.slidesLabel')}>
            <button
              type="button"
              style={toolbarButton(true)}
              aria-label={t('toolbar.addSlide')}
              title={t('toolbar.addSlide')}
              onClick={addSlide}
            >
              <PptxToolbarIcon name="addSlide" />
            </button>
            <button
              type="button"
              style={toolbarButton(slideCount > 1)}
              disabled={slideCount <= 1}
              aria-label={t('toolbar.deleteSlide')}
              title={t('toolbar.deleteSlide')}
              onClick={deleteSlide}
            >
              <PptxToolbarIcon name="deleteSlide" />
            </button>
          </div>
          <span style={styles.divider} aria-hidden="true" />
          <div style={styles.toolbarGroup} role="group" aria-label={t('toolbar.objectsLabel')}>
            <button
              type="button"
              style={toolbarButton(slideCount > 0)}
              disabled={slideCount === 0}
              aria-label={t('toolbar.addTextBox')}
              title={t('toolbar.addTextBox')}
              onClick={addTextBox}
            >
              <PptxToolbarIcon name="textBox" />
            </button>
          </div>
        </div>
        {remotePeers.length > 0 ? (
          <div style={styles.presenceStrip} role="list" aria-label="Collaborators">
            {toolbarPresence.visible.map((peer) => {
              const sameSlide = peer.state.cursor?.slideId === currentSlideId;
              return (
                <span
                  key={peer.state.clientId}
                  role="listitem"
                  style={presenceChip(peer.state.user.color, sameSlide)}
                  title={peer.state.user.name}
                  aria-label={peer.state.user.name}
                >
                  {presenceInitials(peer.state.user.name)}
                </span>
              );
            })}
            {toolbarPresence.overflow > 0 ? (
              <span
                role="listitem"
                style={presenceChip('#64748b', true)}
                title={`${toolbarPresence.overflow} more collaborators`}
                aria-label={`${toolbarPresence.overflow} more collaborators`}
              >
                +{toolbarPresence.overflow}
              </span>
            ) : null}
          </div>
        ) : null}
      </div>
      <div style={styles.workspace}>
        <aside style={styles.slideStrip} aria-label={t('slides.panelLabel')}>
          {model?.snapshot.slides.map((slide, index) => {
            const slidePresence = remotePeersBySlide.get(slide.id);
            return (
              <button
                type="button"
                key={slide.id}
                aria-current={index === currentSlide ? 'page' : undefined}
                style={slideButton(index === currentSlide)}
                onClick={() => selectSlide(index)}
              >
                <span style={styles.slideNumber}>{index + 1}</span>
                <span style={styles.slidePreview}>
                  {model.thumbnails.get(slide.id) ? (
                    <SlideThumbnail
                      frame={model.thumbnails.get(slide.id)!}
                      resolveImage={(assetId) =>
                        resolveImage(assetId, handleRef, imageCacheRef, decodeImageError)
                      }
                    />
                  ) : (
                    <span style={styles.slideTitle}>
                      {slideTitle(slide.shapes) ||
                        t('slides.fallbackTitle', { number: index + 1 })}
                    </span>
                  )}
                  {slide.id !== currentSlideId &&
                  slidePresence &&
                  (slidePresence.visible.length > 0 || slidePresence.overflow > 0) ? (
                      <span style={styles.thumbnailPresence} aria-hidden="true">
                        {slidePresence.visible.map((peer) => (
                          <span
                            key={peer.state.clientId}
                            style={{
                              ...styles.thumbnailPresenceDot,
                              backgroundColor: peer.state.user.color,
                            }}
                          />
                        ))}
                        {slidePresence.overflow > 0 ? (
                          <span style={styles.thumbnailPresenceOverflow}>
                            +{slidePresence.overflow}
                          </span>
                        ) : null}
                      </span>
                    ) : null}
                  </span>
              </button>
            );
          })}
        </aside>
        <div
          ref={stageRef}
          style={styles.stage}
          tabIndex={0}
          role="application"
          aria-label={t('editor.appLabel')}
          onKeyDown={keyDown}
        >
          <div ref={canvasHostRef} style={styles.canvasHost}>
            {model?.frame ? (
              <div
                style={{
                  ...styles.canvasFrame,
                  width: model.frame.width * scale,
                  height: model.frame.height * scale,
                }}
              >
                <canvas
                  ref={canvasRef}
                  style={{
                    ...styles.canvas,
                    cursor: selection ? 'text' : selectedShapeMovable ? 'move' : 'default',
                  }}
                  onPointerDown={pointerDown}
                  onPointerMove={pointerMove}
                  onPointerUp={pointerUp}
                  onPointerCancel={cancelPointerGesture}
                  onLostPointerCapture={cancelPointerGesture}
                  onDoubleClick={pointerDoubleClick}
                  aria-label={
                    selectedShape
                      ? t('slides.canvasLabelWithSelection', {
                          current: currentSlide + 1,
                          total: slideCount,
                          name: selectedShape.name,
                        })
                      : t('slides.canvasLabel', {
                          current: currentSlide + 1,
                          total: slideCount,
                        })
                  }
                />
                <canvas ref={overlayCanvasRef} style={styles.canvasOverlay} aria-hidden="true" />
                {remoteShapePresence.visible.map(
                  ({ peer, peerCount, shapeId, bounds }, index) => (
                    <RemoteShapeOutline
                      key={shapeId}
                      peer={peer}
                      peerCount={peerCount}
                      bounds={bounds}
                      scale={scale}
                      labelOffset={index}
                    />
                  )
                )}
                {remoteShapePresence.overflow > 0 ? (
                  <span style={styles.remoteShapeOverflow} aria-hidden="true">
                    +{remoteShapePresence.overflow} selections
                  </span>
                ) : null}
                {selectedShapeBounds ? (
                  <span
                    style={{
                      ...styles.shapeSelection,
                      left: (selectedShapeBounds.x + (shapeDragDelta?.x ?? 0)) * scale,
                      top: (selectedShapeBounds.y + (shapeDragDelta?.y ?? 0)) * scale,
                      width: Math.max(1, selectedShapeBounds.width * scale),
                      height: Math.max(1, selectedShapeBounds.height * scale),
                    }}
                    aria-hidden="true"
                  />
                ) : null}
              </div>
            ) : (
              <div style={styles.empty}>
                {loading
                  ? t('editor.opening')
                  : file
                    ? t('editor.noSlides')
                    : t('editor.openPrompt')}
              </div>
            )}
          </div>
          {error ? <div style={styles.error}>{error}</div> : null}
        </div>
      </div>
    </div>
  );
}

function RemoteShapeOutline({
  peer,
  peerCount,
  bounds,
  scale,
  labelOffset,
}: {
  peer: PptxPresencePeer;
  peerCount: number;
  bounds: FrameBounds;
  scale: number;
  labelOffset: number;
}) {
  const [labelVisible, setLabelVisible] = useState(
    () => Date.now() - peer.cursorMovedAt < PRESENCE_LABEL_DURATION_MS
  );
  useEffect(() => {
    const remaining = PRESENCE_LABEL_DURATION_MS - (Date.now() - peer.cursorMovedAt);
    if (remaining <= 0) {
      setLabelVisible(false);
      return;
    }
    setLabelVisible(true);
    const timer = setTimeout(() => setLabelVisible(false), remaining);
    return () => clearTimeout(timer);
  }, [peer.cursorMovedAt]);

  return (
    <span
      style={{
        ...styles.remoteShapeSelection,
        left: bounds.x * scale,
        top: bounds.y * scale,
        width: Math.max(1, bounds.width * scale),
        height: Math.max(1, bounds.height * scale),
        borderColor: peer.state.user.color,
      }}
      aria-hidden="true"
    >
      <span
        style={{
          ...styles.remoteShapeWash,
          backgroundColor: peer.state.user.color,
        }}
      />
      {labelVisible ? (
        <span
          style={{
            ...styles.remoteShapeLabel,
            top: -4 - (labelOffset % 3) * 20,
            backgroundColor: peer.state.user.color,
          }}
        >
          {peer.state.user.name}
          {peerCount > 1 ? ` +${peerCount - 1}` : ''}
        </span>
      ) : null}
    </span>
  );
}

function SlideThumbnail({
  frame,
  resolveImage,
}: {
  frame: SlideDisplayList;
  resolveImage: CanvasImageResolver;
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const scale = 128 / frame.width;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(canvas, frame, dpr, scale);
    void paintSlide(ctx, frame, dpr, scale, { resolveImage }).catch(() => undefined);
  }, [frame, resolveImage]);
  return <canvas ref={canvasRef} style={styles.thumbnailCanvas} aria-hidden="true" />;
}

function clampSlideIndex(index: number, count: number): number {
  if (count === 0) return 0;
  return Math.max(0, Math.min(count - 1, index));
}

function useStableFontFaces(fonts: ReadonlyArray<PptxFontFace>): ReadonlyArray<PptxFontFace> {
  const stable = useRef(fonts);
  if (!fontFacesEqual(stable.current, fonts)) stable.current = fonts;
  return stable.current;
}

function fontFacesEqual(
  left: ReadonlyArray<PptxFontFace>,
  right: ReadonlyArray<PptxFontFace>
): boolean {
  if (left === right) return true;
  if (left.length !== right.length) return false;
  return left.every((face, index) => fontFaceEqual(face, right[index]));
}

function fontFaceEqual(left: PptxFontFace, right: PptxFontFace): boolean {
  return (
    left.family === right.family &&
    (left.bold ?? false) === (right.bold ?? false) &&
    (left.italic ?? false) === (right.italic ?? false) &&
    bytesEqual(left.bytes, right.bytes)
  );
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left === right) return true;
  if (left.byteLength !== right.byteLength) return false;
  return left.every((byte, index) => byte === right[index]);
}

function storyText(story: StorySnapshot): string {
  return story.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

function previousTextIndex(story: StorySnapshot, index: number): number {
  const prefix = storyText(story).slice(0, index);
  const characters = Array.from(prefix);
  const character = characters[characters.length - 1];
  return character ? index - character.length : index;
}

function nextTextIndex(story: StorySnapshot, index: number): number {
  const character = Array.from(storyText(story).slice(index))[0];
  return character ? index + character.length : index;
}

async function installBrowserFonts(fonts: ReadonlyArray<PptxFontFace>): Promise<FontFace[]> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return [];
  const installed = await Promise.all(
    fonts.map(async (font) => {
      const source = font.bytes.slice().buffer as ArrayBuffer;
      const face = new FontFace(font.family, source, {
        style: font.italic ? 'italic' : 'normal',
        weight: font.bold ? '700' : '400',
      });
      const loaded = await face.load();
      document.fonts.add(loaded);
      return loaded;
    })
  );
  return installed;
}

function removeBrowserFonts(fonts: FontFace[]): void {
  if (typeof document === 'undefined') return;
  for (const font of fonts) document.fonts.delete(font);
}

function resolveImage(
  assetId: string,
  handleRef: { current: PresentationHandle | null },
  cacheRef: { current: Map<string, Promise<CanvasImageSource | null>> },
  errorMessage: string
): Promise<CanvasImageSource | null> {
  const cached = cacheRef.current.get(assetId);
  if (cached) return cached;
  const pending = decodeImage(handleRef.current?.mediaBytes(assetId), errorMessage);
  cacheRef.current.set(assetId, pending);
  return pending;
}

async function decodeImage(
  bytes: Uint8Array | undefined,
  errorMessage: string
): Promise<CanvasImageSource | null> {
  if (!bytes) return null;
  const blob = new Blob([bytes.slice()]);
  if (typeof createImageBitmap === 'function') return createImageBitmap(blob);
  const url = URL.createObjectURL(blob);
  try {
    return await new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error(errorMessage));
      image.src = url;
    });
  } finally {
    URL.revokeObjectURL(url);
  }
}

function paintSelection(
  ctx: CanvasRenderingContext2D,
  frame: SlideDisplayList,
  selection: PptxTextSelection | null,
  dpr: number,
  scale: number
): void {
  if (!selection) return;
  const textBox = frame.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' &&
      primitive.storyId === selection.storyId &&
      primitive.shapeId === selection.shapeId
  );
  if (!textBox) return;
  const start = Math.min(selection.anchor, selection.focus);
  const end = Math.max(selection.anchor, selection.focus);
  ctx.save();
  ctx.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);
  if (start !== end) {
    ctx.fillStyle = 'rgba(59, 130, 246, 0.26)';
    for (const line of textBox.lines) {
      const lineStart = Math.max(start, line.start);
      const lineEnd = Math.min(end, line.end);
      if (lineStart >= lineEnd) continue;
      const x1 = caretX(line, lineStart);
      const x2 = caretX(line, lineEnd);
      ctx.fillRect(Math.min(x1, x2), line.y, Math.max(1, Math.abs(x2 - x1)), line.height);
    }
  } else {
    const line =
      textBox.lines.find((candidate) => start >= candidate.start && start <= candidate.end) ??
      textBox.lines[textBox.lines.length - 1];
    if (line) {
      const x = caretX(line, start);
      ctx.fillStyle = '#1d4ed8';
      ctx.fillRect(x, line.y, 1.5, line.height);
    }
  }
  ctx.restore();
}

function caretX(line: TextBoxPrimitive['lines'][number], position: number): number {
  const first = line.caretStops[0];
  if (!first) return line.x;
  return line.caretStops.reduce(
    (nearest, stop) =>
      Math.abs(stop.position - position) < Math.abs(nearest.position - position) ? stop : nearest,
    first
  ).x;
}

function slideTitle(shapes: DeckSnapshot['slides'][number]['shapes']): string {
  for (const shape of shapes) {
    for (const story of shape.textStories) {
      const value = storyText(story).trim();
      if (value) return value.slice(0, 40);
    }
  }
  return '';
}

const styles: Record<string, CSSProperties> = {
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    height: '100%',
    minHeight: 480,
    overflow: 'hidden',
    color: '#172033',
    background: '#f3f5f8',
    fontFamily: 'ui-sans-serif, system-ui, sans-serif',
  },
  toolbarShell: {
    display: 'flex',
    alignItems: 'center',
    flex: '0 0 auto',
    padding: '4px 0 5px',
    background: '#ffffff',
    borderBottom: '1px solid #e2e8f0',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: 0,
    flex: '1 1 auto',
    minWidth: 0,
    minHeight: 36,
    margin: '0 8px',
    padding: '4px 8px',
    background: '#f1f5f9',
    borderRadius: 999,
    boxSizing: 'border-box',
    overflowX: 'auto',
    overflowY: 'hidden',
  },
  presenceStrip: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: '0 0 auto',
    maxWidth: '35%',
    padding: '0 10px 0 2px',
    overflowX: 'auto',
  },
  toolbarGroup: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 1,
    flex: '0 0 auto',
  },
  divider: {
    width: 1,
    height: 24,
    margin: '0 6px',
    background: '#e2e8f0',
    flex: '0 0 auto',
  },
  numberInput: {
    appearance: 'textfield',
    width: 40,
    height: 28,
    marginLeft: 2,
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    padding: '0 4px',
    background: '#f8fafc',
    color: '#0f172a',
    font: '12px ui-sans-serif, system-ui, sans-serif',
    textAlign: 'center',
    outlineColor: '#2563eb',
  },
  colorInput: {
    width: 28,
    height: 28,
    marginLeft: 2,
    padding: 2,
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    background: '#f8fafc',
    cursor: 'pointer',
    outlineColor: '#2563eb',
  },
  workspace: { display: 'flex', flex: 1, minHeight: 0 },
  slideStrip: {
    width: 184,
    padding: '14px 10px',
    overflowY: 'auto',
    background: '#eef1f5',
    borderRight: '1px solid #d8dee9',
    boxSizing: 'border-box',
  },
  slideNumber: { width: 18, flex: '0 0 auto', paddingTop: 3, fontSize: 11, color: '#647087', textAlign: 'right' },
  slidePreview: {
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 132,
    aspectRatio: '16 / 9',
    padding: 8,
    overflow: 'hidden',
    background: '#ffffff',
    boxSizing: 'border-box',
    boxShadow: '0 1px 4px rgba(20, 31, 50, 0.16)',
  },
  slideTitle: { fontSize: 9, lineHeight: 1.25, color: '#39445a', textAlign: 'center' },
  thumbnailCanvas: { display: 'block', maxWidth: '100%', height: 'auto' },
  thumbnailPresence: {
    position: 'absolute',
    top: 4,
    right: 4,
    display: 'flex',
    gap: 3,
    padding: 2,
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.88)',
    boxShadow: '0 1px 3px rgba(15, 23, 42, 0.2)',
  },
  thumbnailPresenceDot: {
    width: 7,
    height: 7,
    borderRadius: 999,
    boxShadow: '0 0 0 1px rgba(255, 255, 255, 0.9)',
  },
  thumbnailPresenceOverflow: {
    color: '#475569',
    fontSize: 8,
    fontWeight: 700,
    lineHeight: '8px',
  },
  stage: { position: 'relative', flex: 1, minWidth: 0, outline: 'none', overflow: 'hidden' },
  canvasHost: { display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', overflow: 'auto' },
  canvasFrame: { position: 'relative', flex: '0 0 auto' },
  canvas: { display: 'block', flex: '0 0 auto', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' },
  canvasOverlay: { position: 'absolute', inset: 0, display: 'block', pointerEvents: 'none' },
  shapeSelection: {
    position: 'absolute',
    zIndex: 3,
    border: '2px solid #2563eb',
    boxSizing: 'border-box',
    boxShadow: '0 0 0 1px rgba(255, 255, 255, 0.9)',
    pointerEvents: 'none',
  },
  remoteShapeSelection: {
    position: 'absolute',
    zIndex: 2,
    border: '2px solid',
    borderRadius: 2,
    boxSizing: 'border-box',
    boxShadow: '0 0 0 1px rgba(255, 255, 255, 0.88)',
    pointerEvents: 'none',
  },
  remoteShapeWash: {
    position: 'absolute',
    inset: 0,
    opacity: 0.12,
  },
  remoteShapeLabel: {
    position: 'absolute',
    left: -2,
    maxWidth: 180,
    overflow: 'hidden',
    padding: '3px 6px',
    borderRadius: '4px 4px 4px 0',
    color: '#ffffff',
    fontSize: 11,
    fontWeight: 650,
    lineHeight: '14px',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    transform: 'translateY(-100%)',
    boxShadow: '0 1px 3px rgba(15, 23, 42, 0.25)',
  },
  remoteShapeOverflow: {
    position: 'absolute',
    top: 8,
    right: 8,
    zIndex: 3,
    padding: '3px 6px',
    borderRadius: 999,
    backgroundColor: '#475569',
    color: '#ffffff',
    fontSize: 10,
    fontWeight: 700,
    pointerEvents: 'none',
  },
  empty: { margin: 'auto', color: '#6b7587', fontSize: 14 },
  error: { position: 'absolute', left: 16, right: 16, bottom: 14, padding: '9px 12px', color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0', borderRadius: 6, fontSize: 12 },
};

function toolbarButton(enabled: boolean): CSSProperties {
  return {
    appearance: 'none',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 28,
    height: 28,
    padding: 0,
    border: 0,
    borderRadius: 6,
    color: enabled ? '#64748b' : '#94a3b8',
    background: 'transparent',
    cursor: enabled ? 'pointer' : 'default',
    opacity: enabled ? 1 : 0.32,
  };
}

function formatButton(active: boolean): CSSProperties {
  return {
    ...toolbarButton(true),
    fontWeight: 700,
    color: active ? '#ffffff' : '#64748b',
    background: active ? '#0f172a' : 'transparent',
  };
}

function presenceInitials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  return parts
    .slice(0, 2)
    .map((part) => part[0])
    .join('')
    .toUpperCase();
}

function presenceChip(color: string, sameSlide: boolean): CSSProperties {
  return {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 28,
    height: 28,
    flex: '0 0 auto',
    border: '2px solid #ffffff',
    borderRadius: 999,
    backgroundColor: color,
    boxShadow: '0 0 0 1px rgba(15, 23, 42, 0.14)',
    color: '#ffffff',
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: '-0.02em',
    opacity: sameSlide ? 1 : 0.38,
  };
}

function slideButton(active: boolean): CSSProperties {
  return {
    display: 'flex',
    alignItems: 'flex-start',
    gap: 7,
    width: '100%',
    marginBottom: 12,
    padding: 4,
    border: active ? '2px solid #325ee6' : '2px solid transparent',
    borderRadius: 5,
    background: 'transparent',
    cursor: 'pointer',
  };
}
