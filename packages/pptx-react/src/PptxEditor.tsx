import {
  initWasm,
  openPresentation,
  paintSlide,
  sizeCanvasForSlide,
} from '@betteroffice/pptx';
import type {
  CanvasImageResolver,
  DeckSnapshot,
  PptxFontFace,
  PresentationHandle,
  SlideDisplayList,
  StorySnapshot,
  TextBoxPrimitive,
  TextStylePatch,
} from '@betteroffice/pptx';
import type { CollaborationReplica } from '@betteroffice/pptx/collaboration';
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
  KeyboardEvent,
  MouseEvent as ReactMouseEvent,
  PointerEvent,
} from 'react';
import { EditorToolbar } from './components/EditorToolbar';
import type {
  FormattingAction,
  PptxEditorTool,
  PptxZoom,
  SelectionFormatting,
  SlideLayoutOption,
} from './components/Toolbar';
import {
  canMoveShape,
  findShape,
  findTopLevelShape,
  frameBoundsForShape,
  movedShapePosition,
  passedDragThreshold,
  slidePoint,
  textPositionAtPoint,
} from './interactions';
import type { FrameBounds, SlidePoint } from './interactions';
import {
  effectiveStyleFromSelection,
  selectionFormattingFromStory,
  storyFormattingFromStory,
  storyTextRanges,
} from './textFormatting';
import type { EffectiveTextStyle } from './textFormatting';

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
    }
  | {
      kind: 'textBox';
      pointerId: number;
      slideId: string;
      start: SlidePoint;
      last: SlidePoint;
    };

interface ShapeDragPreview {
  shapeId: string;
  delta: SlidePoint;
}

interface TextBoxPreview {
  start: SlidePoint;
  end: SlidePoint;
}

interface RecentShapeClick {
  slideId: string;
  shapeId: string;
  clientX: number;
  clientY: number;
  timeStamp: number;
}

const initialStyle: EffectiveTextStyle = {
  bold: false,
  italic: false,
  underline: 'none',
  fontSizePt: 24,
  color: '#111827',
  fontFamily: 'Arial',
};

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
  const [textBoxPreview, setTextBoxPreview] = useState<TextBoxPreview | null>(null);
  const [textStyle, setTextStyle] = useState(initialStyle);
  const [historyState, setHistoryState] = useState({ canUndo: false, canRedo: false });
  const [zoom, setZoom] = useState<PptxZoom>('fit');
  const [activeTool, setActiveTool] = useState<PptxEditorTool>('select');
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);

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
            (gesture.kind !== 'textBox' &&
              !findShape(activeSlide.shapes, gesture.shapeId)))
        ) {
          pointerGestureRef.current = null;
          recentShapeClickRef.current = null;
          setDragPreview(null);
          setTextBoxPreview(null);
        }
        modelRef.current = next;
        setModel(next);
        setHistoryState({ canUndo: handle.canUndo(), canRedo: handle.canRedo() });
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
    setTextBoxPreview(null);
    setHistoryState({ canUndo: false, canRedo: false });
    setActiveTool('select');
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
    const host = canvasHostRef.current;
    if (!host) return;
    const update = () => setViewport({ width: host.clientWidth, height: host.clientHeight });
    update();
    const observer = new ResizeObserver(update);
    observer.observe(host);
    return () => observer.disconnect();
  }, []);

  const fitScale = useMemo(() => {
    if (!model?.frame || viewport.width <= 0 || viewport.height <= 0) return 1;
    return Math.min(
      (viewport.width - 40) / model.frame.width,
      (viewport.height - 40) / model.frame.height,
      1
    );
  }, [model?.frame, viewport]);
  const scale = zoom === 'fit' ? fitScale : zoom;

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
  const selectedShapeStoryId = selectedShape?.textStories[0]?.id ?? null;

  const selectedShapeBounds = useMemo<FrameBounds | null>(
    () =>
      model?.frame && selectedShape
        ? frameBoundsForShape(model.snapshot, model.frame, selectedShape)
        : null,
    [model, selectedShape]
  );

  useEffect(() => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    try {
      const story = handle.story(selection.storyId);
      setTextStyle(
        effectiveStyleFromSelection(
          story,
          selection.anchor,
          selection.focus,
          initialStyle
        )
      );
    } catch (value) {
      reportError(value);
    }
  }, [model, reportError, selection]);

  const selectionFormatting = useMemo<SelectionFormatting>(() => {
    if (selection && selection.anchor === selection.focus) {
      return selectionFormattingFromStyle(textStyle);
    }
    const handle = handleRef.current;
    if (!handle) return selectionFormattingFromStyle(textStyle);
    try {
      if (!selection) {
        return selectedShapeStoryId
          ? storyFormattingFromStory(handle.story(selectedShapeStoryId), initialStyle)
          : selectionFormattingFromStyle(textStyle);
      }
      return selectionFormattingFromStory(
        handle.story(selection.storyId),
        selection.anchor,
        selection.focus,
        initialStyle
      );
    } catch {
      return selectionFormattingFromStyle(textStyle);
    }
  }, [model, selectedShapeStoryId, selection, textStyle]);

  const slideLayouts = useMemo<SlideLayoutOption[]>(() => {
    const unique = new Set<string | null>();
    for (const slide of model?.snapshot.slides ?? []) unique.add(slide.layoutPartPath);
    return [...unique].map((partPath) => ({ partPath }));
  }, [model?.snapshot]);

  const selectSlide = (index: number) => {
    setSelection(null);
    setShapeSelection(null);
    setDragPreview(null);
    setTextBoxPreview(null);
    pointerGestureRef.current = null;
    recentShapeClickRef.current = null;
    refreshAt(index);
    stageRef.current?.focus();
  };

  const createTextBox = (start: SlidePoint, end: SlidePoint) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current?.frame) return;
    const slide = current.snapshot.slides[current.slideIndex];
    if (!slide) return;
    const dragged = Math.abs(end.x - start.x) >= 6 || Math.abs(end.y - start.y) >= 6;
    const width = dragged ? Math.abs(end.x - start.x) : current.frame.width * 0.36;
    const height = dragged ? Math.abs(end.y - start.y) : current.frame.height * 0.12;
    const x = Math.max(
      0,
      Math.min(
        dragged ? Math.min(start.x, end.x) : start.x,
        current.frame.width - width
      )
    );
    const y = Math.max(
      0,
      Math.min(
        dragged ? Math.min(start.y, end.y) : start.y,
        current.frame.height - height
      )
    );
    try {
      const receipt = handle.addTextBox(slide.id, {
        name: t('objects.defaultTextBoxName'),
        rect: {
          x: Math.round((x * current.snapshot.widthEmu) / current.frame.width),
          y: Math.round((y * current.snapshot.heightEmu) / current.frame.height),
          width: Math.round(
            (Math.max(12, width) * current.snapshot.widthEmu) / current.frame.width
          ),
          height: Math.round(
            (Math.max(12, height) * current.snapshot.heightEmu) / current.frame.height
          ),
        },
        text: '',
        style: textStyle,
      });
      const next = refreshAt(undefined, true);
      setActiveTool('select');
      setShapeSelection(null);
      setDragPreview(null);
      setTextBoxPreview(null);
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
    const slide = current.snapshot.slides[current.slideIndex];
    if (activeTool === 'textBox' && slide) {
      setSelection(null);
      setShapeSelection(null);
      setDragPreview(null);
      setTextBoxPreview({ start: point, end: point });
      recentShapeClickRef.current = null;
      pointerGestureRef.current = {
        kind: 'textBox',
        pointerId: event.pointerId,
        slideId: slide.id,
        start: point,
        last: point,
      };
      event.currentTarget.setPointerCapture(event.pointerId);
      stageRef.current?.focus();
      event.preventDefault();
      return;
    }
    try {
      handle.layoutSlide(current.slideIndex);
      const hit = handle.hitTest(point.x, point.y);
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
    } else if (gesture.kind === 'shape') {
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
    } else {
      gesture.last = point;
      setTextBoxPreview({ start: gesture.start, end: point });
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
    if (gesture.kind === 'textBox') {
      setTextBoxPreview(null);
      createTextBox(gesture.start, gesture.last);
      event.preventDefault();
      return;
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
    setTextBoxPreview(null);
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
    if (!handle) return;
    const modifier = event.metaKey || event.ctrlKey;
    if (event.key === 'Escape') {
      setActiveTool('select');
      setTextBoxPreview(null);
      pointerGestureRef.current = null;
      event.preventDefault();
      return;
    }
    if (modifier && (event.key === 'z' || event.key === 'Z')) {
      event.preventDefault();
      history(event.shiftKey ? 'redo' : 'undo');
      return;
    }
    if (!selection && !selectedShapeStoryId) return;
    if (modifier && (event.key === 'b' || event.key === 'B')) {
      event.preventDefault();
      applyFormatting({
        bold: selection ? !textStyle.bold : !selectionFormatting.bold,
      });
      return;
    }
    if (modifier && (event.key === 'i' || event.key === 'I')) {
      event.preventDefault();
      applyFormatting({
        italic: selection ? !textStyle.italic : !selectionFormatting.italic,
      });
      return;
    }
    if (modifier && (event.key === 'u' || event.key === 'U')) {
      event.preventDefault();
      applyFormatting({
        underline: selection
          ? textStyle.underline === 'none'
            ? 'sng'
            : 'none'
          : selectionFormatting.underline
            ? 'none'
            : 'sng',
      });
      return;
    }
    if (!selection) return;
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
    if (!handle) return;
    try {
      if (selection) {
        if (selection.anchor === selection.focus) return;
        handle.formatText(
          selection.storyId,
          Math.min(selection.anchor, selection.focus),
          Math.max(selection.anchor, selection.focus),
          patch
        );
      } else if (selectedShapeStoryId) {
        const story = handle.story(selectedShapeStoryId);
        formatStory(handle, story, patch);
      } else {
        return;
      }
      refreshAt(undefined, true);
    } catch (value) {
      reportError(value);
    }
  };

  const formatSelection = (action: FormattingAction) => {
    if (action === 'bold') {
      applyFormatting({ bold: !selectionFormatting.bold });
    } else if (action === 'italic') {
      applyFormatting({ italic: !selectionFormatting.italic });
    } else if (action === 'underline') {
      applyFormatting({ underline: selectionFormatting.underline ? 'none' : 'sng' });
    } else if (action.type === 'fontFamily') {
      applyFormatting({ fontFamily: action.value });
    } else if (action.type === 'fontSize') {
      applyFormatting({ fontSizePt: action.value });
    } else if (action.type === 'textColor') {
      applyFormatting({ color: action.value });
    }
  };

  const addSlide = (layoutPartPath?: string | null) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current) return;
    try {
      const index = current.slideIndex + 1;
      const layout =
        layoutPartPath === undefined
          ? current.snapshot.slides[current.slideIndex]?.layoutPartPath ?? undefined
          : layoutPartPath ?? undefined;
      handle.insertSlide(index, layout);
      setSelection(null);
      setShapeSelection(null);
      setDragPreview(null);
      setTextBoxPreview(null);
      setActiveTool('select');
      pointerGestureRef.current = null;
      recentShapeClickRef.current = null;
      refreshAt(index, true, true);
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
      setTextBoxPreview(null);
      setActiveTool('select');
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
      <EditorToolbar
        currentFormatting={selectionFormatting}
        textSelectionActive={selection !== null || selectedShapeStoryId !== null}
        onFormat={formatSelection}
        onInsertSlide={addSlide}
        slideLayouts={slideLayouts}
        currentLayoutPartPath={model?.snapshot.slides[currentSlide]?.layoutPartPath}
        onUndo={() => history('undo')}
        onRedo={() => history('redo')}
        canUndo={historyState.canUndo}
        canRedo={historyState.canRedo}
        zoom={zoom}
        onZoomChange={setZoom}
        activeTool={activeTool}
        onToolChange={(tool) => {
          setActiveTool(tool);
          setTextBoxPreview(null);
          pointerGestureRef.current = null;
          stageRef.current?.focus();
        }}
        disabled={!model || slideCount === 0}
        style={styles.toolbarShell}
      >
        <EditorToolbar.Toolbar />
      </EditorToolbar>
      <div style={styles.workspace}>
        <aside style={styles.slideStrip} aria-label={t('slides.panelLabel')}>
          {model?.snapshot.slides.map((slide, index) => (
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
              </span>
            </button>
          ))}
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
                  data-testid="pptx-slide-canvas"
                  style={{
                    ...styles.canvas,
                    cursor:
                      activeTool === 'textBox'
                        ? 'crosshair'
                        : selection
                          ? 'text'
                          : selectedShapeMovable
                            ? 'move'
                            : 'default',
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
                {textBoxPreview ? (
                  <span
                    style={{
                      ...styles.textBoxPreview,
                      left: Math.min(textBoxPreview.start.x, textBoxPreview.end.x) * scale,
                      top: Math.min(textBoxPreview.start.y, textBoxPreview.end.y) * scale,
                      width:
                        Math.max(1, Math.abs(textBoxPreview.end.x - textBoxPreview.start.x)) *
                        scale,
                      height:
                        Math.max(1, Math.abs(textBoxPreview.end.y - textBoxPreview.start.y)) *
                        scale,
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

function formatStory(
  handle: PresentationHandle,
  story: StorySnapshot,
  patch: TextStylePatch
): void {
  try {
    handle.formatText(story.id, 0, story.length, patch);
    return;
  } catch (value) {
    if (
      !(value instanceof Error) ||
      !value.message.includes('crosses a paragraph boundary')
    ) {
      throw value;
    }
  }
  for (const range of storyTextRanges(story)) {
    if (range.start < range.end) {
      handle.formatText(story.id, range.start, range.end, patch);
    }
  }
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

function selectionFormattingFromStyle(style: EffectiveTextStyle): SelectionFormatting {
  return {
    bold: style.bold,
    italic: style.italic,
    underline: style.underline !== 'none',
    fontSize: style.fontSizePt,
    textColor: style.color,
    fontFamily: style.fontFamily,
  };
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
    flex: '0 0 auto',
    padding: '4px 0 5px',
    background: '#ffffff',
    borderBottom: '1px solid #e2e8f0',
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
  stage: { position: 'relative', flex: 1, minWidth: 0, outline: 'none', overflow: 'hidden' },
  canvasHost: { display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', overflow: 'auto' },
  canvasFrame: { position: 'relative', flex: '0 0 auto' },
  canvas: { display: 'block', flex: '0 0 auto', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' },
  canvasOverlay: { position: 'absolute', inset: 0, display: 'block', pointerEvents: 'none' },
  shapeSelection: {
    position: 'absolute',
    border: '2px solid #2563eb',
    boxSizing: 'border-box',
    boxShadow: '0 0 0 1px rgba(255, 255, 255, 0.9)',
    pointerEvents: 'none',
  },
  textBoxPreview: {
    position: 'absolute',
    border: '1.5px dashed #1a73e8',
    background: 'rgba(26, 115, 232, 0.08)',
    boxSizing: 'border-box',
    pointerEvents: 'none',
  },
  empty: { margin: 'auto', color: '#6b7587', fontSize: 14 },
  error: { position: 'absolute', left: 16, right: 16, bottom: 14, padding: '9px 12px', color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0', borderRadius: 6, fontSize: 12 },
};

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
