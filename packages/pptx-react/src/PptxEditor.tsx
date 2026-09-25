import {
  initWasm,
  openPresentation,
  paintSlide,
  presentationImageBlob,
  PRESENCE_LABEL_DURATION_MS,
  sizeCanvasForSlide,
  slideToPng,
  StaleProposalError,
} from '@betteroffice/pptx';
import type {
  CanvasImageResolver,
  CollaborationReplica,
  DeckSnapshot,
  HitTestResult,
  ParagraphAlignment,
  PptxCaretAnchor,
  PptxPresence,
  PptxPresencePeer,
  PptxFontFace,
  PresentationHandle,
  Proposal,
  ProposalDiffSlide,
  SlideDisplayList,
  StorySnapshot,
  TextBoxPrimitive,
  TextStylePatch,
} from '@betteroffice/pptx';
import type { Translations } from '@betteroffice/pptx-i18n';
import { LocaleProvider, useTranslation } from './i18n';
import { ProposalsPanel } from './components/ProposalsPanel';
import { ProposalCanvasOverlay, ProposalCanvasToolbar, ProposalNotesDiff, useProposalCanvas } from './components/ProposalCanvas';
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
  PointerEvent,
  ReactNode,
} from 'react';
import { EditorToolbar } from './components/EditorToolbar';
import { EditorChromeContext } from './components/EditorToolbarContext';
import { PresentationOverlay } from './components/PresentationOverlay';
import type {
  PptxEditorTool,
  PptxShapePreset,
  PptxZoom,
  ShapeFormattingAction,
} from './components/toolbarTypes';
import { SHAPE_PRESETS, shapePresetFromTool } from './components/toolbarTypes';
import { createPptxCommandController } from './commands/createPptxCommandStore';
import { usePptxCommands, usePptxCommandState } from './commands/hooks';
import { useDisabledDescription } from './components/ui/ToolbarPrimitives';
import {
  createInputCoordinator,
  PptxInputFailure,
  PptxInputStale,
} from './commands/inputCoordinator';
import { PptxCommandContext } from './commands/PptxCommandProvider';
import type { PptxCommandStore } from './commands/types';
import { useCommandShortcuts } from './commands/useCommandShortcuts';
import {
  usePptxCommandBinding,
  type PptxExportOutcome,
  type PptxSaveOutcome,
} from './commands/usePptxCommands';
import {
  RESIZE_HANDLES,
  canResizeShape,
  canMoveShape,
  findShape,
  findTopLevelShape,
  frameBoundsForShape,
  gestureOwnsPointer,
  pointerTargetAtPoint,
  indexShapes,
  movedShapePosition,
  passedDragThreshold,
  handleAnchor,
  resizeCommitDelta,
  resizeCursor,
  resizedShapeBox,
  resizedShapeBounds,
  shapeTargetExists,
  slidePoint,
  textLocationAtPoint,
} from './interactions';
import type { FrameBounds, HoverTarget, ResizeHandle, SlidePoint } from './interactions';
import {
  groupPresenceBySlide,
  groupShapePresence,
  limitPresence,
  type BoundedPresence,
} from './presence-rendering';
import { effectiveStyleFromSelection } from './textFormatting';
import { useSyncedState } from './useSyncedState';
import type { EffectiveTextStyle } from './textFormatting';
import {
  caretGoalX,
  caretLineIndex,
  extendTextRange,
  lineEdge,
  sameCaretGoalKey,
  textRangeAt,
  verticalCaretMove,
  wordBoundary,
} from './textSelection';
import type { CaretLine, TextSelectionGranularity } from './textSelection';

export interface PptxTextSelection {
  shapeId: string;
  storyId: string;
  anchor: number;
  focus: number;
  focusLine?: number;
}

export interface PptxTextSelectionTarget {
  /** 1-based slide number. */
  slide: number;
  shapeId: string;
  storyId: string;
  start: number;
  end: number;
}

export type PptxPointPosition = HitTestResult & { slide: number; slideId: string };

export interface PptxEditorApi {
  /** Waits for accepted input; rejects during unfinished pointer gestures. */
  flushPendingInput: () => Promise<void>;
  /** Client coordinates; returns null outside slide content. */
  getPositionAtPoint: (clientX: number, clientY: number) => PptxPointPosition | null;
  clearSelection: () => void;
  focus: () => void;
  /** Accepts a 1-based slide number. */
  goToSlide: (slide: number) => boolean;
  handle: PresentationHandle;
  /** The editor's commands, shared by built-in and host chrome; `api.handle` edits bypass them. */
  readonly commands: PptxCommandStore;
  refresh: () => void;
  refreshProposals: () => void;
  /** Serialize the presentation back to .pptx bytes, edits included. */
  save: () => Uint8Array;
  /** Also navigates to the slide. */
  selectText: (target: PptxTextSelectionTarget) => boolean;
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
  /** Download name for the save button; falls back to `presentation.pptx`. */
  fileName?: string;
  /** 1-based; clamped to the deck. */
  initialSlide?: number;
  onReady?: (api: PptxEditorApi) => void;
  onChange?: (snapshot: DeckSnapshot) => void;
  onError?: (error: Error) => void;
  /** Receives the saved bytes; without it, saving downloads the file. */
  onSave?: (bytes: Uint8Array) => void;
  /** Return true for built-in saving; false or void handles/cancels the request. */
  onSaveRequest?: () => boolean | void | Promise<boolean | void>;
  /** Blocks user edits; navigation and selection remain available. */
  readOnly?: boolean;
  /**
   * Replaces the toolbar: omit it for the default controls, pass `null` for
   * none, or pass chrome composed from the toolbar parts, which also renders
   * while `readOnly`. `showToolbar={false}` hides the whole region.
   */
  toolbar?: ReactNode;
  /** Shows the toolbar region; defaults to true. */
  showToolbar?: boolean;
}

interface PictureOrigin {
  handle: PresentationHandle;
  generation: number;
  slideId: string | null;
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
      focus: number;
      focusLine?: number;
      granularity: 'character' | TextSelectionGranularity;
      clickCount: number;
      clickShapeId: string;
      startClientX: number;
      startClientY: number;
      dragging: boolean;
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
      clickCount: number;
      dragging: boolean;
    }
  | {
      kind: 'textBox';
      pointerId: number;
      slideId: string;
      start: SlidePoint;
      last: SlidePoint;
    }
  | {
      kind: 'shapeInsert';
      pointerId: number;
      slideId: string;
      geometry: PptxShapePreset;
      start: SlidePoint;
      last: SlidePoint;
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

interface TextBoxPreview {
  start: SlidePoint;
  end: SlidePoint;
}

interface RecentCanvasClick {
  slideId: string;
  shapeId: string;
  clientX: number;
  clientY: number;
  timeStamp: number;
  count: number;
}

const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

// the png download name derived from the deck name: swap .pptx for the slide.
function pngName(fileName: string | undefined, slideIndex: number): string {
  const stem = (fileName ?? 'presentation.pptx').replace(/\.pptx$/i, '');
  return `${stem}-slide-${slideIndex + 1}.png`;
}

// trigger a browser download of a byte blob under the given name and mime type.
function downloadBytes(bytes: Uint8Array, name: string, mime: string): void {
  const blob = new Blob([new Uint8Array(bytes)], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** Reads a file as a `data:` URL, e.g. for handing an image to `<img>`/`Image`. */
function readFileAsDataUrl(file: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error ?? new Error('failed to read the file'));
    reader.readAsDataURL(file);
  });
}

/** The pixel size a `data:` image URL decodes to. */
function loadImageSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
    image.onerror = () => reject(new Error('failed to decode the image'));
    image.src = dataUrl;
  });
}

const MAX_INSERT_IMAGE_BYTES = 8 * 1024 * 1024;
const INSERT_IMAGE_TYPES: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', webp: 'image/webp', svg: 'image/svg+xml',
};

/** Windows/Office caret phase. */
const CARET_BLINK_MS = 530;

/** Screen-space diameter of a resize grip. */
const HANDLE_SIZE = 9;

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
      <PptxEditorContent {...props} i18n={i18n} />
    </LocaleProvider>
  );
}

function PptxEditorContent({
  file,
  fonts,
  clientId,
  collaboration,
  className,
  fileName,
  initialSlide,
  onReady,
  onChange,
  onError,
  onSave,
  onSaveRequest,
  readOnly = false,
  toolbar,
  showToolbar = true,
  i18n,
}: PptxEditorProps) {
  const { t } = useTranslation();
  const decodeImageError = t('errors.decodeSlideImage');
  const collaborationClientId = collaboration?.clientId ?? clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const collaborationOnReplica = collaboration?.onReplica;
  const collaborationPresence = collaboration?.presence;
  const handleRef = useRef<PresentationHandle | null>(null);
  const modelRef = useRef<EditorModel | null>(null);
  const generationRef = useRef(0);
  const [commands] = useState(createPptxCommandController);
  const commandStore = commands.store;
  const typingRef = useRef<TypingTarget | null>(null);
  const [coordinator] = useState(() => {
    const created = createInputCoordinator(() => {
      if (!created.busy()) typingRef.current = null;
      commands.refresh();
    });
    return created;
  });
  const pendingSaveRef = useRef<Promise<PptxSaveOutcome> | null>(null);
  const pickerOriginRef = useRef<PictureOrigin | null>(null);
  const hostPointRef = useRef<(x: number, y: number) => PptxPointPosition | null>(() => null);
  const flushPendingInput = useCallback(async (opened: PresentationHandle) => {
    if (handleRef.current !== opened) throw new Error('Presentation is no longer open');
    try {
      await coordinator.run(() => {
        if (handleRef.current !== opened) throw new Error('Presentation changed while flushing input');
        if (pointerGestureRef.current || resizeRef.current) {
          throw new Error('Finish the pointer gesture before flushing input');
        }
      });
    } catch (value) {
      if (value instanceof PptxInputStale) throw new Error('Presentation changed while flushing input');
      throw value instanceof PptxInputFailure ? value.cause : value;
    }
  }, [coordinator]);
  const initialSlideRef = useRef(initialSlide);
  const onReadyRef = useRef(onReady);
  const onChangeRef = useRef(onChange);
  const onErrorRef = useRef(onError);
  const rootRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const canvasHostRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pictureInputRef = useRef<HTMLInputElement>(null);
  const [stageFocused, setStageFocused] = useState(false);
  const caretGoalRef = useRef<{
    shapeId: string;
    position: number;
    lineIndex?: number;
    goalX: number;
  } | null>(null);
  const resizeRef = useRef<{
    pointerId: number;
    handle: ResizeHandle;
    start: SlidePoint;
    slideId: string;
    shapeId: string;
    delta: SlidePoint;
  } | null>(null);
  const [resizeDelta, setResizeDelta] = useState<SlidePoint | null>(null);
  const pointerGestureRef = useRef<PointerGesture | null>(null);
  const recentClickRef = useRef<RecentCanvasClick | null>(null);
  const imageCacheRef = useRef(new Map<string, Promise<CanvasImageSource | null>>());
  const stableFonts = useStableFontFaces(fonts);
  const [model, setModel] = useState<EditorModel | null>(null);
  const [selection, setSelection, selectionRef] = useSyncedState<PptxTextSelection | null>(null);
  const [shapeSelection, setShapeSelection, shapeSelectionRef] =
    useSyncedState<PptxShapeSelection | null>(null);
  const [dragPreview, setDragPreview] = useState<ShapeDragPreview | null>(null);
  const [textBoxPreview, setTextBoxPreview] = useState<TextBoxPreview | null>(null);
  const [textStyle, setTextStyle, textStyleRef] = useSyncedState(initialStyle);
  const [zoom, setZoom, zoomRef] = useSyncedState<PptxZoom>('fit');
  const [activeTool, setActiveTool, activeToolRef] = useSyncedState<PptxEditorTool>('select');
  const [hoverTarget, setHoverTarget] = useState<HoverTarget | null>(null);
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);
  const [remotePeers, setRemotePeers] = useState<readonly PptxPresencePeer[]>([]);
  const [presenting, setPresenting] = useState(false);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsOpen, setProposalsOpen, proposalsOpenRef] = useSyncedState(false);
  const proposalButtonRef = useRef<HTMLButtonElement>(null);
  const canvasReview = useProposalCanvas(
    handleRef.current,
    readOnly ? [] : proposals,
    model?.snapshot,
    model?.slideIndex ?? 0
  );
  const [paintedReview, setPaintedReview] = useState<ProposalDiffSlide | null>(null);
  const resolveProposalImage = useCallback((assetId: string) =>
    resolveImage(assetId, handleRef, imageCacheRef, decodeImageError), [decodeImageError]);

  useEffect(() => {
    if (!canvasReview.reviewing) return;
    setSelection(null);
    setShapeSelection(null);
    pointerGestureRef.current = null;
    resizeRef.current = null;
    setResizeDelta(null);
    setDragPreview(null);
    setTextBoxPreview(null);
  }, [canvasReview.reviewing]);

  useEffect(() => {
    if (!readOnly) return;
    setActiveTool('select');
    setProposalsOpen(false);
    pointerGestureRef.current = null;
    resizeRef.current = null;
    recentClickRef.current = null;
    setResizeDelta(null);
    setDragPreview(null);
    setTextBoxPreview(null);
  }, [readOnly]);

  onReadyRef.current = onReady;
  initialSlideRef.current = initialSlide;
  onChangeRef.current = onChange;
  onErrorRef.current = onError;
  modelRef.current = model;
  const readOnlyRef = useRef(readOnly);
  readOnlyRef.current = readOnly;
  const reviewingRef = useRef(canvasReview.reviewing);
  reviewingRef.current = canvasReview.reviewing;

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
        caretGoalRef.current = null;
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
              gesture.kind !== 'shapeInsert' &&
              !findShape(activeSlide.shapes, gesture.shapeId)))
        ) {
          pointerGestureRef.current = null;
          recentClickRef.current = null;
          setDragPreview(null);
          setTextBoxPreview(null);
        }
        const resize = resizeRef.current;
        if (resize && !shapeTargetExists(activeSlide, resize)) {
          resizeRef.current = null;
          setResizeDelta(null);
        }
        modelRef.current = next;
        setModel(next);
        setProposals(handle.listProposals());
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

  const clearSelection = useCallback(() => {
    typingRef.current = null;
    caretGoalRef.current = null;
    resizeRef.current = null;
    pointerGestureRef.current = null;
    recentClickRef.current = null;
    setResizeDelta(null);
    setSelection(null);
    setShapeSelection(null);
    setDragPreview(null);
    setTextBoxPreview(null);
  }, []);

  const goToSlide = useCallback(
    (slide: number): boolean => {
      const current = modelRef.current;
      if (
        !current ||
        !Number.isInteger(slide) ||
        slide < 1 ||
        slide > current.snapshot.slides.length
      ) {
        return false;
      }
      clearSelection();
      if (!refreshAt(slide - 1)) return false;
      stageRef.current?.focus();
      return true;
    },
    [clearSelection, refreshAt]
  );

  const selectText = useCallback(
    (target: PptxTextSelectionTarget): boolean => {
      const handle = handleRef.current;
      const current = modelRef.current;
      if (
        !handle ||
        !current ||
        !Number.isInteger(target.slide) ||
        !Number.isInteger(target.start) ||
        !Number.isInteger(target.end) ||
        target.start < 0 ||
        target.end < target.start
      ) {
        return false;
      }
      const slide = current.snapshot.slides[target.slide - 1];
      const shape = slide ? findShape(slide.shapes, target.shapeId) : null;
      if (!shape?.textStories.some((story) => story.id === target.storyId)) return false;
      try {
        const story = handle.story(target.storyId);
        if (target.end > story.length) return false;
        clearSelection();
        if (!refreshAt(target.slide - 1)) return false;
        setSelection({
          shapeId: target.shapeId,
          storyId: target.storyId,
          anchor: target.start,
          focus: target.end,
        });
        stageRef.current?.focus();
        return true;
      } catch {
        return false;
      }
    },
    [clearSelection, refreshAt]
  );

  const navigateProposalTarget = useCallback((slideId: string, shapeId: string | null, proposalId?: string) => {
    const index = modelRef.current?.snapshot.slides.findIndex((slide) => slide.id === slideId) ?? -1;
    if (index < 0) return;
    const refreshed = refreshAt(index);
    const slide = refreshed?.snapshot.slides[refreshed.slideIndex];
    setSelection(null);
    setShapeSelection(shapeId && slide && findShape(slide.shapes, shapeId) ? { slideId, shapeId } : null);
    if (proposalId) canvasReview.select(proposalId);
    canvasReview.setEnabled(true);
  }, [refreshAt, canvasReview.select, canvasReview.setEnabled]);

  const refreshProposals = useCallback(() => {
    try {
      setProposals(handleRef.current?.listProposals() ?? []);
    } catch (value) {
      reportError(value);
    }
  }, [reportError]);

  const acceptProposal = useCallback((id: string, force: boolean): 'accepted' | 'stale' => {
    const handle = handleRef.current;
    if (!handle) throw new Error('Presentation is no longer open');
    try {
      handle.acceptProposal(id, { force });
      refreshAt(undefined, true, true);
      return 'accepted';
    } catch (value) {
      refreshProposals();
      if (value instanceof StaleProposalError) return 'stale';
      throw value;
    }
  }, [refreshAt, refreshProposals]);

  const rejectProposal = useCallback((id: string): boolean => {
    const rejected = handleRef.current?.rejectProposal(id) ?? false;
    refreshProposals();
    return rejected;
  }, [refreshProposals]);

  useEffect(() => {
    let disposed = false;
    let handle: PresentationHandle | null = null;
    let browserFaces: FontFace[] = [];
    let unsubscribeUpdates = () => {};
    handleRef.current?.dispose();
    handleRef.current = null;
    generationRef.current += 1;
    coordinator.reset();
    typingRef.current = null;
    setCollaborationReplica(null);
    modelRef.current = null;
    setModel(null);
    setSelection(null);
    setShapeSelection(null);
    setDragPreview(null);
    setTextBoxPreview(null);
    setResizeDelta(null);
    setActiveTool('select');
    setPresenting(false);
    setProposals([]);
    setProposalsOpen(false);
    pointerGestureRef.current = null;
    resizeRef.current = null;
    caretGoalRef.current = null;
    recentClickRef.current = null;
    setError(null);
    pendingSaveRef.current = null;
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
          const requestedSlide = initialSlideRef.current;
          refreshAt(
            typeof requestedSlide === 'number' && Number.isInteger(requestedSlide)
              ? requestedSlide - 1
              : 0
          );
          setLoading(false);
          setCollaborationReplica(handle);
          const opened = handle;
          onReadyRef.current?.({
            clearSelection,
            commands: commandStore,
            goToSlide,
            handle: opened,
            refresh,
            refreshProposals,
            flushPendingInput: () => flushPendingInput(opened),
            getPositionAtPoint: (x, y) => handleRef.current === opened ? hostPointRef.current(x, y) : null,
            save: () => {
              if (handleRef.current !== opened) throw new Error('Presentation is no longer open');
              if (coordinator.busy() || pointerGestureRef.current || resizeRef.current) {
                throw new Error('Await flushPendingInput before saving pending input');
              }
              return opened.save();
            },
            selectText,
            focus: () => stageRef.current?.focus(),
          });
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
    clearSelection,
    commandStore,
    coordinator,
    file,
    flushPendingInput,
    goToSlide,
    stableFonts,
    refresh,
    refreshAt,
    reportError,
    selectText,
    setActiveTool,
    setSelection,
    setShapeSelection,
  ]);

  useEffect(() => {
    if (!collaborationOnReplica || !collaborationReplica) return;
    collaborationOnReplica(collaborationReplica);
    return () => collaborationOnReplica(null);
  }, [collaborationOnReplica, collaborationReplica]);

  // hover goes untracked under a non-select tool, so drop it rather than let a
  // stale target paint the cursor when the tool switches back
  useEffect(() => setHoverTarget(null), [activeTool]);

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

  const editedShapeBounds = useMemo<FrameBounds | null>(() => {
    if (!model?.frame || !selection || selectedShapeBounds) return null;
    const slide = model.snapshot.slides[model.slideIndex];
    const shape = slide ? findShape(slide.shapes, selection.shapeId) : null;
    return shape ? frameBoundsForShape(model.snapshot, model.frame, shape) : null;
  }, [model, selectedShapeBounds, selection]);

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

  const selectSlide = (index: number) => {
    goToSlide(index + 1);
  };

  const createTextBox = (start: SlidePoint, end: SlidePoint) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current?.frame || readOnlyRef.current) return;
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
        style: textStyleRef.current,
      });
      const next = refreshAt(undefined, true);
      setActiveTool('select');
      setShapeSelection(null);
      setDragPreview(null);
      setTextBoxPreview(null);
      pointerGestureRef.current = null;
      recentClickRef.current = null;
      const shape = next?.snapshot.slides[next.slideIndex]?.shapes.find(
        (candidate) => candidate.id === receipt.shapeId
      );
      const story = shape?.textStories[0];
      if (story) {
        setSelection({
          shapeId: shape.id,
          storyId: story.id,
          anchor: 0,
          focus: 0,
          focusLine: 0,
        });
        stageRef.current?.focus();
      }
    } catch (value) {
      reportError(value);
      throw value;
    }
  };

  const createShape = (
    geometry: PptxShapePreset,
    start: SlidePoint,
    end: SlidePoint
  ) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current?.frame || readOnlyRef.current) return;
    const slide = current.snapshot.slides[current.slideIndex];
    const preset = SHAPE_PRESETS.find((candidate) => candidate.geometry === geometry);
    if (!slide || !preset) return;
    const dragged = Math.abs(end.x - start.x) >= 6 || Math.abs(end.y - start.y) >= 6;
    const width = dragged ? Math.abs(end.x - start.x) : current.frame.width * 0.22;
    const height = dragged ? Math.abs(end.y - start.y) : current.frame.height * 0.18;
    const x = Math.max(
      0,
      Math.min(
        dragged ? Math.min(start.x, end.x) : start.x - width / 2,
        current.frame.width - width
      )
    );
    const y = Math.max(
      0,
      Math.min(
        dragged ? Math.min(start.y, end.y) : start.y - height / 2,
        current.frame.height - height
      )
    );
    try {
      const receipt = handle.addShape(slide.id, {
        name: t(preset.labelKey),
        geometry,
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
        fill: '#d9eaf7',
      });
      const next = refreshAt(undefined, true);
      setActiveTool('select');
      setSelection(null);
      setShapeSelection({ slideId: slide.id, shapeId: receipt.shapeId });
      setDragPreview(null);
      setTextBoxPreview(null);
      pointerGestureRef.current = null;
      recentClickRef.current = null;
      if (next) stageRef.current?.focus();
    } catch (value) {
      reportError(value);
      throw value;
    }
  };

  /** Inserts a chosen picture on the slide it was chosen for, once decoded, if the gate still passes. */
  const insertPicture = async (file: File, origin: PictureOrigin) => {
    const { handle } = origin;
    const allowed = () =>
      handleRef.current === handle &&
      generationRef.current === origin.generation &&
      liveGate('insertImage').enabled;
    try {
      if (file.size > MAX_INSERT_IMAGE_BYTES) {
        throw new Error(
          `image is ${file.size} bytes, exceeds the ${MAX_INSERT_IMAGE_BYTES}-byte limit`
        );
      }
      const extension = file.name.split('.').pop()?.toLowerCase() ?? '';
      const contentType = file.type === 'image/jpg' ? 'image/jpeg' : file.type || INSERT_IMAGE_TYPES[extension];
      if (!Object.values(INSERT_IMAGE_TYPES).includes(contentType)) {
        throw new Error(`unsupported image type ${file.type || extension}`);
      }
      const dataUrl = await readFileAsDataUrl(file);
      const mediaBase64 = dataUrl.slice(dataUrl.indexOf(',') + 1);
      let previewUrl = `data:${contentType};base64,${mediaBase64}`;
      if (contentType === 'image/tiff') {
        const bytes = Uint8Array.from(atob(mediaBase64), (char) => char.charCodeAt(0));
        previewUrl = await readFileAsDataUrl(presentationImageBlob(bytes));
      }
      const natural = await loadImageSize(previewUrl);
      const current = modelRef.current;
      if (!allowed() || !current?.frame) return;
      if (natural.width <= 0 || natural.height <= 0) throw new Error('image dimensions must be positive');
      const slideIndex = current.snapshot.slides.findIndex((slide) => slide.id === origin.slideId);
      const slide = current.snapshot.slides[slideIndex];
      if (!slide) return;
      const maxWidth = current.frame.width * 0.5;
      const maxHeight = current.frame.height * 0.5;
      const scale = Math.min(maxWidth / natural.width, maxHeight / natural.height, 1);
      const width = Math.max(1, natural.width * scale);
      const height = Math.max(1, natural.height * scale);
      const x = (current.frame.width - width) / 2;
      const y = (current.frame.height - height) / 2;
      const receipt = handle.addPicture(slide.id, {
        name: file.name || t('objects.defaultPictureName'),
        rect: {
          x: Math.round((x * current.snapshot.widthEmu) / current.frame.width),
          y: Math.round((y * current.snapshot.heightEmu) / current.frame.height),
          width: Math.round((width * current.snapshot.widthEmu) / current.frame.width),
          height: Math.round((height * current.snapshot.heightEmu) / current.frame.height),
        },
        contentType,
        mediaBase64,
      });
      const next = refreshAt(undefined, true);
      setActiveTool('select');
      setDragPreview(null);
      setTextBoxPreview(null);
      pointerGestureRef.current = null;
      recentClickRef.current = null;
      if (next && slideIndex === next.slideIndex) {
        setSelection(null);
        setShapeSelection({ slideId: slide.id, shapeId: receipt.shapeId });
        const focused = document.activeElement;
        if (!focused || focused === document.body || stageRef.current?.contains(focused)) {
          stageRef.current?.focus();
        }
      }
    } catch (value) {
      if (allowed()) reportError(value);
      throw value;
    }
  };

  /** The presentation and slide a picture chosen now is for. */
  const pictureOrigin = (): PictureOrigin | null => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle) return null;
    return {
      handle,
      generation: generationRef.current,
      slideId: current?.snapshot.slides[current.slideIndex]?.id ?? null,
    };
  };

  /**
   * Accepts a chosen picture as input, ordered with the edits and commands
   * around it, for the presentation and slide the picker opened on.
   */
  const admitPicture = (file: File) => {
    const origin = pickerOriginRef.current ?? pictureOrigin();
    pickerOriginRef.current = null;
    if (
      !origin ||
      origin.handle !== handleRef.current ||
      origin.generation !== generationRef.current
    ) {
      return;
    }
    void coordinator.input(() => insertPicture(file, origin)).catch(() => {});
  };

  const pointerDown = (event: PointerEvent<HTMLCanvasElement>) => {
    if (!event.isPrimary || event.button !== 0) return;
    typingRef.current = null;
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
    caretGoalRef.current = null;
    resizeRef.current = null;
    setResizeDelta(null);
    const slide = current.snapshot.slides[current.slideIndex];
    const shapePreset = shapePresetFromTool(activeTool);
    if (!readOnly && (activeTool === 'textBox' || shapePreset) && slide) {
      setSelection(null);
      setShapeSelection(null);
      setDragPreview(null);
      setTextBoxPreview({ start: point, end: point });
      recentClickRef.current = null;
      pointerGestureRef.current = shapePreset
        ? {
            kind: 'shapeInsert',
            pointerId: event.pointerId,
            slideId: slide.id,
            geometry: shapePreset,
            start: point,
            last: point,
          }
        : {
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
      const engineHit = handle.hitTest(point.x, point.y);
      // The edge band grabs the box rather than typing in it, so a click there
      // reads as a shape hit and matches the cursor the pointer showed.
      const hit =
        engineHit?.kind === 'text' &&
        pointerTargetAtPoint(current.frame, point, scale) === 'shape'
          ? { kind: 'shape' as const, shapeId: engineHit.shapeId }
          : engineHit;
      const shape = slide && hit ? findTopLevelShape(slide, hit.shapeId) : null;
      const recentClick = recentClickRef.current;
      const repeatedClick = Boolean(
        slide &&
          shape &&
          recentClick?.slideId === slide.id &&
          recentClick.shapeId === shape.id &&
          event.timeStamp - recentClick.timeStamp <= 600 &&
          Math.hypot(
            event.clientX - recentClick.clientX,
            event.clientY - recentClick.clientY
          ) <= 6
      );
      const clickCount =
        repeatedClick && recentClick ? Math.min(recentClick.count + 1, 3) : 1;
      const hitLocation =
        hit?.kind === 'text'
          ? textLocationAtPoint(current.frame, hit.shapeId, hit.storyId, point)
          : null;
      recentClickRef.current = null;
      if (
        slide &&
        hit?.kind === 'text' &&
        shape &&
        clickCount >= 2
      ) {
        const story = handle.story(hit.storyId);
        const granularity = clickCount >= 3 ? 'paragraph' : 'word';
        const range = textRangeAt(storyText(story), hit.position, granularity);
        setSelection({
          shapeId: hit.shapeId,
          storyId: hit.storyId,
          anchor: range.start,
          focus: range.end,
          focusLine: hitLocation?.lineIndex,
        });
        setShapeSelection(null);
        setDragPreview(null);
        pointerGestureRef.current = {
          kind: 'text',
          pointerId: event.pointerId,
          slideId: slide.id,
          shapeId: hit.shapeId,
          storyId: hit.storyId,
          anchor: range.start,
          focus: range.end,
          focusLine: hitLocation?.lineIndex,
          granularity,
          clickCount,
          clickShapeId: shape.id,
          startClientX: event.clientX,
          startClientY: event.clientY,
          dragging: false,
        };
        event.currentTarget.setPointerCapture(event.pointerId);
      } else if (
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
          focusLine: hitLocation?.lineIndex,
        });
        setShapeSelection(null);
        setDragPreview(null);
        pointerGestureRef.current = {
          kind: 'text',
          pointerId: event.pointerId,
          slideId: slide.id,
          shapeId: hit.shapeId,
          storyId: hit.storyId,
          anchor,
          focus: hit.position,
          focusLine: hitLocation?.lineIndex,
          granularity: 'character',
          clickCount,
          clickShapeId: shape?.id ?? hit.shapeId,
          startClientX: event.clientX,
          startClientY: event.clientY,
          dragging: false,
        };
        event.currentTarget.setPointerCapture(event.pointerId);
      } else if (slide && hit) {
        if (shape) {
          setSelection(null);
          setShapeSelection({ slideId: slide.id, shapeId: shape.id });
          setDragPreview(null);
          if (!readOnly && canMoveShape(shape)) {
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
              clickCount,
              dragging: false,
            };
            event.currentTarget.setPointerCapture(event.pointerId);
          } else {
            pointerGestureRef.current = null;
          }
        } else {
          setSelection(null);
          setShapeSelection(null);
          recentClickRef.current = null;
          pointerGestureRef.current = null;
        }
      } else {
        setSelection(null);
        setShapeSelection(null);
        setDragPreview(null);
        recentClickRef.current = null;
        pointerGestureRef.current = null;
      }
      stageRef.current?.focus();
      event.preventDefault();
    } catch (value) {
      reportError(value);
    }
  };

  const updatePointerGesture = (event: PointerEvent<HTMLCanvasElement>): boolean => {
    const gesture = pointerGestureRef.current;
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!gesture || gesture.pointerId !== event.pointerId || !handle || !current?.frame) {
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
      if (
        !gesture.dragging &&
        passedDragThreshold(
          gesture.startClientX,
          gesture.startClientY,
          event.clientX,
          event.clientY
        )
      ) {
        gesture.dragging = true;
      }
      const focus = textLocationAtPoint(
        current.frame,
        gesture.shapeId,
        gesture.storyId,
        point
      );
      if (focus) {
        const range =
          gesture.granularity === 'character'
            ? { anchor: gesture.anchor, focus: focus.position }
            : extendTextRange(
                { start: gesture.anchor, end: gesture.focus },
                textRangeAt(
                  storyText(handle.story(gesture.storyId)),
                  focus.position,
                  gesture.granularity
                )
              );
        setSelection({
          shapeId: gesture.shapeId,
          storyId: gesture.storyId,
          ...range,
          focusLine: focus.lineIndex,
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
    if (updatePointerGesture(event)) {
      event.preventDefault();
      return;
    }
    // a non-select tool paints its own cursor, and a live gesture holds the one it started with
    if (pointerGestureRef.current || activeTool !== 'select') return;
    const current = modelRef.current;
    if (!current?.frame) return;
    const point = slidePoint(
      event.currentTarget.getBoundingClientRect(),
      current.frame,
      event.clientX,
      event.clientY
    );
    setHoverTarget(point ? pointerTargetAtPoint(current.frame, point, scale) : null);
  };

  const pointerLeave = () => {
    if (!pointerGestureRef.current) setHoverTarget(null);
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
      void coordinator.input(() => createTextBox(gesture.start, gesture.last)).catch(() => {});
      event.preventDefault();
      return;
    }
    if (gesture.kind === 'shapeInsert') {
      setTextBoxPreview(null);
      void coordinator
        .input(() => createShape(gesture.geometry, gesture.start, gesture.last))
        .catch(() => {});
      event.preventDefault();
      return;
    }
    if (gesture.kind === 'text') {
      recentClickRef.current = gesture.dragging
        ? null
        : {
            slideId: gesture.slideId,
            shapeId: gesture.clickShapeId,
            clientX: event.clientX,
            clientY: event.clientY,
            timeStamp: event.timeStamp,
            count: gesture.clickCount,
          };
      return;
    }
    if (gesture.kind !== 'shape') return;
    if (!gesture.dragging) {
      recentClickRef.current = {
        slideId: gesture.slideId,
        shapeId: gesture.shapeId,
        clientX: event.clientX,
        clientY: event.clientY,
        timeStamp: event.timeStamp,
        count: gesture.clickCount,
      };
      return;
    }
    recentClickRef.current = null;
    event.preventDefault();
    void coordinator.input(() => moveShape(gesture)).catch(() => {});
  };

  const moveShape = (gesture: Extract<PointerGesture, { kind: 'shape' }>) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    const slide = current?.snapshot.slides[current.slideIndex];
    if (!handle || !current?.frame || slide?.id !== gesture.slideId || readOnlyRef.current) return;
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
    } catch (value) {
      reportError(value);
      throw value;
    }
  };

  const cancelPointerGesture = (event: PointerEvent<HTMLCanvasElement>) => {
    if (pointerGestureRef.current?.pointerId !== event.pointerId) return;
    pointerGestureRef.current = null;
    setDragPreview(null);
    setTextBoxPreview(null);
    recentClickRef.current = null;
  };

  const commit = (nextSelection: PptxTextSelection | null): PptxTextSelection | null => {
    if (readOnlyRef.current) return selectionRef.current;
    const committed = nextSelection ? { ...nextSelection, focusLine: undefined } : null;
    setSelection(committed);
    setShapeSelection(null);
    recentClickRef.current = null;
    refreshAt(undefined, true);
    return committed;
  };

  const keyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (!handleRef.current) return;
    if (event.key === 'Escape') {
      canvasReview.setEnabled(false);
      setActiveTool('select');
      setTextBoxPreview(null);
      pointerGestureRef.current = null;
      resizeRef.current = null;
      setResizeDelta(null);
      event.preventDefault();
      return;
    }
    const input: KeyInput = {
      key: event.key,
      shiftKey: event.shiftKey,
      metaKey: event.metaKey,
      ctrlKey: event.ctrlKey,
      altKey: event.altKey,
    };
    const busy = coordinator.busy();
    const typing = busy ? typingRef.current : null;
    const hasTarget = typing ? typing.selection !== null : selectionRef.current !== null;
    if (!hasTarget || reviewingRef.current || !isEditingKey(input)) return;
    if (busy) {
      event.preventDefault();
      if (!typingRef.current) {
        const selection = selectionRef.current;
        typingRef.current = { selection: selection ? anchorSelection(selection) : null };
      }
      const queued = typingRef.current;
      void coordinator.input(() => typeQueued(input, queued), 'keyboard').catch(() => {});
      return;
    }
    if (applyKey(input, selectionRef.current!).handled) event.preventDefault();
  };

  /** The selection as positions that later edits, undo and remote updates move along. */
  const anchorSelection = (selection: PptxTextSelection): AnchoredSelection | null => {
    const handle = handleRef.current;
    if (!handle) return null;
    try {
      return {
        shapeId: selection.shapeId,
        storyId: selection.storyId,
        anchor: handle.anchorCaret(selection.storyId, selection.anchor),
        focus: handle.anchorCaret(selection.storyId, selection.focus),
      };
    } catch {
      return null;
    }
  };

  /**
   * Applies a keystroke accepted while earlier work was pending, at the text it
   * was typed into, wherever that text moved meanwhile. Later keystrokes of the
   * same run follow its result.
   */
  const typeQueued = (input: KeyInput, typing: TypingTarget) => {
    const anchored = typing.selection;
    if (!anchored) return;
    const handle = handleRef.current;
    if (!handle) throw new Error('Presentation is no longer open');
    let target: PptxTextSelection;
    try {
      const anchor = handle.resolveCaretAnchor(anchored.anchor);
      const focus = handle.resolveCaretAnchor(anchored.focus);
      if (anchor === null || focus === null) {
        throw new Error('The text changed before the typed input was applied');
      }
      target = { shapeId: anchored.shapeId, storyId: anchored.storyId, anchor, focus };
    } catch (value) {
      reportError(value);
      throw value;
    }
    const selection = applyKey(input, target, true).selection;
    typing.selection = selection ? anchorSelection(selection) : null;
  };

  /**
   * Moves the caret or edits text at `selection`. Returns whether the key was
   * used and the selection after it; `strict` rethrows engine refusals after
   * reporting them.
   */
  const applyKey = (
    input: KeyInput,
    selection: PptxTextSelection,
    strict = false
  ): { handled: boolean; selection: PptxTextSelection | null } => {
    const handle = handleRef.current;
    const unhandled = { handled: false, selection };
    if (!handle || reviewingRef.current) return unhandled;
    const start = Math.min(selection.anchor, selection.focus);
    const end = Math.max(selection.anchor, selection.focus);
    const edited = (next: PptxTextSelection | null) => ({ handled: true, selection: commit(next) });
    try {
      const moved = caretDestination(input, selection);
      if (moved !== null) {
        const next = {
          ...selection,
          anchor: input.shiftKey ? selection.anchor : moved.position,
          focus: moved.position,
          focusLine: moved.lineIndex,
        };
        setSelection(next);
        return { handled: true, selection: next };
      }
      if (readOnlyRef.current) return unhandled;
      if (input.key === 'Backspace') {
        if (start !== end) {
          handle.deleteText(selection.storyId, start, end);
          return edited({ ...selection, anchor: start, focus: start });
        }
        const previous = previousTextIndex(handle.story(selection.storyId), start);
        if (previous < start) {
          handle.deleteText(selection.storyId, previous, start);
          return edited({ ...selection, anchor: previous, focus: previous });
        }
        return { handled: true, selection };
      }
      if (input.key === 'Delete') {
        if (start !== end) {
          handle.deleteText(selection.storyId, start, end);
          return edited({ ...selection, anchor: start, focus: start });
        }
        const next = nextTextIndex(handle.story(selection.storyId), end);
        if (next > end) {
          handle.deleteText(selection.storyId, end, next);
          return edited({ ...selection, anchor: end, focus: end });
        }
        return { handled: true, selection };
      }
      if (input.key === 'Enter') {
        if (start !== end) handle.deleteText(selection.storyId, start, end);
        handle.insertParagraphBreak(selection.storyId, start);
        return edited({ ...selection, anchor: start + 1, focus: start + 1 });
      }
      if (isCharacter(input)) {
        if (start !== end) handle.deleteText(selection.storyId, start, end);
        handle.insertText(selection.storyId, start, input.key, textStyleRef.current);
        const next = start + input.key.length;
        return edited({ ...selection, anchor: next, focus: next });
      }
      return unhandled;
    } catch (value) {
      reportError(value);
      if (strict) throw value;
      return { handled: true, selection };
    }
  };

  const selectedShapeNow = () => {
    const current = modelRef.current;
    const target = shapeSelectionRef.current;
    const slide = current?.snapshot.slides[current.slideIndex];
    return slide && target && slide.id === target.slideId ? findShape(slide.shapes, target.shapeId) : null;
  };

  /** Formats the selected range, the whole story of a selected shape, or the caret's stored style. */
  const applyFormatting = (patch: TextStylePatch): boolean => {
    const handle = handleRef.current;
    const selection = selectionRef.current;
    const shapeStoryId = selection ? null : (selectedShapeNow()?.textStories[0]?.id ?? null);
    if (!handle || (!selection && !shapeStoryId)) return false;
    setTextStyle((current) => ({ ...current, ...patch }));
    if (selection) {
      if (selection.anchor === selection.focus) return true;
      handle.formatText(
        selection.storyId,
        Math.min(selection.anchor, selection.focus),
        Math.max(selection.anchor, selection.focus),
        patch
      );
    } else {
      formatStory(handle, handle.story(shapeStoryId!), patch);
    }
    refreshAt(undefined, true);
    return true;
  };

  /** Where a caret-movement key lands, or `null` when the key moves nothing.
   *  Vertical steps hold the column they started from; the goal is tied to the
   *  position it produced, so a click or a keystroke in between drops it. */
  const caretDestination = (
    input: KeyInput,
    selection: PptxTextSelection
  ): { position: number; lineIndex?: number } | null => {
    const handle = handleRef.current;
    if (!handle) return null;
    const story = handle.story(selection.storyId);
    const last = Math.max(0, story.length - 1);
    const clamp = (position: number) => Math.max(0, Math.min(last, position));
    const lines = caretLinesFor(modelRef.current?.frame, selection);
    const current = { position: selection.focus, lineIndex: selection.focusLine };
    // macOS reads Option as word-wise and Command as line-edge; Control is the
    // Windows word-wise modifier.
    const wordWise = input.altKey || (input.ctrlKey && !input.metaKey);
    const toEdge = input.metaKey;

    if (input.key === 'ArrowUp' || input.key === 'ArrowDown') {
      const goalKey = { shapeId: selection.shapeId, ...current };
      const goal =
        caretGoalRef.current && sameCaretGoalKey(caretGoalRef.current, goalKey)
          ? caretGoalRef.current.goalX
          : caretGoalX(lines, current);
      const direction = input.key === 'ArrowUp' ? 'up' : 'down';
      const moved = verticalCaretMove(lines, current, direction, goal);
      const destination = { ...moved, position: clamp(moved.position) };
      caretGoalRef.current =
        goal === undefined
          ? null
          : {
              shapeId: selection.shapeId,
              goalX: goal,
              position: destination.position,
              lineIndex: destination.lineIndex,
            };
      return destination;
    }
    caretGoalRef.current = null;
    if (input.key === 'ArrowLeft' || input.key === 'ArrowRight') {
      const direction = input.key === 'ArrowLeft' ? -1 : 1;
      if (toEdge) {
        const destination = lineEdge(lines, current, direction < 0 ? 'start' : 'end');
        return { ...destination, position: clamp(destination.position) };
      }
      const position = wordWise
        ? wordBoundary(storyText(story), selection.focus, direction)
        : selection.focus + direction;
      return { position: clamp(position), lineIndex: selection.focusLine };
    }
    if (input.key === 'Home') {
      if (toEdge || input.ctrlKey) return { position: 0, lineIndex: 0 };
      const destination = lineEdge(lines, current, 'start');
      return { ...destination, position: clamp(destination.position) };
    }
    if (input.key === 'End') {
      if (toEdge || input.ctrlKey) {
        return { position: last, lineIndex: Math.max(0, lines.length - 1) };
      }
      const destination = lineEdge(lines, current, 'end');
      return { ...destination, position: clamp(destination.position) };
    }
    return null;
  };

  const applyAlignment = (alignment: ParagraphAlignment): boolean => {
    const handle = handleRef.current;
    const selection = selectionRef.current;
    const storyId = selection?.storyId ?? selectedShapeNow()?.textStories[0]?.id;
    if (!handle || !storyId) return false;
    if (selection) {
      handle.setParagraphAlignment(
        storyId,
        Math.min(selection.anchor, selection.focus),
        Math.max(selection.anchor, selection.focus),
        alignment
      );
    } else {
      handle.setParagraphAlignment(storyId, 0, handle.story(storyId).length, alignment);
    }
    refreshAt(undefined, true);
    return true;
  };

  const formatShape = (action: ShapeFormattingAction): boolean => {
    const handle = handleRef.current;
    const target = shapeSelectionRef.current;
    const shape = selectedShapeNow();
    if (!handle || !target || !shape) return false;
    if (action.type === 'fillColor') {
      handle.setShapeFill(target.slideId, target.shapeId, action.value);
    } else if (action.type === 'strokeColor') {
      handle.setShapeStroke(target.slideId, target.shapeId, action.value ? { color: action.value } : {});
    } else if (action.type === 'strokeWidth') {
      handle.setShapeStroke(
        target.slideId,
        target.shapeId,
        action.value === null ? {} : { widthPt: action.value }
      );
    } else if (action.type === 'adjust') {
      handle.setShapeAdjust(target.slideId, target.shapeId, {
        ...shape.adjustValues,
        [action.name]: action.value,
      });
    } else {
      const zOrder = {
        front: handle.bringShapeToFront,
        back: handle.sendShapeToBack,
        forward: handle.bringShapeForward,
        backward: handle.sendShapeBackward,
      }[action.value];
      zOrder(target.slideId, target.shapeId);
    }
    refreshAt(undefined, true);
    return true;
  };

  const addSlide = (layoutPartPath?: string | null): boolean => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current) return false;
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
    recentClickRef.current = null;
    resizeRef.current = null;
    setResizeDelta(null);
    refreshAt(index, true, true);
    return true;
  };

  const history = (direction: 'undo' | 'redo'): boolean => {
    const handle = handleRef.current;
    if (!handle) return false;
    if (direction === 'undo') handle.undo();
    else handle.redo();
    setSelection(null);
    setDragPreview(null);
    setTextBoxPreview(null);
    setActiveTool('select');
    pointerGestureRef.current = null;
    recentClickRef.current = null;
    resizeRef.current = null;
    setResizeDelta(null);
    refreshAt(undefined, true);
    return true;
  };

  const changeTool = (tool: PptxEditorTool): boolean => {
    const changed = activeToolRef.current !== tool;
    canvasReview.setEnabled(false);
    setActiveTool(tool);
    if (tool !== 'select') {
      setSelection(null);
      setShapeSelection(null);
    }
    setTextBoxPreview(null);
    pointerGestureRef.current = null;
    resizeRef.current = null;
    setResizeDelta(null);
    return changed;
  };

  /**
   * Runs the host's save request outside the input queue, which it may flush,
   * then serializes after everything accepted by then.
   */
  const save = (): Promise<PptxSaveOutcome> => {
    if (pendingSaveRef.current) return pendingSaveRef.current;
    const handle = handleRef.current;
    if (!handle) return Promise.resolve('replaced');
    const generation = generationRef.current;
    const current = () => handleRef.current === handle && generationRef.current === generation;
    const pending = Promise.resolve()
      .then(async (): Promise<PptxSaveOutcome> => {
        if (onSaveRequest && (await onSaveRequest()) !== true) return 'requested';
        if (!current()) return 'replaced';
        return coordinator.run((): PptxSaveOutcome => {
          if (!current()) return 'replaced';
          if (pointerGestureRef.current || resizeRef.current) return 'gesture';
          const bytes = handle.save();
          if (onSave) onSave(bytes);
          else downloadBytes(bytes, fileName ?? 'presentation.pptx', PPTX_MIME);
          return 'saved';
        });
      })
      .catch((value: unknown): PptxSaveOutcome => {
        if (value instanceof PptxInputFailure) return 'input-failed';
        if (value instanceof PptxInputStale || !current()) return 'replaced';
        reportError(value);
        return 'failed';
      })
      .finally(() => {
        if (pendingSaveRef.current === pending) pendingSaveRef.current = null;
      });
    pendingSaveRef.current = pending;
    return pending;
  };

  hostPointRef.current = (clientX, clientY) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    const canvas = canvasRef.current;
    if (!handle || !current?.frame || !canvas || canvasReview.reviewing) return null;
    const point = slidePoint(canvas.getBoundingClientRect(), current.frame, clientX, clientY);
    if (!point || point.x < 0 || point.y < 0 ||
        point.x >= current.frame.width || point.y >= current.frame.height) return null;
    handle.layoutSlide(current.slideIndex);
    const hit = handle.hitTest(point.x, point.y);
    const slide = current.snapshot.slides[current.slideIndex];
    return hit && slide ? { ...hit, slide: current.slideIndex + 1, slideId: slide.id } : null;
  };

  const slidePointFromClient = (clientX: number, clientY: number): SlidePoint | null => {
    const canvas = canvasRef.current;
    const frame = model?.frame;
    if (!canvas || !frame) return null;
    return slidePoint(canvas.getBoundingClientRect(), frame, clientX, clientY);
  };

  const resizePointerDown =
    (handle: ResizeHandle) => (event: PointerEvent<HTMLSpanElement>) => {
      if (readOnly || !shapeSelection || !selectedShape || !canResizeShape(selectedShape)) return;
      const point = slidePointFromClient(event.clientX, event.clientY);
      if (!point) return;
      event.preventDefault();
      event.stopPropagation();
      resizeRef.current = {
        pointerId: event.pointerId,
        handle,
        start: point,
        slideId: shapeSelection.slideId,
        shapeId: shapeSelection.shapeId,
        delta: { x: 0, y: 0 },
      };
      setResizeDelta({ x: 0, y: 0 });
      event.currentTarget.setPointerCapture(event.pointerId);
    };

  /** The gesture's geometry lives on the ref, and the state only mirrors it for
   *  the preview: the release commits the pointer's own position, so a move
   *  whose render has not landed yet cannot resize the shape to a stale size. */
  const resizeDeltaFrom = (event: PointerEvent<HTMLSpanElement>): SlidePoint | null => {
    const gesture = resizeRef.current;
    if (!gestureOwnsPointer(gesture, event.pointerId)) return null;
    return resizeCommitDelta(
      gesture.start,
      gesture.delta,
      slidePointFromClient(event.clientX, event.clientY)
    );
  };

  const resizePointerMove = (event: PointerEvent<HTMLSpanElement>) => {
    const gesture = resizeRef.current;
    const delta = resizeDeltaFrom(event);
    if (!gestureOwnsPointer(gesture, event.pointerId) || !delta) return;
    gesture.delta = delta;
    setResizeDelta(delta);
  };

  const endResizeGesture = (event: PointerEvent<HTMLSpanElement>) => {
    const gesture = resizeRef.current;
    if (!gestureOwnsPointer(gesture, event.pointerId)) return null;
    resizeRef.current = null;
    setResizeDelta(null);
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
    return gesture;
  };

  /** A cancelled pointer abandons the drag; the shape keeps the size it had. */
  const resizePointerCancel = (event: PointerEvent<HTMLSpanElement>) => {
    endResizeGesture(event);
  };

  const resizePointerUp = (event: PointerEvent<HTMLSpanElement>) => {
    const delta = resizeDeltaFrom(event);
    const gesture = endResizeGesture(event);
    if (!gesture || !delta || (delta.x === 0 && delta.y === 0)) return;
    void coordinator.input(() => resizeShape(gesture, delta)).catch(() => {});
  };

  const resizeShape = (
    gesture: NonNullable<typeof resizeRef.current>,
    delta: SlidePoint
  ) => {
    const api = handleRef.current;
    const current = modelRef.current;
    const slide = current?.snapshot.slides[current.slideIndex];
    const shape = slide ? findShape(slide.shapes, gesture.shapeId) : null;
    if (readOnlyRef.current || !api || !current?.frame || slide?.id !== gesture.slideId || !shape) {
      return;
    }
    try {
      const box = resizedShapeBox(
        current.snapshot,
        current.frame,
        shape,
        gesture.handle,
        delta
      );
      if (
        box &&
        (box.x !== shape.x ||
          box.y !== shape.y ||
          box.width !== shape.width ||
          box.height !== shape.height)
      ) {
        api.setShapeRect(gesture.slideId, gesture.shapeId, box);
        refreshAt(undefined, true);
      }
    } catch (value) {
      reportError(value);
      throw value;
    }
  };

  const slideCount = model?.snapshot.slides.length ?? 0;
  const currentSlide = model?.slideIndex ?? 0;

  const startPresenting = () => {
    if (!modelRef.current?.snapshot.slides.length) return;
    setPresenting(true);
  };

  // Export the current slide through the same canvas painter the editor draws
  // with, so the png matches what is on screen. Only capturing the slide waits
  // for accepted input; painting runs outside the input queue.
  const exportPng = async (): Promise<PptxExportOutcome> => {
    let captured: { frame: SlideDisplayList; slideIndex: number; handle: PresentationHandle };
    try {
      captured = await coordinator.run(() => {
        const current = modelRef.current;
        const handle = handleRef.current;
        if (!current?.frame || !handle) throw new PptxInputStale();
        return { frame: current.frame, slideIndex: current.slideIndex, handle };
      });
    } catch (value) {
      return value instanceof PptxInputFailure ? 'input-failed' : 'replaced';
    }
    const { frame, slideIndex, handle } = captured;
    // Painting is async, so the deck can be replaced or disposed mid-export.
    // Pin the handle the frame came from, behind its own decode cache, rather
    // than reading whichever deck `handleRef` holds by the time an asset
    // resolves.
    const pinned = { current: handle };
    const cache = { current: new Map<string, Promise<CanvasImageSource | null>>() };
    try {
      const blob = await slideToPng(frame, {
        scale: window.devicePixelRatio || 1,
        resolveImage: (assetId) => resolveImage(assetId, pinned, cache, decodeImageError),
      });
      const bytes = new Uint8Array(await blob.arrayBuffer());
      if (handleRef.current !== handle) return 'replaced';
      downloadBytes(bytes, pngName(fileName, slideIndex), 'image/png');
      return 'exported';
    } catch (value) {
      if (handleRef.current !== handle) return 'replaced';
      reportError(value);
      return 'failed';
    }
  };
  const shapeDragDelta =
    dragPreview && dragPreview.shapeId === shapeSelection?.shapeId ? dragPreview.delta : null;
  const resizingHandle = resizeRef.current?.handle;
  const resizePreview =
    model?.frame && selectedShape && resizeDelta && resizingHandle
      ? resizedShapeBounds(
          model.snapshot,
          model.frame,
          selectedShape,
          resizingHandle,
          resizeDelta
        )
      : null;
  const selectionBox = selectedShapeBounds
    ? resizeDelta && resizingHandle
      ? resizePreview ?? selectedShapeBounds
      : {
          ...selectedShapeBounds,
          x: selectedShapeBounds.x + (shapeDragDelta?.x ?? 0),
          y: selectedShapeBounds.y + (shapeDragDelta?.y ?? 0),
        }
    : null;

  const status: 'ready' | 'loading' | 'empty' = !file
    ? 'empty'
    : model && handleRef.current
      ? 'ready'
      : loading || !error
        ? 'loading'
        : 'empty';
  const proposalsAvailable = Boolean(handleRef.current?.isProposalsAvailable());
  const liveGate = usePptxCommandBinding(commands, {
    handle: handleRef.current,
    handleRef,
    modelRef,
    selectionRef,
    shapeSelectionRef,
    textStyleRef,
    activeToolRef,
    generationRef,
    coordinator,
    status,
    readOnlyRef,
    zoomRef,
    proposalsRef,
    proposalsOpenRef,
    reviewEnabledRef: canvasReview.enabledRef,
    reviewSelectedIdRef: canvasReview.selectedIdRef,
    proposalsAvailable,
    gestureActive: () => Boolean(pointerGestureRef.current || resizeRef.current),
    baseStyle: initialStyle,
    i18n,
    t,
    actions: {
      format: applyFormatting,
      align: applyAlignment,
      formatShape,
      insertSlide: addSlide,
      history,
      setTool: changeTool,
      setZoom: (next) => {
        const changed = next !== zoomRef.current;
        setZoom(next);
        return changed;
      },
      openPicker: () => {
        pickerOriginRef.current = pictureOrigin();
        pictureInputRef.current?.click();
      },
      save,
      exportPng,
      present: startPresenting,
      setProposalsOpen: (open) => {
        if (open) refreshProposals();
        setProposalsOpen(open);
      },
      selectProposal: (proposalId) => {
        canvasReview.select(proposalId);
        canvasReview.setEnabled(true);
      },
      showProposalDiff: (enabled) => canvasReview.setEnabled(enabled),
      acceptProposal,
      rejectProposal,
      reportError,
      focusEditor: () => stageRef.current?.focus(),
    },
    stamp: [
      model,
      selection,
      shapeSelection,
      textStyle,
      activeTool,
      zoom,
      readOnly,
      status,
      canvasReview.enabled,
      canvasReview.selected,
      canvasReview.available,
      proposals,
      proposalsOpen,
      proposalsAvailable,
      t,
      i18n,
    ],
  });
  const presentingRef = useRef(presenting);
  presentingRef.current = presenting;
  useCommandShortcuts({ commands, containerRef: rootRef, suspended: () => presentingRef.current });

  const toolbarRegion =
    !showToolbar || toolbar === null ? null : (
      <div style={styles.toolbarShell}>
        <EditorChromeContext.Provider value>
          {toolbar !== undefined ? (
            <div style={styles.toolbar}>{toolbar}</div>
          ) : !readOnly ? (
            <>
              <EditorToolbar mode="commands" style={styles.toolbar} />
              {proposalsAvailable && (
                <ProposalsButton buttonRef={proposalButtonRef} count={proposals.length} />
              )}
            </>
          ) : null}
        </EditorChromeContext.Provider>
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
        {toolbar === undefined && <PresentButton />}
      </div>
    );

  const editor = (
    <div ref={rootRef} className={className} style={styles.root}>
      {toolbarRegion}
      <input
        ref={pictureInputRef}
        type="file"
        accept={Object.values(INSERT_IMAGE_TYPES).join(',')}
        disabled={readOnly || canvasReview.reviewing}
        tabIndex={-1}
        aria-hidden="true"
        data-testid="pptx-insert-image-input"
        style={styles.hiddenFileInput}
        onChange={(event) => {
          const chosen = event.currentTarget.files?.[0];
          event.currentTarget.value = '';
          if (chosen) admitPicture(chosen);
        }}
      />
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
          onFocus={() => setStageFocused(true)}
          onBlur={() => setStageFocused(false)}
        >
          {!readOnly && (
            <ProposalCanvasToolbar review={canvasReview} ready={paintedReview === canvasReview.diff} />
          )}
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
                  aria-hidden={canvasReview.reviewing}
                  style={{
                    ...styles.canvas,
                    cursor:
                      activeTool !== 'select'
                        ? 'crosshair'
                        : hoverTarget === 'text'
                          ? 'text'
                          : hoverTarget === 'shape'
                            ? 'move'
                            : 'default',
                  }}
                  onPointerDown={canvasReview.reviewing ? undefined : pointerDown}
                  onPointerMove={pointerMove}
                  onPointerUp={pointerUp}
                  onPointerLeave={pointerLeave}
                  onPointerCancel={cancelPointerGesture}
                  onLostPointerCapture={cancelPointerGesture}
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
                {!canvasReview.reviewing && <>
                <SelectionOverlay
                  frame={model.frame}
                  selection={selection}
                  scale={scale}
                  focused={stageFocused}
                />
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
                {selectionBox ? (
                  <>
                    <span
                      style={{
                        ...styles.shapeSelection,
                        left: selectionBox.x * scale,
                        top: selectionBox.y * scale,
                        width: Math.max(1, selectionBox.width * scale),
                        height: Math.max(1, selectionBox.height * scale),
                      }}
                      aria-hidden="true"
                    />
                    {!readOnly && selectedShape && canResizeShape(selectedShape)
                      ? RESIZE_HANDLES.map((handle) => {
                          const anchor = handleAnchor(selectionBox, handle);
                          return (
                            <span
                              key={handle}
                              data-testid={`pptx-resize-${handle}`}
                              onPointerDown={resizePointerDown(handle)}
                              onPointerMove={resizePointerMove}
                              onPointerUp={resizePointerUp}
                              onPointerCancel={resizePointerCancel}
                              onLostPointerCapture={resizePointerCancel}
                              style={{
                                ...styles.resizeHandle,
                                left: anchor.x * scale - HANDLE_SIZE / 2,
                                top: anchor.y * scale - HANDLE_SIZE / 2,
                                cursor: resizeCursor(handle),
                              }}
                            />
                          );
                        })
                      : null}
                  </>
                ) : null}
                {editedShapeBounds ? (
                  <span
                    style={{
                      ...styles.textEditOutline,
                      left: editedShapeBounds.x * scale,
                      top: editedShapeBounds.y * scale,
                      width: Math.max(1, editedShapeBounds.width * scale),
                      height: Math.max(1, editedShapeBounds.height * scale),
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
                </>}
                {canvasReview.diff && <ProposalCanvasOverlay
                  key={`${canvasReview.diff.proposal.id}:${model.slideIndex}`}
                  diff={canvasReview.diff} current={model.snapshot} frame={model.frame}
                  slideIndex={model.slideIndex} scale={scale} resolveImage={resolveProposalImage}
                  onPainted={setPaintedReview}
                  onTarget={(slideId, shapeId) => {
                    navigateProposalTarget(slideId, shapeId, canvasReview.selected?.id);
                    setProposalsOpen(true);
                  }} />}
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
        {!readOnly && proposalsOpen && model && handleRef.current && (
          <ProposalsPanel handle={handleRef.current} proposals={proposals} snapshot={model.snapshot}
            resolveImage={(assetId) => resolveImage(assetId, handleRef, imageCacheRef, decodeImageError)}
            onNavigate={navigateProposalTarget}
            onClose={() => proposalButtonRef.current?.focus()} />
        )}
      </div>
      {activeSlide && canvasReview.diff?.proposal.changes.some((change) => change.slideId === activeSlide.id && !change.shapeId && change.oldText !== change.newText) ? (
        <ProposalNotesDiff diff={canvasReview.diff} slideId={activeSlide.id} />
      ) : activeSlide ? (
        <NotesPanel
          key={activeSlide.id}
          value={activeSlide.notes ?? ''}
          disabled={!model || canvasReview.reviewing || readOnly}
          label={t('notes.panelLabel')}
          placeholder={t('notes.placeholder')}
          onCommit={(text) => {
            const handle = handleRef.current;
            if (!handle || readOnly) return;
            try {
              handle.setSlideNotes(activeSlide.id, text);
              refreshAt(undefined, true);
            } catch (value) {
              reportError(value);
            }
          }}
        />
      ) : null}
      {presenting && slideCount > 0 && handleRef.current ? (
        <PresentationOverlay
          handle={handleRef.current}
          slideCount={slideCount}
          startIndex={currentSlide}
          resolveImage={(assetId) =>
            resolveImage(assetId, handleRef, imageCacheRef, decodeImageError)
          }
          counterLabel={(current, total) => t('presentation.slideCounter', { current, total })}
          label={t('presentation.label')}
          exitLabel={t('presentation.exit')}
          previousLabel={t('presentation.previousSlide')}
          nextLabel={t('presentation.nextSlide')}
          onExit={() => setPresenting(false)}
          onError={reportError}
        />
      ) : null}
    </div>
  );
  return <PptxCommandContext.Provider value={commandStore}>{editor}</PptxCommandContext.Provider>;
}

function ProposalsButton({
  buttonRef,
  count,
}: {
  buttonRef: React.RefObject<HTMLButtonElement | null>;
  count: number;
}) {
  const store = usePptxCommands();
  const state = usePptxCommandState('proposalsPanel');
  const { t } = useTranslation();
  const described = useDisabledDescription(!state.enabled, reasonOf(state));
  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        data-testid="pptx-proposals-button"
        {...described.props}
        aria-expanded={state.value ?? false}
        title={described.reason}
        style={styles.presentButton}
        onClick={() => {
          if (state.enabled) void store.execute('proposalsPanel', null);
        }}
      >
        {t('proposals.title')} <span data-testid="pptx-proposals-count">{count}</span>
      </button>
      {described.node}
    </>
  );
}

function PresentButton() {
  const store = usePptxCommands();
  const state = usePptxCommandState('slideshow');
  const { t } = useTranslation();
  const described = useDisabledDescription(!state.enabled, reasonOf(state));
  return (
    <>
      <button
        type="button"
        onClick={() => {
          if (state.enabled) void store.execute('slideshow', null);
        }}
        {...described.props}
        title={described.reason}
        data-testid="pptx-present"
        style={{ ...styles.presentButton, opacity: state.enabled ? 1 : 0.6 }}
      >
        <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
          <path d="M6 4.5v15l13-7.5Z" />
        </svg>
        {t('toolbar.present')}
      </button>
      {described.node}
    </>
  );
}

function reasonOf(state: { enabled: boolean; disabledReason?: { message: string } }): string | undefined {
  return state.enabled ? undefined : state.disabledReason?.message;
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

function NotesPanel({
  value,
  disabled,
  label,
  placeholder,
  onCommit,
}: {
  value: string;
  disabled: boolean;
  label: string;
  placeholder: string;
  onCommit: (text: string) => void;
}) {
  return (
    <div style={styles.notesPanel}>
      <span style={styles.notesLabel}>{label}</span>
      <textarea
        aria-label={label}
        style={styles.notesTextarea}
        value={value}
        disabled={disabled}
        placeholder={placeholder}
        data-testid="pptx-notes-textarea"
        onChange={(event) => onCommit(event.target.value)}
      />
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

export function SelectionOverlay({
  frame,
  selection,
  scale,
  focused,
}: {
  frame: SlideDisplayList;
  selection: PptxTextSelection | null;
  scale: number;
  focused: boolean;
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [pageVisible, setPageVisible] = useState(
    () => typeof document === 'undefined' || document.visibilityState === 'visible'
  );
  const [caretVisible, setCaretVisible] = useState(false);

  useEffect(() => {
    if (typeof document === 'undefined') return;
    const update = () => setPageVisible(document.visibilityState === 'visible');
    document.addEventListener('visibilitychange', update);
    return () => document.removeEventListener('visibilitychange', update);
  }, []);

  useEffect(() => {
    const blinking =
      focused && pageVisible && selection !== null && selection.anchor === selection.focus;
    setCaretVisible(blinking);
    if (!blinking) return;
    const timer = window.setInterval(
      () => setCaretVisible((visible) => !visible),
      CARET_BLINK_MS
    );
    return () => window.clearInterval(timer);
  }, [focused, pageVisible, selection]);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(canvas, frame, dpr, scale);
    paintSelection(ctx, frame, selection, dpr, scale, caretVisible);
  }, [caretVisible, frame, scale, selection]);

  return <canvas ref={canvasRef} style={styles.canvasOverlay} aria-hidden="true" />;
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
  handle.formatText(story.id, 0, story.length, patch);
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
  const blob = presentationImageBlob(bytes);
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

function caretLinesFor(
  frame: SlideDisplayList | null | undefined,
  selection: PptxTextSelection
): CaretLine[] {
  const textBox = frame?.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' &&
      primitive.storyId === selection.storyId &&
      primitive.shapeId === selection.shapeId
  );
  return textBox?.lines ?? [];
}

export function paintSelection(
  ctx: CanvasRenderingContext2D,
  frame: SlideDisplayList,
  selection: PptxTextSelection | null,
  dpr: number,
  scale: number,
  caretVisible = true
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
  } else if (caretVisible) {
    const line = textBox.lines[
      caretLineIndex(textBox.lines, {
        position: start,
        lineIndex: selection.focusLine,
      })
    ];
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
    flex: '1 1 auto',
    minWidth: 0,
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
  stage: { position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, outline: 'none', overflow: 'hidden' },
  canvasHost: { display: 'flex', flex: 1, minHeight: 0, alignItems: 'center', justifyContent: 'center', width: '100%', overflow: 'auto' },
  canvasFrame: { position: 'relative', flex: '0 0 auto' },
  canvas: { display: 'block', flex: '0 0 auto', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' },
  canvasOverlay: { position: 'absolute', inset: 0, display: 'block', pointerEvents: 'none' },
  textEditOutline: {
    position: 'absolute',
    zIndex: 3,
    border: '2px dashed #6b7280',
    boxSizing: 'border-box',
    boxShadow: '0 0 0 1px rgba(255, 255, 255, 0.9)',
    pointerEvents: 'none',
  },
  resizeHandle: {
    position: 'absolute',
    zIndex: 4,
    width: HANDLE_SIZE,
    height: HANDLE_SIZE,
    borderRadius: '50%',
    border: '1px solid #2563eb',
    background: '#ffffff',
    boxShadow: '0 1px 2px rgba(15, 23, 42, 0.35)',
    boxSizing: 'border-box',
    touchAction: 'none',
  },
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
  textBoxPreview: {
    position: 'absolute',
    border: '1.5px dashed #1a73e8',
    background: 'rgba(26, 115, 232, 0.08)',
    boxSizing: 'border-box',
    pointerEvents: 'none',
  },
  notesPanel: {
    flex: '0 0 auto',
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
    padding: '8px 14px 10px',
    background: '#ffffff',
    borderTop: '1px solid #e2e8f0',
  },
  notesLabel: {
    fontSize: 11,
    fontWeight: 650,
    color: '#647087',
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  notesTextarea: {
    width: '100%',
    height: 64,
    resize: 'vertical',
    border: '1px solid #d8dee9',
    borderRadius: 6,
    padding: '6px 8px',
    font: '13px ui-sans-serif, system-ui, sans-serif',
    color: '#172033',
    boxSizing: 'border-box',
    outline: 'none',
  },
  empty: { margin: 'auto', color: '#6b7587', fontSize: 14 },
  error: { position: 'absolute', left: 16, right: 16, bottom: 14, padding: '9px 12px', color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0', borderRadius: 6, fontSize: 12 },
  hiddenFileInput: {
    position: 'absolute',
    width: 1,
    height: 1,
    padding: 0,
    margin: -1,
    overflow: 'hidden',
    clip: 'rect(0, 0, 0, 0)',
    whiteSpace: 'nowrap',
    border: 0,
  },
  presentButton: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 6,
    flex: '0 0 auto',
    margin: '0 12px 0 4px',
    padding: '0 14px',
    height: 30,
    border: 0,
    borderRadius: 15,
    background: '#1a73e8',
    color: '#ffffff',
    font: '600 13px ui-sans-serif, system-ui, sans-serif',
    cursor: 'pointer',
  },

};

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


interface AnchoredSelection {
  shapeId: string;
  storyId: string;
  anchor: PptxCaretAnchor;
  focus: PptxCaretAnchor;
}

/** Where keystrokes queued behind pending work type. */
interface TypingTarget {
  selection: AnchoredSelection | null;
}

interface KeyInput {
  key: string;
  shiftKey: boolean;
  metaKey: boolean;
  ctrlKey: boolean;
  altKey: boolean;
}

const EDITING_KEYS = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
  'Home',
  'End',
  'Backspace',
  'Delete',
  'Enter',
]);

function isCharacter(input: KeyInput): boolean {
  return (
    input.key.length > 0 &&
    !input.metaKey &&
    !input.ctrlKey &&
    !input.altKey &&
    Array.from(input.key).length === 1
  );
}

function isEditingKey(input: KeyInput): boolean {
  return EDITING_KEYS.has(input.key) || isCharacter(input);
}
