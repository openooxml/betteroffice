/**
 * PagedEditor Component
 *
 * Main paginated editing component that integrates:
 * - YrsInput: browser input translated directly into yrs operations
 * - Layout engine: computes page layout from yrs story blocks
 * - Canvas renderer: replays the Rust display list on visible pages
 * - Selection overlay: renders caret and selection highlights
 *
 * Architecture:
 * 1. User clicks on visible pages → hit test → update the yrs selection
 * 2. User types → YrsInput applies a yrs operation
 * 3. Yrs story blocks → measure → layout → display list
 * 4. Selection changes → compute rects → update overlay
 */

import React, { useEffect, useRef, useState, useCallback, useMemo, forwardRef, memo } from 'react';
import type { CSSProperties } from 'react';
import { createPortal } from 'react-dom';

// Internal components
import {
  YrsInput,
  type YrsDisplaySelection,
  type YrsInputRef,
  type YrsStoredFormatting,
} from './YrsInput';
import { CanvasSelectionOverlay } from './overlays/CanvasSelectionOverlay';
import { CanvasCellSelectionOverlay } from './overlays/CanvasCellSelectionOverlay';
import { CanvasImageSelectionOverlay } from './overlays/CanvasImageSelectionOverlay';
import { CanvasTableResizeOverlay } from './overlays/CanvasTableResizeOverlay';
import {
  CanvasParagraphFlashOverlay,
  type CanvasParagraphFlashRequest,
} from './overlays/CanvasParagraphFlashOverlay';

// Layout engine
import type { Layout } from '@betteroffice/docx/layout/pagination';
import {
  computeAnchorPositionsFromYrs,
  createYrsSidebarProjection,
  extractTrackedChangesFromYrs,
  resolveDisplayPageClientRect,
  type TrackedChangesResult,
  type DisplayList,
  type DisplayListQueries,
  type DisplayListRect,
} from '@betteroffice/docx/layout/render';
import type { ParagraphHighlightOptions, ScrollToParaIdOptions } from '@betteroffice/docx/utils';

// Layout bridge
import { DEFAULT_PAGE_HEIGHT_PX, type BundledFontProvider } from '@betteroffice/docx/layout';

// Selection sync
import { LayoutSelectionGate } from './internals/LayoutSelectionGate';

// Types
import type {
  Document,
  TextFormatting,
  Theme,
  StyleDefinitions,
  SectionProperties,
  HeaderFooter,
} from '@betteroffice/docx/types/document';
import type { WrapType } from '@betteroffice/docx/docx/wrapTypes';
import {
  yrsLocToDisplayPosition as yrsLocToLocalDisplayPosition,
  type YrsInlineFormatDelta,
  type YrsLoc,
  type YrsRenderEnv,
  type YrsResidentCaretSnapshot,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { resolveImageLayoutAttrs } from '@betteroffice/docx/docx';
import type { RenderedDomContext } from '../../plugin-api/types';
import {
  DEFAULT_PAGE_WIDTH,
  DEFAULT_PAGE_GAP,
  VIEWPORT_PADDING_BOTTOM,
  VIEWPORT_PADDING_TOP,
  containerStyles,
  pluginOverlaysStyles,
} from './internals/styles';
import { viewportMinHeightPx } from './internals/scrollUtils';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '../../plugin-api/RenderedDomContext';
import { useLayoutPipeline } from './hooks/useLayoutPipeline';
import type { ResidentFrameApplyResult } from './hooks/useDisplayList';
import { useRustMeasurement, type RustFontChainsProvider } from './hooks/useRustMeasurement';
import { useYrsCoreSession } from './hooks/useYrsCoreSession';
import { useSelectionOverlay } from './hooks/useSelectionOverlay';
import { useImageInteractions } from './hooks/useImageInteractions';
import { usePagedScrollApi } from './hooks/usePagedScrollApi';
import { usePagesPointer } from './hooks/usePagesPointer';
import { usePagedEditorRefApi } from './hooks/usePagedEditorRefApi';
import { useLayoutTriggers } from './hooks/useLayoutTriggers';
import { TableInsertButton } from './overlays/TableInsertButton';
import { HyperlinkPopup, type HyperlinkPopupData } from '../ui/HyperlinkPopup';
import type { FormattingAction } from '../Toolbar';
import {
  applyYrsToolbarFormatting,
  currentYrsToolbarSelection,
  storedYrsToolbarFormatting,
  withStoredYrsFormatting,
  type YrsToolbarSelection,
} from './yrsToolbar';
import {
  currentYrsTableTarget,
  currentYrsSelectionRange,
  setYrsTableProperty,
  setYrsSelectionInCell,
  yrsEmbedIdForProjectedNode,
  yrsImageGeometryForProjectedNode,
  yrsImageTransformForProjectedNode,
  yrsHyperlinkAtSelection,
  yrsSelectionNearTable,
  yrsTableSelectionRange,
  type YrsEditorCommand,
} from './yrsCommands';
import { YrsPositionProjection } from './internals/yrsPositionProjection';
import type { DocxEditorCollaborationOptions } from './types';

export { DEFAULT_PAGE_WIDTH };

function yrsDeltaForTextFormatting(formatting: TextFormatting | undefined): YrsInlineFormatDelta {
  if (!formatting) return {};
  const delta: YrsInlineFormatDelta = {};
  if (formatting.bold !== undefined) delta.bold = formatting.bold;
  if (formatting.italic !== undefined) delta.italic = formatting.italic;
  if (formatting.underline !== undefined) {
    delta.underline = formatting.underline
      ? { style: formatting.underline.style, color: formatting.underline.color?.rgb }
      : null;
  }
  if (formatting.strike !== undefined) delta.strike = formatting.strike;
  if (formatting.color?.rgb) delta.color = { rgb: formatting.color.rgb };
  else if (formatting.color?.themeColor) {
    delta.color = { themeColor: formatting.color.themeColor };
  }
  if (formatting.highlight !== undefined) {
    delta.highlight = formatting.highlight === 'none' ? null : formatting.highlight;
  }
  if (formatting.fontSize !== undefined) delta.fontSize = formatting.fontSize / 2;
  if (formatting.fontFamily) {
    const ascii = formatting.fontFamily.ascii ?? formatting.fontFamily.hAnsi;
    if (ascii) {
      delta.fontFamily = {
        ascii,
        hAnsi: formatting.fontFamily.hAnsi ?? formatting.fontFamily.ascii,
      };
    }
  }
  return delta;
}

// =============================================================================
// TYPES
// =============================================================================

export interface PagedEditorProps {
  /** The document to edit. */
  document: Document | null;
  /** Source document used only when seeding a replacement yrs session. */
  yrsSeedDocument?: Document | null;
  /** Collaboration identity and replica lifecycle callback. */
  collaboration?: DocxEditorCollaborationOptions;
  /** Document styles for style resolution. */
  styles?: StyleDefinitions | null;
  /** Theme for styling. */
  theme?: Theme | null;
  /** Section properties (page size, margins). */
  sectionProperties?: SectionProperties | null;
  /** Body-level final section properties, used after the last explicit section break. */
  finalSectionProperties?: SectionProperties | null;
  /** Header content for all pages (or pages 2+ when titlePg is set). */
  headerContent?: HeaderFooter | null;
  /** Footer content for all pages (or pages 2+ when titlePg is set). */
  footerContent?: HeaderFooter | null;
  /** Header content for first page only (when titlePg is set). */
  firstPageHeaderContent?: HeaderFooter | null;
  /** Footer content for first page only (when titlePg is set). */
  firstPageFooterContent?: HeaderFooter | null;
  /** Whether the editor is read-only. */
  readOnly?: boolean;
  /** Gap between pages in pixels. */
  pageGap?: number;
  /** Zoom level (1 = 100%). */
  zoom?: number;
  /** Callback when document changes. */
  onDocumentChange?: (document: Document) => void;
  /** Callback when the native Yrs undo/redo availability changes. */
  onYrsHistoryChange?: (canUndo: boolean, canRedo: boolean) => void;
  /** Callback when selection changes. */
  onSelectionChange?: (from: number, to: number) => void;
  /** Yrs-authoritative toolbar state, emitted while standard yrs input is active. */
  onYrsSelectionChange?: (selection: YrsToolbarSelection) => void;
  /** Callback when editor is ready. */
  onReady?: (ref: PagedEditorRef) => void;
  /** Callback when rendered DOM context is ready. */
  onRenderedDomContextReady?: (context: RenderedDomContext) => void;
  /** Plugin overlays to render inside the viewport. */
  pluginOverlays?: React.ReactNode;
  /** Callback when header or footer is double-clicked for editing. */
  onHeaderFooterDoubleClick?: (position: 'header' | 'footer', pageNumber?: number) => void;
  /** Active header/footer editing mode (dims body, intercepts body clicks). */
  hfEditMode?: 'header' | 'footer' | null;
  /** Relationship id of the exact HF part active on the edited page. */
  hfEditRId?: string | null;
  /** Called when user clicks the body area while in HF editing mode. */
  onBodyClick?: () => void;
  /** Custom class name. */
  className?: string;
  /** Custom styles. */
  style?: CSSProperties;
  /** Whether comments sidebar is open (shifts document left). */
  commentsSidebarOpen?: boolean;
  /** Sidebar overlay rendered inside the scroll container (scrolls with document). */
  sidebarOverlay?: React.ReactNode;
  /** Ref callback for the scroll container element. */
  scrollContainerRef?: React.Ref<HTMLDivElement>;
  /** Callback when a hyperlink is clicked (for showing popup). */
  onHyperlinkClick?: (data: {
    href: string;
    displayText: string;
    tooltip?: string;
    position: { top: number; left: number };
  }) => void;
  /** Hyperlink popup state (null = hidden). */
  hyperlinkPopupData?: HyperlinkPopupData | null;
  /** Called when user wants to navigate to the link. */
  onHyperlinkPopupNavigate?: (href: string) => void;
  /** Called when user wants to copy the URL. */
  onHyperlinkPopupCopy?: (href: string) => void;
  /** Called when user saves hyperlink edits. */
  onHyperlinkPopupEdit?: (displayText: string, href: string) => void;
  /** Called when user removes the hyperlink. */
  onHyperlinkPopupRemove?: () => void;
  /** Called when the popup should close. */
  onHyperlinkPopupClose?: () => void;
  /** Callback when user right-clicks on the pages (for context menu).
   *  When the right-click target resolves to an image node, `image` carries
   *  the image's PM doc position, current wrap type, current cssFloat (lets
   *  the menu disambiguate Square Left vs Square Right), and — for inline
   *  images only — the rendered EMU offset of the image relative to the
   *  page content origin. The host promotes that offset into the new
   *  anchor's `wp:positionH/V` if the user converts inline → anchor. */
  onContextMenu?: (data: {
    x: number;
    y: number;
    hasSelection: boolean;
    image?: {
      pos: number;
      wrapType: WrapType;
      cssFloat?: 'left' | 'right' | 'none' | null;
      inlinePositionEmu?: { horizontalEmu: number; verticalEmu: number };
    } | null;
  }) => void;
  /** Callback with pre-computed Y positions for comment/tracked-change anchors (for sidebar positioning without DOM queries). */
  onAnchorPositionsChange?: (positions: Map<string, number>) => void;
  /** Comment ids whose sticky yrs coverage should be projected into sidebar anchors. */
  sidebarCommentIds?: readonly (string | number)[];
  /** yrs-authoritative tracked-change list, emitted while standard yrs input is active. */
  onYrsTrackedChangesChange?: (result: TrackedChangesResult) => void;
  /** Sticky yrs HF selection, expressed in the active HF root's display positions. */
  onYrsHfSelectionChange?: (rId: string, selection: { from: number; to: number }) => void;
  /**
   * Callback fired when the page count changes after a layout pass.
   * Parents use this to keep their own page counters (e.g. scroll indicator,
   * `getTotalPages()` ref method) in sync without having to poll `getLayout()`.
   */
  onTotalPagesChange?: (totalPages: number) => void;
  /** Layout of each pass (null on reset) — canvas renderer plumbing. */
  onLayoutComputed?: (layout: Layout | null, engine?: YrsSession | null) => void;
  /** One-call resident body-text edit supplied by the canvas frame owner. */
  applyResidentInput?: (text: string) => Promise<ResidentFrameApplyResult | null>;
  /** One-call resident body-text deletion supplied by the canvas frame owner. */
  applyResidentDelete?: (
    direction: 'backward' | 'forward'
  ) => Promise<ResidentFrameApplyResult | null>;
  /** Set of resolved comment IDs — hides highlight for these comments */
  resolvedCommentIds?: Set<number>;
  /** Suggestion mode active state */
  isSuggesting?: boolean;
  /** Active author for suggestion mode */
  author?: string;
  /** Bundled metric-compatible font provider for Rust measurement. */
  measurementFontProvider?: BundledFontProvider;
  /**
   * Host slot the Rust measure source fills with the merged doc-wide font
   * chains; the canvas display-list build (mounted in DocxEditor) reads it to
   * gate GlyphRun emission. Undefined ⇒ the measure source has nowhere to
   * publish its chains (no canvas renderer to feed).
   */
  rustFontChainsProviderRef?: React.RefObject<RustFontChainsProvider | null>;
  /**
   * Display-list query source (null until the first display-list build lands).
   * Routes pointer resolution, selection rects / caret, and sidebar anchor Ys
   * through Rust queries over the display list.
   */
  displayListQueries?: DisplayListQueries | null;
  canvasDisplayList?: DisplayList | null;
  displayListFrameEpoch?: number | null;
  residentCaret?: YrsResidentCaretSnapshot | null;
  residentCaretAuthoritative?: boolean;
  /** Worker-painted caret line is on screen; hide the DOM blink caret. */
  paintedCaretActive?: boolean;
  /** Document-mutating input notification for the painted-caret mode machine. */
  onCaretInput?(): void;
  /** Text input dispatched: hide the DOM caret before the worker round-trip. */
  onCaretInputDispatched?(): void;
  /** Selection move / blur / IME / mode change: swap to the DOM caret now. */
  onCaretInterrupt?(): void;
  /** `.canvas-pages` host element — canvas-path pointer events attach here. */
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  /**
   * Portal target for the caret / selection overlay — resolves to
   * `editorContentRef.current` (the positioned ancestor sharing the
   * `.canvas-pages` host's top-left). The selection overlay portals here so it
   * shares the visible canvas pages' coordinate space.
   */
  canvasOverlayTarget?: HTMLElement | null;
}

export interface PagedEditorRef {
  /** Get the current document (projected from yrs on demand in direct-input mode). */
  getDocument(): Document | null;
  /** Focus the editor. */
  focus(): void;
  /** Blur the editor. */
  blur(): void;
  /** Check if focused. */
  isFocused(): boolean;
  /** Undo. */
  undo(): boolean;
  /** Redo. */
  redo(): boolean;
  /** Check whether the authoritative input surface has undo history. */
  canUndo(): boolean;
  /** Check whether the authoritative input surface has redo history. */
  canRedo(): boolean;
  /** Set selection by display position. */
  setSelection(anchor: number, head?: number): void;
  /** Insert plain text through the authoritative Yrs input path. */
  insertText(text: string): void;
  /** Delete the current authoritative selection. */
  deleteSelection(): void;
  /** Select the full active Yrs story. */
  selectAll(): void;
  /** Get the current display-position selection. */
  getSelectionRange(): { from: number; to: number } | null;
  /** Resolve a display position into the authoritative Yrs location. */
  displayPositionToYrsLoc(position: number): YrsLoc | null;
  /** Live authoritative yrs session. */
  getYrsSession(): YrsSession | null;
  /** Paragraph-local stored inline formatting for the current yrs caret. */
  getYrsStoredFormatting(): YrsStoredFormatting | null;
  /** Resolve a live yrs Loc to the display position used by overlays. */
  yrsLocToDisplayPosition(loc: YrsLoc): number | null;
  /** Publish a yrs selection/mutation through the direct-input refresh path. */
  syncYrsInputState(docChanged: boolean): boolean;
  /** Apply a body-toolbar command through yrs. Never falls back to PM. */
  applyYrsFormatting(action: FormattingAction): boolean;
  /** Apply a non-toolbar body command through yrs. Never falls back to PM. */
  applyYrsCommand(command: YrsEditorCommand): boolean;
  /** Get current layout. */
  getLayout(): Layout | null;
  /** Force re-layout. */
  relayout(): void;
  /** Scroll the visible pages to bring a display position into view. */
  scrollToPosition(position: number): void;
  /**
   * Scroll to the paragraph identified by Word `w14:paraId`.
   * Pass `options.highlight` to briefly flash rendered paragraph fragments.
   * @returns whether a matching paragraph was found
   */
  scrollToParaId(paraId: string, options?: ScrollToParaIdOptions): boolean;
  /**
   * Scroll the paginated view so `pageNumber` (1-indexed) is in view.
   * No-op if the layout isn't ready yet or pageNumber is out of range.
   */
  scrollToPage(pageNumber: number): void;
  /**
   * Scroll to the comment identified by `commentId` and select its range so
   * the selection overlay highlights it. Resolves the id to a live Yrs range;
   * returns `false` (not a throw, not a silent no-op)
   * when the id no longer resolves so the caller can surface a "location no
   * longer exists" affordance.
   */
  scrollToCommentId(commentId: number): boolean;
  /**
   * Scroll to the tracked change identified by `revisionId` and select its
   * range so the selection overlay highlights it. Resolves the id to a live
   * Yrs range; returns `false` when the id no longer
   * resolves (the change was accepted/rejected/deleted).
   */
  scrollToChangeId(revisionId: number): boolean;
  /**
   * Select the display-position range `[from, to]` so the selection overlay
   * highlights it, and scroll its start into view. No-op for a malformed
   * range or a `from` past the document end; `to` is clamped to the document
   * size (raw caller positions, so out-of-range must not throw).
   */
  highlightRange(from: number, to: number): void;
}

// =============================================================================
// COMPONENT (module-scope helpers live in per-domain files — see imports)
// =============================================================================

/**
 * PagedEditor - Main paginated editing component.
 */
const PagedEditorComponent = forwardRef<PagedEditorRef, PagedEditorProps>(
  function PagedEditor(props, ref) {
    const {
      document,
      yrsSeedDocument = document,
      collaboration,
      styles,
      theme: _theme,
      headerContent,
      footerContent,
      firstPageHeaderContent,
      firstPageFooterContent,
      readOnly = false,
      pageGap = DEFAULT_PAGE_GAP,
      zoom = 1,
      onDocumentChange,
      onYrsHistoryChange,
      onSelectionChange,
      onYrsSelectionChange,
      onYrsHfSelectionChange,
      onReady,
      onRenderedDomContextReady,
      pluginOverlays,
      onHeaderFooterDoubleClick,
      hfEditMode,
      hfEditRId,
      onBodyClick,
      className,
      style,
      commentsSidebarOpen = false,
      sidebarOverlay,
      scrollContainerRef: scrollContainerRefProp,
      onHyperlinkClick,
      onContextMenu,
      onAnchorPositionsChange,
      sidebarCommentIds = [],
      onYrsTrackedChangesChange,
      onTotalPagesChange,
      onLayoutComputed,
      applyResidentInput,
      applyResidentDelete,
      hyperlinkPopupData,
      onHyperlinkPopupNavigate,
      onHyperlinkPopupCopy,
      onHyperlinkPopupEdit,
      onHyperlinkPopupRemove,
      onHyperlinkPopupClose,
      isSuggesting = false,
      author = 'User',
      measurementFontProvider,
      rustFontChainsProviderRef,
      displayListQueries = null,
      canvasDisplayList = null,
      displayListFrameEpoch = null,
      residentCaret = null,
      residentCaretAuthoritative = false,
      paintedCaretActive = false,
      onCaretInput,
      onCaretInputDispatched,
      onCaretInterrupt,
      canvasHostRef,
      canvasOverlayTarget = null,
    } = props;
    const yrsStyleResolver = useMemo(() => (styles ? createStyleResolver(styles) : null), [styles]);

    // Resolve the scroll container: prefer parent-provided ref, fallback to own container
    const getScrollContainer = useCallback((): HTMLDivElement | null => {
      if (scrollContainerRefProp && typeof scrollContainerRefProp === 'object') {
        return (scrollContainerRefProp as React.RefObject<HTMLDivElement | null>).current;
      }
      return containerRef.current;
    }, [scrollContainerRefProp]);

    // Refs
    const containerRef = useRef<HTMLDivElement>(null);
    const pagesContainerRef = useRef<HTMLDivElement>(null);
    useEffect(() => {
      if (!canvasHostRef || canvasHostRef.current) return;
      const host = pagesContainerRef.current;
      if (!host) return;
      canvasHostRef.current = host;
      return () => {
        if (canvasHostRef.current === host) canvasHostRef.current = null;
      };
    }, [canvasHostRef]);
    /** Viewport wrapper: sync minHeight/marginBottom in layout pipeline before scroll restore. */
    const viewportLayoutRef = useRef<HTMLDivElement>(null);
    const yrsInputRef = useRef<YrsInputRef>(null);

    const yrsCore = useYrsCoreSession(true, document, yrsSeedDocument, collaboration);
    const yrsRenderEnv = useMemo<YrsRenderEnv>(() => {
      const themeColors: Record<string, string> = {};
      for (const [name, value] of Object.entries(_theme?.colorScheme ?? {})) {
        if (typeof value === 'string') themeColors[name] = value;
      }
      return {
        themeColors,
        defaultTabStopTwips: document?.package.settings?.defaultTabStop ?? null,
        numericIds: {},
      };
    }, [_theme?.colorScheme, document?.package.settings?.defaultTabStop]);
    const activeYrsRootStory = hfEditMode && hfEditRId ? `hf:${hfEditRId}` : 'body';
    const yrsInputPositionMap = useCallback(
      (storyId = activeYrsRootStory) => yrsCore.inputPositionMap(storyId),
      [activeYrsRootStory, yrsCore.inputPositionMap]
    );
    const yrsDisplayPositionToLoc = useCallback(
      (position: number, story = 'body') => yrsCore.displayPositionToLoc(position, story),
      [yrsCore.displayPositionToLoc]
    );
    const getYrsPositionProjectionRef = useRef<(rootStory: string) => YrsPositionProjection | null>(
      () => null
    );
    const yrsLocToDisplayPosition = useCallback(
      (loc: YrsLoc): number | null => {
        const map = yrsCore.inputPositionMap(loc.story);
        if (!map) return null;
        const local = yrsLocToLocalDisplayPosition(map, loc);
        const rootStory = loc.story.startsWith('hf:') ? activeYrsRootStory : 'body';
        if (loc.story === rootStory) return local;
        return getYrsPositionProjectionRef.current(rootStory)?.positionForLoc(loc) ?? null;
      },
      [activeYrsRootStory, yrsCore.inputPositionMap]
    );
    // Store callbacks in refs to avoid infinite re-render loops
    // when parent passes unstable callback references
    const onSelectionChangeRef = useRef(onSelectionChange);
    const onYrsSelectionChangeRef = useRef(onYrsSelectionChange);
    const onYrsHfSelectionChangeRef = useRef(onYrsHfSelectionChange);
    const onDocumentChangeRef = useRef(onDocumentChange);
    const onYrsHistoryChangeRef = useRef(onYrsHistoryChange);
    const onReadyRef = useRef(onReady);
    // Keep refs in sync with latest props
    onSelectionChangeRef.current = onSelectionChange;
    onYrsSelectionChangeRef.current = onYrsSelectionChange;
    onYrsHfSelectionChangeRef.current = onYrsHfSelectionChange;
    onDocumentChangeRef.current = onDocumentChange;
    onYrsHistoryChangeRef.current = onYrsHistoryChange;
    onReadyRef.current = onReady;

    useEffect(() => {
      const session = yrsCore.session;
      onYrsHistoryChangeRef.current?.(session?.canUndo() ?? false, session?.canRedo() ?? false);
    }, [yrsCore.session]);

    // State
    const [isFocused, setIsFocused] = useState(false);

    useEffect(() => {
      if (hfEditMode) setIsFocused(false);
    }, [hfEditMode]);

    useEffect(() => {
      if (!isFocused) onCaretInterrupt?.();
    }, [isFocused, onCaretInterrupt]);

    // Read-only / suggesting / HF-edit transitions bypass the resident input
    // path, so any painted caret line would go stale — swap to the DOM caret.
    useEffect(() => {
      onCaretInterrupt?.();
    }, [readOnly, isSuggesting, hfEditMode, onCaretInterrupt]);

    // Image selection state — `isImageInteractingRef` lives at the parent so
    // useSelectionOverlay can read it (to gate the deferred image-info clear)
    // while useImageInteractions writes it (during drag / resize).
    const isImageInteractingRef = useRef(false);

    // Selection gate - ensures selection renders only when layout is current
    const syncCoordinator = useMemo(() => new LayoutSelectionGate(), []);
    const getYrsSelectionHead = useCallback(
      () => yrsInputRef.current?.displaySelection()?.head ?? 0,
      []
    );

    // Rust measurement — the sole measurement path. Wiring lives in the
    // hook; engine-ready and fonts-ready re-layouts reach back through its
    // runLayoutPipelineRef, and deferLayoutPass gates provisional passes.
    const {
      deferLayoutPass,
      residentMeasurementConfig,
      runLayoutPipelineRef,
    } = useRustMeasurement({
      document,
      fontProvider: measurementFontProvider,
      fontChainsProviderRef: rustFontChainsProviderRef,
      textEngine: yrsCore.session,
    });

    // Layout pipeline — owns layout/blocks/measures state, the rAF-coalesced
    // scheduler, scroll-restore plumbing, and the page-count
    // notifier.
    const publishResidentLayout = useCallback(
      (nextLayout: Layout | null) => onLayoutComputed?.(nextLayout, yrsCore.session),
      [onLayoutComputed, yrsCore.session]
    );
    const { layout, runLayoutPipeline, scheduleLayout } = useLayoutPipeline({
      document,
      session: yrsCore.session,
      renderEnv: yrsRenderEnv,
      pageGap,
      zoom,
      deferLayoutPass,
      residentMeasurementConfig,
      displayListQueries,
      interactionPageHostRef: canvasHostRef,
      pagesContainerRef,
      viewportLayoutRef,
      getSelectionHead: getYrsSelectionHead,
      syncCoordinator,
      getScrollContainer,
      onTotalPagesChange,
      onLayoutComputed: publishResidentLayout,
      onAnchorPositionsChange,
    });
    runLayoutPipelineRef.current = yrsCore.session ? runLayoutPipeline : null;

    // Selection overlay — caret, range rects, image overlay info, plus the
    // ResizeObserver + post-layout recompute that keep geometry fresh.
    const getYrsDisplaySelection = useCallback(
      () => yrsInputRef.current?.displaySelection() ?? null,
      []
    );
    const getYrsStickySelection = useCallback(
      () => yrsCore.session?.selection() ?? null,
      [yrsCore.session]
    );
    const {
      selectionRects,
      caretPosition,
      setSelectionRects,
      setCaretPosition,
      updateSelectionOverlay,
    } = useSelectionOverlay({
      layout,
      containerRef,
      syncCoordinator,
      displayListQueries,
      displayListFrameEpoch,
      residentCaret,
      residentCaretAuthoritative,
      getYrsDisplaySelection,
      getYrsStickySelection,
    });
    const updateSelectionOverlayRef = useRef(updateSelectionOverlay);
    updateSelectionOverlayRef.current = updateSelectionOverlay;

    // Route body focus to the yrs textarea input surface.
    const focusBodyInput = useCallback((): void => {
      yrsInputRef.current?.focus();
    }, []);

    const yrsProjectionVersionRef = useRef(0);
    const lastYrsToolbarSelectionKeyRef = useRef<string | null>(null);
    const lastPublishedBodySelectionKeyRef = useRef<string | null>(null);
    const documentChangeTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
    const publishYrsDirectInput = useCallback((): void => {
      yrsCore.publishDirectInput();
      // Structural input can mint a paragraph before the existing projection
      // can map its new sticky caret. Invalidate first so emitSelection can
      // rebuild the projection and reach the normal layout-refresh callback.
      yrsProjectionVersionRef.current += 1;
    }, [yrsCore.publishDirectInput]);
    useEffect(
      () => () => {
        if (documentChangeTimerRef.current !== null) {
          clearTimeout(documentChangeTimerRef.current);
        }
      },
      []
    );
    const refreshYrsLayout = useCallback((): void => {
      yrsProjectionVersionRef.current += 1;
      syncCoordinator.incrementStateSeq();
      syncCoordinator.requestRender();
      scheduleLayout();
    }, [scheduleLayout, syncCoordinator]);

    /**
     * Publish sticky yrs selection geometry and refresh the yrs-backed layout.
     */
    const handleYrsStateChange = useCallback(
      (
        selection: YrsDisplaySelection,
        docChanged: boolean,
        residentLayoutReady = false,
        residentCaretReady = false
      ): void => {
        const session = yrsCore.session;
        if (session) {
          onYrsHistoryChangeRef.current?.(session.canUndo(), session.canRedo());
        }
        const liveStory = session?.selection()?.head.story ?? activeYrsRootStory;
        const inputMap = yrsCore.inputPositionMap(liveStory);
        if (session && inputMap) {
          const toolbarSelection = currentYrsToolbarSelection(session, inputMap);
          if (toolbarSelection) {
            const nextToolbarSelection = withStoredYrsFormatting(
              toolbarSelection,
              yrsInputRef.current?.storedFormatting() ?? null
            );
            const key = JSON.stringify([
              nextToolbarSelection.context,
              nextToolbarSelection.tableContext,
              nextToolbarSelection.range.story,
              nextToolbarSelection.startParagraphIndex,
              nextToolbarSelection.endParagraphIndex,
            ]);
            if (lastYrsToolbarSelectionKeyRef.current !== key) {
              lastYrsToolbarSelectionKeyRef.current = key;
              onYrsSelectionChangeRef.current?.(nextToolbarSelection);
            }
          }
        }

        if (activeYrsRootStory.startsWith('hf:') && hfEditRId) {
          onYrsHfSelectionChangeRef.current?.(hfEditRId, {
            from: selection.anchor,
            to: selection.head,
          });
        } else {
          const stickySelection = session?.selection();
          const selectionKey = stickySelection
            ? JSON.stringify([stickySelection.anchor, stickySelection.head])
            : `${selection.anchor}:${selection.head}`;
          const selectionChanged = lastPublishedBodySelectionKeyRef.current !== selectionKey;
          lastPublishedBodySelectionKeyRef.current = selectionKey;
          // Publish body selection before the asynchronous layout/sidebar
          // projection. The deleted PM follower used to provide this ordering;
          // React state now consumes the authoritative yrs positions directly.
          if (!docChanged && selectionChanged) {
            onSelectionChangeRef.current?.(selection.anchor, selection.head);
          }
        }
        // Layout reads the same authoritative selection, but selection state
        // itself is published above from the yrs event rather than inferred
        // later from a projection rebuild.
        if (!residentCaretReady) updateSelectionOverlay();
        if (docChanged && !residentLayoutReady) {
          refreshYrsLayout();
        }
        if (docChanged) {
          // Projecting the complete document is intentionally kept off the
          // synchronous input path. Besides being unnecessary for canvas
          // rendering, doing it once per key starves the browser event loop on
          // larger documents. Coalesce host onChange notifications instead.
          if (documentChangeTimerRef.current !== null) {
            clearTimeout(documentChangeTimerRef.current);
          }
          documentChangeTimerRef.current = setTimeout(() => {
            documentChangeTimerRef.current = null;
            const projected = yrsCore.documentFromYrs();
            if (projected) onDocumentChangeRef.current?.(projected);
          }, 100);
        }
      },
      [
        activeYrsRootStory,
        hfEditRId,
        refreshYrsLayout,
        updateSelectionOverlay,
        yrsCore.documentFromYrs,
        yrsCore.inputPositionMap,
        yrsCore.session,
      ]
    );

    const syncYrsInputState = useCallback(
      (docChanged: boolean): boolean => {
        if (!yrsCore.session) return false;
        const displaySelection = yrsInputRef.current?.displaySelection() ?? { anchor: 0, head: 0 };
        if (docChanged) {
          yrsCore.publishDirectInput();
        }
        handleYrsStateChange(displaySelection, docChanged);
        return true;
      },
      [handleYrsStateChange, yrsCore.publishDirectInput, yrsCore.session]
    );

    useEffect(() => {
      const session = yrsCore.session;
      if (!session) return;
      return session.onUpdate((_update, origin) => {
        if (origin === 'remote') syncYrsInputState(true);
      });
    }, [syncYrsInputState, yrsCore.session]);

    const applyYrsFormatting = useCallback(
      (action: FormattingAction): boolean => {
        if (!yrsCore.session) return false;
        const liveStory = yrsCore.session.selection()?.head.story ?? activeYrsRootStory;
        const map = yrsCore.inputPositionMap(liveStory);
        if (!map) return false;
        try {
          const structuralAuthor = isSuggesting
            ? { name: author, date: new Date().toISOString() }
            : undefined;
          const selection = currentYrsToolbarSelection(yrsCore.session, map);
          if (
            selection &&
            typeof action === 'object' &&
            action.type === 'applyStyle' &&
            yrsStyleResolver
          ) {
            const resolved = yrsStyleResolver.resolveParagraphStyle(action.value);
            const delta = yrsDeltaForTextFormatting(resolved.runFormatting);
            yrsCore.session.applyParagraphStyle(selection.range, action.value, structuralAuthor);
            if (selection.context.hasSelection) {
              if (Object.keys(delta).length > 0) {
                yrsCore.session.formatRange(selection.range, delta);
              }
              yrsInputRef.current?.clearStoredFormatting();
            } else if (Object.keys(delta).length > 0) {
              yrsInputRef.current?.applyStoredFormatting({ type: 'set', delta });
            }
            if (!syncYrsInputState(true)) return false;
            yrsInputRef.current?.focus();
            return true;
          }
          if (selection && !selection.context.hasSelection) {
            const storedAction = storedYrsToolbarFormatting(selection.context, action);
            if (storedAction) {
              yrsInputRef.current?.applyStoredFormatting(storedAction);
              yrsInputRef.current?.focus();
              return true;
            }
          }
          const changed = applyYrsToolbarFormatting(yrsCore.session, map, action, structuralAuthor);
          if (!changed) return false;
          yrsInputRef.current?.clearStoredFormatting();
          if (!syncYrsInputState(true)) return false;
          yrsInputRef.current?.focus();
          return true;
        } catch (error) {
          console.error('[yrs] toolbar formatting failed', error);
          return false;
        }
      },
      [
        activeYrsRootStory,
        author,
        isSuggesting,
        syncYrsInputState,
        yrsStyleResolver,
        yrsCore.inputPositionMap,
        yrsCore.session,
      ]
    );

    const applyYrsCommand = useCallback(
      (command: YrsEditorCommand): boolean => {
        const session = yrsCore.session;
        if (!session || activeYrsRootStory !== 'body') return false;
        const displaySelection = yrsInputRef.current?.displaySelection() ?? { anchor: 0, head: 0 };
        const positionProjection = getYrsPositionProjectionRef.current('body');
        if (!positionProjection) return false;
        try {
          const structuralAuthor = isSuggesting
            ? { name: author, date: new Date().toISOString() }
            : undefined;
          let docChanged = true;
          if (command.type === 'imageGeometry') {
            const node = positionProjection.nodeAt(command.pmPos);
            if (!node || node.kind !== 'image') return false;
            const embedId = yrsEmbedIdForProjectedNode(node);
            const geometry = yrsImageGeometryForProjectedNode(node, command.patch);
            if (!embedId || !geometry) return false;
            session.setImageGeometry(embedId, geometry);
          } else if (command.type === 'imageWrap') {
            const node = positionProjection.nodeAt(command.pmPos);
            if (!node || node.kind !== 'image') return false;
            const embedId = yrsEmbedIdForProjectedNode(node);
            const patch = resolveImageLayoutAttrs(command.target, node.attrs, command.options);
            const geometry = yrsImageGeometryForProjectedNode(node, patch);
            if (!embedId || !geometry) return false;
            session.setImageGeometry(embedId, geometry);
          } else if (command.type === 'imageTransform') {
            const node = positionProjection.nodeAt(command.pmPos);
            if (!node || node.kind !== 'image') return false;
            const embedId = yrsEmbedIdForProjectedNode(node);
            const transform = yrsImageTransformForProjectedNode(node, command.action);
            const geometry = yrsImageGeometryForProjectedNode(node, { transform });
            if (!embedId || !geometry) return false;
            session.setImageGeometry(embedId, geometry);
          } else if (command.type === 'insertImage') {
            const at = session.selection()?.head;
            if (!at) return false;
            session.insertImage(at, command.image, structuralAuthor);
            session.setSelection({ ...at, offset: at.offset + 1 });
          } else if (command.type === 'contentControlValue') {
            const node = positionProjection.nodeAt(command.pmPos);
            const embedId = command.embedId ?? (node ? yrsEmbedIdForProjectedNode(node) : null);
            if (embedId) {
              session.setContentControlValue(embedId, command.value);
            } else {
              if (!node || (node.kind !== 'sdt' && node.kind !== 'blockSdt')) return false;
              const target = positionProjection.targetAt(node.start);
              const at = yrsCore.displayPositionToLoc(target.displayPosition, target.story);
              if (!at) return false;
              session.setContentControlValueAt(at, command.value);
            }
          } else if (command.type === 'paragraphAttrs' || command.type === 'removeTabStop') {
            const map = yrsCore.inputPositionMap('body');
            const selection = map ? currentYrsToolbarSelection(session, map) : null;
            if (!selection) return false;
            if (command.type === 'paragraphAttrs') {
              session.setParagraphAttrs(selection.range, command.attrs, structuralAuthor);
            } else {
              const currentTabs = selection.context.paragraphProperties.tabs;
              const tabs = Array.isArray(currentTabs)
                ? currentTabs.filter(
                    (tab) =>
                      tab != null &&
                      typeof tab === 'object' &&
                      Number((tab as { position?: unknown }).position) !== command.positionTwips
                  )
                : [];
              session.setParagraphAttrs(
                selection.range,
                { tabs: tabs.length > 0 ? (tabs as never) : null },
                structuralAuthor
              );
            }
          } else if (command.type === 'setHyperlink') {
            const selected = currentYrsSelectionRange(session);
            const existing = command.editExisting
              ? yrsHyperlinkAtSelection(session, command.matchHref)
              : null;
            if (!selected) return false;
            let range = existing?.range ?? selected;
            const collapsed =
              range.start.paraId === range.end.paraId && range.start.offset === range.end.offset;
            if ((collapsed || existing) && command.displayText) {
              const at = { story: range.story, ...range.start };
              if (existing) session.replaceRange(range, command.displayText, structuralAuthor);
              else session.insertText(at, command.displayText, structuralAuthor);
              range = {
                story: at.story,
                start: { paraId: at.paraId, offset: at.offset },
                end: { paraId: at.paraId, offset: at.offset + command.displayText.length },
              };
            }
            if (
              range.start.paraId === range.end.paraId &&
              range.start.offset === range.end.offset
            ) {
              return false;
            }
            session.setHyperlink(range, {
              href: command.href,
              tooltip: command.tooltip ?? null,
              rId: null,
            });
            session.setSelection({ story: range.story, ...range.end });
          } else if (command.type === 'removeHyperlink') {
            const selected = currentYrsSelectionRange(session);
            if (!selected) return false;
            const collapsed =
              selected.start.paraId === selected.end.paraId &&
              selected.start.offset === selected.end.offset;
            const range = collapsed
              ? yrsHyperlinkAtSelection(session, command.href)?.range
              : selected;
            if (!range) return false;
            session.setHyperlink(range, null);
          } else if (command.type === 'insertPageBreak' || command.type === 'insertSectionBreak') {
            const at = session.selection()?.head;
            if (!at) return false;
            if (command.type === 'insertPageBreak') {
              // The native renderer consumes page breaks at block boundaries.
              // Split a non-empty prefix first so the facade insert lands
              // between paragraph pilcrows without writing through PM.
              const breakAt =
                at.offset > 0
                  ? {
                      story: at.story,
                      paraId: session.splitParagraph(at).secondParaId,
                      offset: 0,
                    }
                  : at;
              session.setSelection(breakAt);
              session.insertPageBreak(breakAt);
            } else session.insertSectionBreak(at, command.breakType);
          } else if (command.type === 'insertTable') {
            const at = session.selection()?.head;
            if (!at || command.rows < 1 || command.columns < 1) return false;
            // Tables are block embeds in the yrs save projection. When the
            // caret follows text, split first so the embed lands between two
            // paragraph pilcrows instead of inside the current paragraph.
            const tableAt =
              at.offset > 0
                ? {
                    story: at.story,
                    paraId: session.splitParagraph(at).secondParaId,
                    offset: 0,
                  }
                : at;
            const receipt = session.insertTable(
              tableAt,
              command.rows,
              command.columns,
              structuralAuthor
            );
            const firstCell = {
              story: receipt.table.story,
              tableIndex: receipt.table.tableIndex,
              row: 0,
              column: 0,
            };
            session.setCellSelection({ anchor: firstCell, head: firstCell });
            setYrsSelectionInCell(session, firstCell);
          } else if (command.type === 'tableSetBorders') {
            const target = currentYrsTableTarget(session);
            if (!target) return false;
            session.setCellBorders(target.range, command.borders);
          } else if (command.type === 'tableProperties') {
            const target = currentYrsTableTarget(session);
            if (!target) return false;
            if (
              !setYrsTableProperty(
                session,
                { story: target.focused.story, tableIndex: target.focused.tableIndex },
                'tblPr',
                command.properties
              )
            ) {
              return false;
            }
          } else if (command.type === 'tableColumnWidths') {
            if (command.widths.length === 0) return false;
            const table = positionProjection.tableAtStart(command.pmStart);
            if (!table) return false;
            const targets = command.widths.map((width) => ({
              ...width,
              at: {
                story: table.story,
                tableIndex: table.tableIndex,
                row: 0,
                column: width.column,
              },
            }));
            for (const target of targets) {
              session.setColumnWidth(target.at, target.widthTwips);
            }
            setYrsTableProperty(
              session,
              { story: targets[0].at.story, tableIndex: targets[0].at.tableIndex },
              'tblPr',
              { tableLayout: 'fixed' }
            );
          } else {
            const target = currentYrsTableTarget(session);
            if (!target) return false;
            const table = { story: target.focused.story, tableIndex: target.focused.tableIndex };
            const nearTable = yrsSelectionNearTable(session, table);

            if (command.type === 'tableInsertRow') {
              const at = command.at ?? target.focused;
              session.setCellSelection({ anchor: at, head: at });
              session.insertRow(at, command.side, structuralAuthor);
            } else if (command.type === 'tableInsertColumn') {
              const at = command.at ?? target.focused;
              session.setCellSelection({ anchor: at, head: at });
              session.insertColumn(at, command.side);
            } else if (command.type === 'tableDeleteRow') {
              const receipt = session.deleteRow(target.range, structuralAuthor);
              if (receipt.deletedTable) {
                if (nearTable) session.setSelection(nearTable);
              } else {
                const surviving = session.cellSelection()?.anchor ?? target.focused;
                setYrsSelectionInCell(session, surviving);
              }
            } else if (command.type === 'tableDeleteColumn') {
              const receipt = session.deleteColumn(target.range);
              if (receipt.deletedTable) {
                if (nearTable) session.setSelection(nearTable);
              } else {
                const surviving = session.cellSelection()?.anchor ?? target.focused;
                setYrsSelectionInCell(session, surviving);
              }
            } else if (command.type === 'tableDelete') {
              session.deleteTable(table);
              if (nearTable) session.setSelection(nearTable);
            } else if (command.type === 'tableMergeCells') {
              if (
                target.range.anchor.row === target.range.head.row &&
                target.range.anchor.column === target.range.head.column
              ) {
                return false;
              }
              session.mergeCells(target.range);
              const surviving = session.cellSelection()?.anchor ?? target.range.anchor;
              setYrsSelectionInCell(session, surviving);
            } else if (command.type === 'tableSplitCell') {
              session.splitCell(target.focused, command.rows, command.columns);
            } else if (command.type === 'tableCellShading') {
              session.setCellShading(target.range, command.color);
            } else {
              const range = yrsTableSelectionRange(session, target.focused, command.target);
              if (!range) return false;
              session.setCellSelection(range);
              docChanged = false;
            }
          }

          if (docChanged) {
            yrsCore.publishDirectInput();
          }
          handleYrsStateChange(
            yrsInputRef.current?.displaySelection() ?? displaySelection,
            docChanged
          );
          yrsInputRef.current?.focus();
          return true;
        } catch (error) {
          console.error('[yrs] non-toolbar command failed', error);
          return false;
        }
      },
      [
        activeYrsRootStory,
        author,
        handleYrsStateChange,
        isSuggesting,
        yrsCore.inputPositionMap,
        yrsCore.publishDirectInput,
        yrsCore.session,
      ]
    );

    // Paragraph flash (G2): scrollToParaId's flash is drawn as an overlay over
    // display-list geometry. The scroll API bumps the nonce to (re)start; the
    // overlay clears itself after the flash and calls back to drop the request.
    const [canvasFlashRequest, setCanvasFlashRequest] =
      useState<CanvasParagraphFlashRequest | null>(null);
    const requestCanvasParagraphFlash = useCallback(
      (req: { from: number; to: number; options?: ParagraphHighlightOptions }) => {
        setCanvasFlashRequest((prev) => ({
          from: req.from,
          to: req.to,
          color: req.options?.color,
          durationMs: req.options?.durationMs,
          nonce: (prev?.nonce ?? 0) + 1,
        }));
      },
      []
    );
    const handleCanvasFlashDone = useCallback((nonce: number) => {
      setCanvasFlashRequest((prev) => (prev && prev.nonce === nonce ? null : prev));
    }, []);

    // Scroll API exposed via the PagedEditorRef. Owns the AbortController
    // chain that lets a fresh scroll supersede an in-flight paint-settle.
    const { scrollToPositionImpl, scrollToPageImpl, scrollToParaIdImpl } = usePagedScrollApi({
      pagesContainerRef,
      yrsInputRef,
      yrsSession: yrsCore.session,
      yrsLocToDisplayPosition,
      getScrollContainer,
      displayListQueries,
      canvasHostRef,
      requestCanvasParagraphFlash,
    });

    // Display-list positions retain the document tree's integer coordinate
    // space. Build a lightweight index directly from the authoritative yrs
    // stories and cache it across gestures until the next mutation.
    const yrsPositionProjectionCacheRef = useRef<{
      version: number;
      session: YrsSession;
      rootStory: string;
      projection: YrsPositionProjection;
    } | null>(null);
    const getYrsPositionProjection = useCallback(
      (rootStory: string): YrsPositionProjection | null => {
        const session = yrsCore.session;
        if (!session) return null;
        const cached = yrsPositionProjectionCacheRef.current;
        if (
          cached?.version === yrsProjectionVersionRef.current &&
          cached.session === session &&
          cached.rootStory === rootStory
        ) {
          return cached.projection;
        }
        const projection = new YrsPositionProjection(session, rootStory);
        yrsPositionProjectionCacheRef.current = {
          version: yrsProjectionVersionRef.current,
          session,
          rootStory,
          projection,
        };
        return projection;
      },
      [yrsCore.session]
    );
    getYrsPositionProjectionRef.current = getYrsPositionProjection;

    // Pointer routing — every mouse path on the visible pages: cursor
    // placement, drag-to-select (with cell-selection promotion), table
    // resize handles, the floating "+" insert button, hyperlink clicks,
    // header/footer double-clicks, word/paragraph multi-click, and
    // right-click → host context-menu.
    const {
      handlePagesContextMenu,
      handleTableInsertClick,
      tableInsertButton,
      clearTableInsertTimer,
      hideTableInsertButton,
      getPositionFromMouse,
    } = usePagesPointer({
      pagesContainerRef,
      yrsInputRef,
      yrsSession: yrsCore.session!,
      yrsRootStory: activeYrsRootStory,
      getYrsPositionProjection,
      applyYrsCommand,
      syncYrsInputState,
      readOnly,
      hfEditMode,
      displayListQueries,
      canvasHostRef,
      canvasOverlayTarget,
      onBodyClick,
      onContextMenu,
      onHyperlinkClick,
      onHeaderFooterDoubleClick,
      setSelectionRects,
      setCaretPosition,
      setIsFocused,
      scrollToPositionImpl,
    });

    /**
     * Handle focus on container - redirect to hidden PM.
     */
    const handleContainerFocus = useCallback(
      (e: React.FocusEvent) => {
        if (readOnly) return;
        // Don't steal focus from sidebar inputs (textareas, inputs, buttons)
        const target = e.target as HTMLElement;
        if (target.closest('.docx-comments-sidebar') || target.closest('.docx-unified-sidebar'))
          return;
        // Don't steal focus from the hyperlink popup's text/URL inputs —
        // the focus event bubbles up here and would bounce focus back to
        // the body PM, making the inputs impossible to edit.
        if (target.closest('.oox-hyperlink-popup')) return;
        focusBodyInput();
        setIsFocused(true);
      },
      [readOnly, focusBodyInput]
    );

    /**
     * Handle blur from container.
     */
    const handleContainerBlur = useCallback((e: React.FocusEvent) => {
      // Check if focus is staying within the editor container.
      const relatedTarget = e.relatedTarget as HTMLElement | null;
      if (relatedTarget && containerRef.current?.contains(relatedTarget)) {
        return; // Focus staying within editor
      }
      // Keep selection visible when focus moves to toolbar or dropdown portals
      if (
        relatedTarget?.closest(
          '[role="toolbar"], [data-radix-popper-content-wrapper], [data-radix-select-content], .docx-table-options-dropdown'
        )
      ) {
        return;
      }
      setIsFocused(false);
    }, []);

    // Image overlay interactions — resize + drag-to-move. Owns the writes
    // to `isImageInteractingRef` that gate the selection hook's deferred
    // image-info clear during drag/resize gestures.
    const {
      handleImageResize,
      handleImageResizeStart,
      handleImageResizeEnd,
      handleImageDragMove,
      handleImageDragStart,
      handleImageDragEnd,
    } = useImageInteractions({
      pagesContainerRef,
      getPositionProjection: () => getYrsPositionProjection('body'),
      isImageInteractingRef,
      getPositionFromMouse,
      canvasHostRef,
      displayListQueries,
      applyYrsCommand,
    });

    /**
     * Handle keyboard events on container.
     * The hidden textarea handles native keyboard/IME input; this container
     * only intercepts shortcuts that originate outside it.
     */
    const handleKeyDown = useCallback(
      (e: React.KeyboardEvent) => {
        if (readOnly) return;
        // The hidden textarea owns every keyboard/IME event for both body and
        // header/footer roots. Do not re-interpret its bubbled events.
        if (yrsInputRef.current?.isFocused()) return;
        const target = e.target as HTMLElement | null;
        // Don't hijack keystrokes typed into the hyperlink popup's inputs —
        // refocusing the body PM here would steal focus mid-type and route
        // keys (e.g. space) into the document instead of the input.
        if (target?.closest('.oox-hyperlink-popup')) return;
        // Ensure the input surface is focused if the user types.
        if (!yrsInputRef.current?.isFocused()) {
          focusBodyInput();
          setIsFocused(true);
        }

        // PageUp/PageDown - let container handle scrolling
        if (['PageUp', 'PageDown'].includes(e.key) && !e.metaKey && !e.ctrlKey) {
          // Let PM handle the cursor movement first
          // If PM doesn't handle it (at bounds), the container will scroll
        }

        // Cmd/Ctrl+Home - scroll to top and move cursor to start
        if (e.key === 'Home' && (e.metaKey || e.ctrlKey)) {
          const sc = getScrollContainer();
          if (sc) sc.scrollTop = 0;
        }

        // Cmd/Ctrl+End - scroll to bottom and move cursor to end
        if (e.key === 'End' && (e.metaKey || e.ctrlKey)) {
          const sc = getScrollContainer();
          if (sc) sc.scrollTop = sc.scrollHeight;
        }
      },
      [readOnly, getScrollContainer, focusBodyInput]
    );

    /**
     * Handle mousedown on container (outside pages).
     */
    const handleContainerMouseDown = useCallback(
      (e: React.MouseEvent) => {
        if (readOnly) return;
        // Don't steal focus from sidebar inputs
        if (
          (e.target as HTMLElement).closest('.docx-comments-sidebar') ||
          (e.target as HTMLElement).closest('.docx-unified-sidebar')
        )
          return;
        // Focus the input surface if clicking outside pages area
        if (!yrsInputRef.current?.isFocused()) {
          focusBodyInput();
          setIsFocused(true);
        }
      },
      [readOnly, focusBodyInput]
    );

    // =========================================================================
    // Initial Layout
    // =========================================================================

    useEffect(() => {
      if (!yrsCore.session) return;
      runLayoutPipelineRef.current?.();
      updateSelectionOverlayRef.current();
      if (!readOnly) {
        const raf = requestAnimationFrame(() => {
          focusBodyInput();
          setIsFocused(true);
        });
        return () => cancelAnimationFrame(raf);
      }
    }, [focusBodyInput, readOnly, runLayoutPipelineRef, yrsCore.session]);

    // Canvas renderer: re-derive sidebar anchor Ys from the display list once
    // each build lands. The layout pipeline's pass writes Ys in the DOM
    // viewport's coordinate space (VIEWPORT_PADDING_TOP + pageGap), which
    // doesn't correspond to the canvas page stack; this effect overrides them
    // with `.canvas-pages`-host offsets derived from Rust range_rects
    // queries. No-op (and no listener churn) on the default DOM-painter path.
    //
    // Under direct yrs input, the session's sticky comment coverage and revision
    // ranges are projected to the body/HF display-list regions without touching
    // an editor view.
    useEffect(() => {
      const session = yrsCore.session;
      if (!session || !displayListQueries || !onAnchorPositionsChange) {
        return;
      }
      let cancelled = false;
      let hostRaf: number | null = null;
      const emit = (): void => {
        if (cancelled) return;
        const host = canvasHostRef?.current ?? pagesContainerRef.current;
        if (!host?.classList.contains('canvas-pages')) {
          hostRaf = requestAnimationFrame(emit);
          return;
        }
        const target = canvasOverlayTarget ?? host?.parentElement ?? null;
        if (!target) return;
        const targetRect = target.getBoundingClientRect();
        const projectY = (rect: DisplayListRect): number | null => {
          const pageRect = resolveDisplayPageClientRect(host, displayListQueries, rect.pageIndex);
          const pageSize = displayListQueries.pageSize(rect.pageIndex);
          if (!pageRect || !pageSize || pageSize.height <= 0) return null;
          return pageRect.top - targetRect.top + rect.y * (pageRect.height / pageSize.height);
        };
        const projection = createYrsSidebarProjection(session);
        const revisions = session.listRevisions();
        onYrsTrackedChangesChange?.(extractTrackedChangesFromYrs(revisions, projection));
        const hfRegions = new Map<string, 'header' | 'footer'>();
        for (const rId of document?.package?.headers?.keys() ?? []) {
          hfRegions.set(rId, 'header');
        }
        for (const rId of document?.package?.footers?.keys() ?? []) {
          if (!hfRegions.has(rId)) hfRegions.set(rId, 'footer');
        }
        onAnchorPositionsChange(
          computeAnchorPositionsFromYrs(
            session,
            sidebarCommentIds,
            revisions,
            projection,
            displayListQueries,
            hfRegions,
            projectY
          )
        );
      };
      emit();
      void displayListQueries.whenReady().then(emit, () => undefined);
      return () => {
        cancelled = true;
        if (hostRaf !== null) cancelAnimationFrame(hostRaf);
      };
    }, [
      canvasHostRef,
      canvasOverlayTarget,
      displayListQueries,
      document,
      onAnchorPositionsChange,
      onYrsTrackedChangesChange,
      pagesContainerRef,
      sidebarCommentIds,
      yrsCore.session,
    ]);

    // Canvas renderer (H2): re-back the plugin-facing RenderedDomContext with
    // the a11y mirror. The pipeline's painter-backed context resolves against
    // the parked 0×0 stage while the canvas paints, so this effect owns the
    // context on the canvas path (the pipeline skips its own emit when
    // source. The mirror speaks the same data-doc-* semantics
    // contract, so third-party plugins keep resolving geometry unchanged.
    useEffect(() => {
      if (!displayListQueries || !onRenderedDomContextReady) return;
      let cancelled = false;
      let hostRaf: number | null = null;
      const emit = (): void => {
        if (cancelled) return;
        const host = canvasHostRef?.current;
        if (!host?.classList.contains('canvas-pages')) {
          hostRaf = requestAnimationFrame(emit);
          return;
        }
        onRenderedDomContextReady(
          createRenderedDomContext(host, zoom, {
            displayListQueries,
            projector: createCanvasHostProjector(host, displayListQueries, zoom),
          })
        );
      };
      emit();
      return () => {
        cancelled = true;
        if (hostRaf !== null) cancelAnimationFrame(hostRaf);
      };
    }, [displayListQueries, onRenderedDomContextReady, canvasHostRef, zoom]);

    // Re-layout triggers: web-font load complete + header/footer content changes.
    useLayoutTriggers({
      runLayoutPipeline,
      updateSelectionOverlay,
      headerContent,
      footerContent,
      firstPageHeaderContent,
      firstPageFooterContent,
    });

    // Imperative-handle setup — exposes PagedEditorRef + mirrors via onReady.
    usePagedEditorRefApi({
      ref,
      yrsInputRef,
      layout,
      runLayoutPipeline,
      scrollToPositionImpl,
      scrollToParaIdImpl,
      scrollToPageImpl,
      setIsFocused,
      onReadyRef,
      documentFromYrs: yrsCore.documentFromYrs,
      yrsSession: yrsCore.session,
      yrsLocToDisplayPosition,
      syncYrsInputState,
      applyYrsFormatting,
      applyYrsCommand,
      getYrsPositionProjection: () => getYrsPositionProjection('body'),
      displayPositionToYrsLoc: (position) => {
        const target = getYrsPositionProjection('body')?.targetAt(position);
        return target ? yrsCore.displayPositionToLoc(target.displayPosition, target.story) : null;
      },
    });

    // =========================================================================
    // Render
    // =========================================================================

    // Min-height of the viewport wrapper. Delegates to `viewportMinHeightPx`
    // so the same math runs in both the JSX commit and the imperative write
    // the layout pipeline does mid-pipeline (needed for scroll-restore math
    // before React commits).
    const totalHeight = useMemo(() => {
      if (!layout) return DEFAULT_PAGE_HEIGHT_PX + VIEWPORT_PADDING_TOP + VIEWPORT_PADDING_BOTTOM;
      return viewportMinHeightPx(layout, pageGap);
    }, [layout, pageGap]);

    // Canvas path: the painted `.layout-run-image` the DOM overlay measures is
    // parked, so `selectedImageInfo` never populates. Derive the selected image
    // node's display position straight from the live selection instead — the
    // CanvasImageSelectionOverlay sources its rect from the display list. Both a
    // Both an image-node selection and a one-atom text range count as
    // "image selected". Re-derived every render; selection changes already
    // re-render PagedEditor (handleSelectionChange sets overlay state).
    const canvasSelectedImagePos = (() => {
      if (!canvasOverlayTarget || !displayListQueries) return null;
      const selection = yrsInputRef.current?.displaySelection();
      if (!selection || Math.abs(selection.anchor - selection.head) !== 1) return null;
      const from = Math.min(selection.anchor, selection.head);
      return getYrsPositionProjection('body')?.nodeAt(from)?.kind === 'image' ? from : null;
    })();
    const interactionPageHostRef =
      canvasHostRef?.current != null ? canvasHostRef : pagesContainerRef;

    return (
      <div
        ref={containerRef}
        className={`oox-root paged-editor ${className ?? ''}`}
        style={{ ...containerStyles, ...style, display: 'contents' }}
        tabIndex={0}
        onFocus={handleContainerFocus}
        onBlur={handleContainerBlur}
        onKeyDown={handleKeyDown}
        onMouseDown={handleContainerMouseDown}
      >
        {/* Hidden textarea input surface that writes directly to yrs. */}
        <YrsInput
          ref={yrsInputRef}
          enabled
          readOnly={readOnly || (!!hfEditMode && !hfEditRId)}
          session={yrsCore.session}
          story={activeYrsRootStory}
          isSuggesting={isSuggesting}
          author={author}
          inputPositionMap={yrsInputPositionMap}
          displayPositionToLoc={yrsDisplayPositionToLoc}
          locToDisplayPosition={yrsLocToDisplayPosition}
          nextParagraphStyleId={(styleId) => yrsStyleResolver?.getNextStyleId(styleId) ?? null}
          displayListQueries={activeYrsRootStory === 'body' ? displayListQueries : null}
          displayListFrameEpoch={displayListFrameEpoch}
          residentCaret={residentCaret}
          residentCaretAuthoritative={residentCaretAuthoritative}
          canvasHostRef={canvasHostRef ?? pagesContainerRef}
          onStateChange={handleYrsStateChange}
          onDirectInput={publishYrsDirectInput}
          applyResidentInput={applyResidentInput}
          applyResidentDelete={applyResidentDelete}
          onFocusChange={setIsFocused}
          onCaretInput={activeYrsRootStory === 'body' ? onCaretInput : undefined}
          onCaretInputDispatched={activeYrsRootStory === 'body' ? onCaretInputDispatched : undefined}
          onCaretInterrupt={onCaretInterrupt}
        />

        {/* Non-rendering orchestration host. Visible pages are canvas-only;
            this empty ref remains while legacy interaction signatures are
            collapsed in later retirement phases. */}
        <div ref={viewportLayoutRef} style={{ display: 'contents' }}>
          <div
            ref={pagesContainerRef}
            className="paged-editor__layout-host"
            style={{ display: 'none' }}
            aria-hidden="true"
          />

          {canvasOverlayTarget &&
            canvasDisplayList &&
            (residentCaretAuthoritative || displayListQueries) && (
              <CanvasSelectionOverlay
                selectionRects={hfEditMode ? [] : selectionRects}
                // While the worker paints the caret into the presented frame,
                // the DOM blink caret stays unmounted (two-mode caret).
                caretPosition={hfEditMode || paintedCaretActive ? null : caretPosition}
                isFocused={isFocused && !hfEditMode}
                readOnly={readOnly}
                overlayTarget={canvasOverlayTarget}
                canvasHostRef={interactionPageHostRef}
                displayList={canvasDisplayList}
                displayListQueries={displayListQueries}
                directProjection={residentCaretAuthoritative}
                sidebarOpen={commentsSidebarOpen}
                zoom={zoom}
              />
            )}

          {canvasOverlayTarget && displayListQueries && !hfEditMode && (
            <CanvasCellSelectionOverlay
              session={yrsCore.session}
              positionProjection={getYrsPositionProjection('body')}
              overlayTarget={canvasOverlayTarget}
              canvasHostRef={interactionPageHostRef}
              displayListQueries={displayListQueries}
              sidebarOpen={commentsSidebarOpen}
              zoom={zoom}
            />
          )}

          {canvasOverlayTarget && displayListQueries && (
            <CanvasImageSelectionOverlay
              pmPos={hfEditMode ? null : canvasSelectedImagePos}
              isFocused={isFocused && !hfEditMode}
              readOnly={readOnly}
              overlayTarget={canvasOverlayTarget}
              canvasHostRef={interactionPageHostRef}
              displayListQueries={displayListQueries}
              sidebarOpen={commentsSidebarOpen}
              zoom={zoom}
              onResize={handleImageResize}
              onResizeStart={handleImageResizeStart}
              onResizeEnd={handleImageResizeEnd}
              onDragMove={handleImageDragMove}
              onDragStart={handleImageDragStart}
              onDragEnd={handleImageDragEnd}
              onContextMenu={handlePagesContextMenu}
            />
          )}

          {canvasOverlayTarget && displayListQueries && !hfEditMode && (
            <CanvasTableResizeOverlay
              overlayTarget={canvasOverlayTarget}
              canvasHostRef={interactionPageHostRef}
              displayListQueries={displayListQueries}
              positionProjection={getYrsPositionProjection('body')}
              applyYrsCommand={applyYrsCommand}
              readOnly={readOnly}
              sidebarOpen={commentsSidebarOpen}
              zoom={zoom}
            />
          )}

          {/* Canvas-path scrollToParaId flash (G2) — painter is parked, so the
              flash is drawn over the paragraph's display-list geometry. */}
          {canvasOverlayTarget && displayListQueries && (
            <CanvasParagraphFlashOverlay
              request={canvasFlashRequest}
              overlayTarget={canvasOverlayTarget}
              canvasHostRef={interactionPageHostRef}
              displayListQueries={displayListQueries}
              sidebarOpen={commentsSidebarOpen}
              zoom={zoom}
              onDone={handleCanvasFlashDone}
            />
          )}

          {/* Table quick action insert button */}
          {tableInsertButton &&
            (() => {
              const button = (
                <TableInsertButton
                  type={tableInsertButton.type}
                  x={tableInsertButton.x}
                  y={tableInsertButton.y}
                  onMouseDown={handleTableInsertClick}
                  onMouseEnter={clearTableInsertTimer}
                  onMouseLeave={hideTableInsertButton}
                />
              );
              return canvasOverlayTarget ? createPortal(button, canvasOverlayTarget) : button;
            })()}

          {/* Plugin overlays (highlights, annotations) */}
          {pluginOverlays &&
            (() => {
              const overlay = (
                <div className="paged-editor__plugin-overlays" style={pluginOverlaysStyles}>
                  {pluginOverlays}
                </div>
              );
              return canvasOverlayTarget ? createPortal(overlay, canvasOverlayTarget) : overlay;
            })()}
        </div>

        {/* Sidebar overlay — positioned to match visual document height, visible overflow for sidebar items */}
        {sidebarOverlay && (
          <div
            style={{
              position: 'absolute',
              top: 0,
              left: 0,
              right: 0,
              height: totalHeight * zoom,
              pointerEvents: 'none',
              overflow: 'visible',
            }}
          >
            <div style={{ pointerEvents: 'auto' }}>{sidebarOverlay}</div>
          </div>
        )}

        {/* Hyperlink popup — rendered inside containerRef so it shares a
            scroll context with the link. position: absolute + coords in
            container space mean the browser repositions on scroll for free. */}
        {hyperlinkPopupData &&
          onHyperlinkPopupNavigate &&
          onHyperlinkPopupCopy &&
          onHyperlinkPopupEdit &&
          onHyperlinkPopupRemove &&
          onHyperlinkPopupClose &&
          (() => {
            const popup = (
              <HyperlinkPopup
                data={hyperlinkPopupData}
                onNavigate={onHyperlinkPopupNavigate}
                onCopy={onHyperlinkPopupCopy}
                onEdit={onHyperlinkPopupEdit}
                onRemove={onHyperlinkPopupRemove}
                onClose={onHyperlinkPopupClose}
                readOnly={readOnly}
              />
            );
            return canvasOverlayTarget ? createPortal(popup, canvasOverlayTarget) : popup;
          })()}
      </div>
    );
  }
);

export const PagedEditor = memo(PagedEditorComponent);

export default PagedEditor;
