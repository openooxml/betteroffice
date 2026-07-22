/**
 * DocxEditor Component
 *
 * Main component integrating all editor features:
 * - Toolbar for formatting
 * - Yrs-backed editing and canvas rendering
 * - Zoom control
 * - Error boundary
 * - Loading states
 */

import { useRef, useCallback, useState, useEffect, useMemo, forwardRef } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { Document, Theme } from '@betteroffice/docx/types/document';
import type { YrsLoc, YrsSession, YrsStoryRange } from '@betteroffice/docx/yrs';
import type { BundledFontProvider } from '@betteroffice/docx/layout';
import {
  createYrsSidebarProjection,
  extractTrackedChangesFromYrs,
  yrsIdToNumericId,
  type TrackedChangesResult,
} from '@betteroffice/docx/layout/render';

import { cn } from '../lib/utils';
import { type SelectionFormatting } from './Toolbar';
import type {
  DocxEditorCollaborationOptions,
  SelectionState,
  TableContextInfo,
} from './DocxEditor/types';
import { useOutlineSidebar } from './DocxEditor/hooks/useOutlineSidebar';
import { useKeyboardShortcuts } from './DocxEditor/hooks/useKeyboardShortcuts';
import { useFileIO } from './DocxEditor/hooks/useFileIO';
import { usePageSetupControls } from './DocxEditor/hooks/usePageSetupControls';
import { useWatermarkControls } from './DocxEditor/hooks/useWatermarkControls';
import { useHyperlinkActions } from './DocxEditor/hooks/useHyperlinkActions';
import { useFindReplaceBridge, type YrsFindMatch } from './DocxEditor/hooks/useFindReplaceBridge';
import { CanvasFindHighlightOverlay } from './DocxEditor/overlays/CanvasFindHighlightOverlay';
import {
  CanvasSidebarBrightenOverlay,
  type CanvasBrightenRange,
} from './DocxEditor/overlays/CanvasSidebarBrightenOverlay';
import { useCanvasOverlayTarget } from './DocxEditor/internals/useCanvasOverlayTarget';
import { isWithinPageArea } from './DocxEditor/internals/pageAreaRouting';
import { useFormattingActions } from './DocxEditor/hooks/useFormattingActions';
import { useImageActions } from './DocxEditor/hooks/useImageActions';
import { useDocxEditorRefApi } from './DocxEditor/hooks/useDocxEditorRefApi';
import { useControllableBoolean } from './DocxEditor/hooks/useControllableBoolean';
import { useTableDialogs } from './DocxEditor/hooks/useTableDialogs';
import { useHeaderFooterEditing } from './DocxEditor/hooks/useHeaderFooterEditing';
import { useDocumentLoader } from './DocxEditor/hooks/useDocumentLoader';
import { useContextMenus } from './DocxEditor/hooks/useContextMenus';
import { useCommentManagement } from './DocxEditor/hooks/useCommentManagement';
import { useCommentLifecycle } from './DocxEditor/hooks/useCommentLifecycle';
import {
  useSelectionTracker,
  type SelectionStateDelta,
} from './DocxEditor/hooks/useSelectionTracker';
import { useFloatingCommentBtn } from './DocxEditor/hooks/useFloatingCommentBtn';
import { useActiveEditor } from './DocxEditor/hooks/useActiveEditor';
import { useScrollPageInfo } from './DocxEditor/hooks/useScrollPageInfo';
import { DocxEditorOverlays } from './DocxEditor/DocxEditorOverlays';
import { DocxEditorDialogs } from './DocxEditor/DocxEditorDialogs';
import { DocxEditorToolbar } from './DocxEditor/DocxEditorToolbar';
import { DocxEditorPagedArea } from './DocxEditor/DocxEditorPagedArea';
import { ContentControlWidgets } from './DocxEditor/ContentControlWidgets';
import { CanvasPagedArea } from './DocxEditor/CanvasPagesView';
import { useCanvasRenderer } from './DocxEditor/hooks/useDisplayList';
import type { RustFontChainsProvider } from './DocxEditor/hooks/useRustMeasurement';
import { useResetEditorState } from './DocxEditor/hooks/useResetEditorState';
import type { YrsToolbarSelection } from './DocxEditor/yrsToolbar';
import { DocxEditorShell } from './DocxEditor/DocxEditorShell';
import type { FontOption } from './ui/FontPicker';
import { OUTLINE_BUTTON_RESERVED_SPACE, OUTLINE_RESERVED_SPACE } from './DocumentOutline';
import { RULER_WIDTH } from './ui/VerticalRuler';
import { SIDEBAR_DOCUMENT_SHIFT } from './sidebar/constants';
import { useCommentSidebarItems, type CommentCallbacks } from '../hooks/useCommentSidebarItems';
import type { ReactSidebarItem } from '../plugin-api/types';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Translations } from '@betteroffice/docx-i18n';
import { type PrintOptions } from './ui/PrintPreview';
// Dialog hooks and utilities (static imports — lightweight, no UI)
import { useFindReplace } from './dialogs/FindReplaceDialog';
import { useHyperlinkDialog } from './dialogs/HyperlinkDialog';
import { DefaultLoadingIndicator, DefaultPlaceholder, ParseError } from './DocxEditorHelpers';
import { type DocxInput } from '@betteroffice/docx/utils';
import type { FontDefinition, ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import { useFontLifecycle } from '../hooks/useFontLifecycle';
import { useTableSelection } from '../hooks/useTableSelection';
import { useDocumentHistory } from '../hooks/useHistory';

import { createStyleResolver } from '@betteroffice/docx/styles';
import { useIsDark } from './DocxEditor/hooks/useIsDark';

// Paginated editor
import { type PagedEditorRef, DEFAULT_PAGE_WIDTH } from './DocxEditor/PagedEditor';

// Plugin API types
import type { RenderedDomContext } from '../plugin-api/types';

// ============================================================================
// TYPES
// ============================================================================

export type { DocxEditorCollaborationOptions } from './DocxEditor/types';

/**
 * DocxEditor props
 */
export interface DocxEditorProps {
  /** Document data — ArrayBuffer, Uint8Array, Blob, or File */
  documentBuffer?: DocxInput | null;
  /** Pre-parsed document (alternative to documentBuffer) */
  document?: Document | null;
  /** Callback when document is saved */
  onSave?: (buffer: ArrayBuffer) => void;
  /** Configure the Yrs collaboration replica used by the editor. */
  collaboration?: DocxEditorCollaborationOptions;
  /**
   * Callback when a DOCX file is selected through `File > Open` or Cmd/Ctrl+O.
   * Pass it to route the picked file through your own import pipeline. Omit it
   * to keep the built-in local document load behavior.
   */
  onOpen?: (file: File) => void | Promise<void>;
  /** Author name used for comments and track changes */
  author?: string;
  /** Callback when document changes */
  onChange?: (document: Document) => void;
  /** Callback when selection changes */
  onSelectionChange?: (state: SelectionState | null) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
  /** Callback when fonts are loaded */
  onFontsLoaded?: () => void;
  /** Color theme mode for UI styling. `'system'` follows the OS preference. */
  colorMode?: 'light' | 'dark' | 'system';
  /** Document theme schema object */
  theme?: Theme | null;
  /** Whether to show toolbar (default: true) */
  showToolbar?: boolean;
  /**
   * Whether to show `File > Open` and enable Cmd/Ctrl+O (default: true).
   * Set false when you provide your own open action elsewhere.
   */
  showFileOpen?: boolean;
  /** Whether to show the Help menu in the menu bar (default: true) */
  showHelpMenu?: boolean;
  /** Whether to show zoom control (default: true) */
  showZoomControl?: boolean;
  /** Whether to show page margin guides/boundaries (default: false) */
  showMarginGuides?: boolean;
  /** Color for margin guides (default: '#c0c0c0') */
  marginGuideColor?: string;
  /** Whether to show horizontal ruler (default: false) */
  showRuler?: boolean;
  /** Unit for ruler display (default: 'inch') */
  rulerUnit?: 'inch' | 'cm';
  /** Initial zoom level (default: 1.0) */
  initialZoom?: number;
  /** Whether the editor is read-only. When true, hides toolbar and rulers */
  readOnly?: boolean;
  /**
   * When true, the editor does not intercept Cmd/Ctrl+F or Cmd/Ctrl+H.
   * This lets the browser or host app handle native find/history shortcuts.
   */
  disableFindReplaceShortcuts?: boolean;
  /** Custom toolbar actions */
  toolbarExtra?: ReactNode;
  /** Additional CSS class name */
  className?: string;
  /** Additional inline styles */
  style?: CSSProperties;
  /** Placeholder when no document */
  placeholder?: ReactNode;
  /** Loading indicator */
  loadingIndicator?: ReactNode;
  /** Whether to show the document outline sidebar (default: false) */
  showOutline?: boolean;
  /** Whether to show the floating outline toggle button (default: true) */
  showOutlineButton?: boolean;
  /**
   * Custom list of fonts shown in the toolbar's font-family dropdown.
   * Strings render in the "Other" group; pass `FontOption[]` for category
   * grouping and CSS fallback chains. Omit to use the built-in 12-font
   * default. An empty array renders an empty (but enabled) dropdown.
   *
   * Pass a stable reference (memoized or module-level) — inline arrays
   * create a new identity per render and invalidate the picker's memo.
   *
   * @example fontFamilies={['Arial', 'Roboto']}
   * @example fontFamilies={[{ name: 'Roboto', fontFamily: 'Roboto, sans-serif', category: 'sans-serif' }]}
   */
  fontFamilies?: ReadonlyArray<string | FontOption>;
  /**
   * Custom font faces to register with the browser before the editor measures
   * text. Each entry injects an `@font-face` rule. Pass a URL (woff2/woff/
   * ttf/otf), an ArrayBuffer, or omit `src` to load by name from Google Fonts.
   * Multiple entries can share `family` to register different weights/styles.
   *
   * Pass a stable reference — inline arrays re-register faces on each render
   * (the loader dedupes by `family|weight|style`, so it's harmless but wastes
   * work).
   *
   * @example
   * fonts={[
   *   { family: 'Custom Sans', src: '/fonts/CustomSans-Regular.woff2' },
   *   { family: 'Custom Sans', src: '/fonts/CustomSans-Bold.woff2', weight: 700 },
   * ]}
   */
  fonts?: ReadonlyArray<FontDefinition>;
  /**
   * Text-watermark presets shown in the watermark dialog's preset dropdown.
   * Omit to use the built-in MS Word phrases (`DEFAULT_WATERMARK_PRESETS`:
   * CONFIDENTIAL, DRAFT, DO NOT COPY, SAMPLE, URGENT, ASAP). Pass an empty
   * array to hide the preset dropdown and require custom text.
   *
   * @example watermarkPresets={['INTERNAL', 'PROPRIETARY', 'COPY']}
   */
  watermarkPresets?: readonly string[];
  /** Print options for print preview */
  printOptions?: PrintOptions;
  /**
   * Callback when print is triggered. Pass it to enable the `File > Print`
   * menu entry; omit to hide. The imperative `ref.current.print()` also
   * invokes this callback.
   */
  onPrint?: () => void;
  /** Callback when content is copied */
  onCopy?: () => void;
  /** Callback when content is cut */
  onCut?: () => void;
  /** Callback when content is pasted */
  onPaste?: () => void;
  /** Editor mode: 'editing' (direct edits), 'suggesting' (track changes), or 'viewing' (read-only). Default: 'editing' */
  mode?: EditorMode;
  /** Callback when the editing mode changes */
  onModeChange?: (mode: EditorMode) => void;
  /** Callback when a comment is added via the UI */
  onCommentAdd?: (comment: Comment) => void;
  /** Callback when a comment is resolved via the UI */
  onCommentResolve?: (comment: Comment) => void;
  /** Callback when a comment is deleted via the UI */
  onCommentDelete?: (comment: Comment) => void;
  /** Callback when a reply is added to a comment via the UI */
  onCommentReply?: (reply: Comment, parent: Comment) => void;
  /**
   * Controlled comments array. When provided, the editor reads comment thread
   * metadata (text, author, replies, resolved status) from this prop instead
   * of internal state, and emits every change through `onCommentsChange`.
   *
   * Use this with collaboration backends (Yjs, Liveblocks, Automerge, …) so
   * comment threads sync across peers — the PM document only carries the
   * range markers; thread metadata lives outside the doc and needs its own
   * sync channel.
   *
   * If omitted, the editor falls back to internal state (current behavior).
   * The granular `onCommentAdd`/`onCommentResolve`/`onCommentDelete`/
   * `onCommentReply` callbacks fire in both modes.
   */
  comments?: Comment[];
  /** Fires whenever the comments array changes (controlled mode). */
  onCommentsChange?: (comments: Comment[]) => void;
  /** Controlled comments-sidebar visibility; source of truth when set. Pair with `onCommentsSidebarOpenChange`; omit for the default self-managed behavior. */
  commentsSidebarOpen?: boolean;
  /** Fires with the next open state whenever the editor wants to show or hide the comments sidebar. Fires in both controlled and uncontrolled modes. */
  onCommentsSidebarOpenChange?: (open: boolean) => void;
  /**
   * Callback when rendered DOM context is ready (for plugin overlays).
   * Used by PluginHost to get access to the rendered page DOM for positioning.
   */
  onRenderedDomContextReady?: (context: RenderedDomContext) => void;
  /**
   * Plugin overlays to render inside the editor viewport.
   * Passed from PluginHost to render plugin-specific overlays.
   */
  pluginOverlays?: ReactNode;
  /** Sidebar items from plugins (passed from PluginHost). */
  pluginSidebarItems?: ReactSidebarItem[];
  /** Rendered DOM context from PluginHost (for sidebar position resolution). */
  pluginRenderedDomContext?: RenderedDomContext | null;
  /** Custom logo/icon for the title bar */
  renderLogo?: () => ReactNode;
  /** Document name shown in the title bar */
  documentName?: string;
  /** Callback when document name changes */
  onDocumentNameChange?: (name: string) => void;
  /** Whether the document name is editable (default: true) */
  documentNameEditable?: boolean;
  /** Custom right-side actions for the title bar */
  renderTitleBarRight?: () => ReactNode;
  /** Translation overrides. Import a locale JSON file and pass it directly. */
  i18n?: Translations;
  /**
   * Bundled metric-compatible font bytes (e.g. Carlito for Calibri) for
   * families the document does not embed. Injected as an interface so the
   * editor never depends on a font bundle — build one from
   * `@betteroffice/docx-fonts` or any custom source. Without it, families with
   * no embedded bytes measure with deterministic synthetic metrics — wire a
   * provider with `resolveLastResort` for faithful output.
   */
  measurementFontProvider?: BundledFontProvider;
}

/**
 * DocxEditor ref interface
 */
export interface DocxEditorRef {
  /** Get the current document */
  getDocument: () => Document | null;
  /** Get the editor ref */
  getEditorRef: () => PagedEditorRef | null;
  /** Save the document to a buffer. */
  save: () => Promise<ArrayBuffer | null>;
  /** Set zoom level */
  setZoom: (zoom: number) => void;
  /** Get current zoom level */
  getZoom: () => number;
  /** Focus the editor */
  focus: () => void;
  /** Get current page number */
  getCurrentPage: () => number;
  /** Get total page count */
  getTotalPages: () => number;
  /**
   * Scroll the paginated view so the given page is in view.
   * Page numbers are 1-indexed (matches `getCurrentPage` / `getTotalPages`).
   * No-op for out-of-range or non-integer values.
   * @example ref.current?.scrollToPage(2)
   */
  scrollToPage: (pageNumber: number) => void;
  /**
   * Scroll the paginated view to the paragraph with the given Word `w14:paraId`.
   * Pass `options.highlight` to briefly flash it in a custom color.
   * @returns whether a matching paragraph exists in the live document
   * @example ref.current?.scrollToParaId('1A2B3C4D', { highlight: { color: 'rgba(255, 235, 59, 0.55)' } })
   */
  scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  /**
   * Scroll the paginated view to a specific display position.
   * For Word `w14:paraId` use
   * `scrollToParaId` instead.
   * @example ref.current?.scrollToPosition(42)
   */
  scrollToPosition: (displayPosition: number) => void;
  /**
   * Scroll the paginated view to the comment with the given id and select its
   * anchored range so the selection overlay highlights it. Resolves the id
   * against the live comment marks at call time.
   * @returns `false` when the id no longer resolves (the comment was deleted
   *   or its anchored text removed between render and click), so the caller
   *   can surface a "location no longer exists" affordance rather than
   *   silently no-op'ing.
   * @example ref.current?.scrollToCommentId(3)
   */
  scrollToCommentId: (commentId: number) => boolean;
  /**
   * Scroll the paginated view to the tracked change with the given Word
   * revision `w:id` and select its range so the selection overlay highlights
   * it. Resolves the id against the live tracked-change marks at call time
   * (matching coalesced revisions the way the changes sidebar does).
   * @returns `false` when the id no longer resolves (the change was
   *   accepted, rejected, or deleted between render and click).
   * @example ref.current?.scrollToChangeId(42)
   */
  scrollToChangeId: (revisionId: number) => boolean;
  /**
   * Select the display-position range `[from, to]` so the selection
   * overlay highlights it, and scroll its start into view. The selection
   * persists until it next changes (there is no auto-clearing flash). No-op
   * for a malformed range or a `from` past the document end; `to` is clamped
   * to the document size.
   * @example ref.current?.highlightRange(10, 24)
   */
  highlightRange: (from: number, to: number) => void;
  /** Open print preview */
  openPrintPreview: () => void;
  /** Print the document directly */
  print: () => void;
  /** Load a pre-parsed document programmatically */
  loadDocument: (doc: Document) => void;
  /** Load a DOCX buffer programmatically (ArrayBuffer, Uint8Array, Blob, or File) */
  loadDocumentBuffer: (buffer: DocxInput) => Promise<void>;
  /** Add a comment programmatically. Anchored by Word `w14:paraId` so
   * it survives unrelated edits. Returns the comment ID, or null if
   * the paraId is unknown or the search text isn't found / is ambiguous. */
  addComment: (options: {
    paraId: string;
    text: string;
    author: string;
    /** Optional: anchor to a specific phrase within the paragraph (must be unique). */
    search?: string;
  }) => number | null;
  /** Reply to an existing comment. Returns the reply comment ID. */
  replyToComment: (commentId: number, text: string, author: string) => number | null;
  /** Resolve (mark as done) a comment. */
  resolveComment: (commentId: number) => void;
  /** Suggest a tracked change. Pass `replaceWith: ''` to delete the matched text;
   * pass `search: ''` to insert at paragraph end. Returns false on missing paraId,
   * missing/ambiguous search, or attempt to layer on an existing tracked change. */
  proposeChange: (options: {
    paraId: string;
    search: string;
    replaceWith: string;
    author: string;
  }) => boolean;
  /** Locate every paragraph containing `query` (case-insensitive substring).
   * Returns a stable handle (paraId + the matched phrase) the agent can pass
   * back to `addComment` / `proposeChange`. */
  findInDocument: (
    query: string,
    options?: { caseSensitive?: boolean; limit?: number }
  ) => Array<{ paraId: string; match: string; before: string; after: string }>;
  /**
   * Apply character formatting (bold / italic / color / size / font / etc.)
   * to a paragraph or to a unique phrase within it. This is a direct edit,
   * not a tracked change. Returns false on missing paraId or ambiguous search.
   */
  applyFormatting: (options: {
    paraId: string;
    search?: string;
    marks: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean | { style?: string };
      strike?: boolean;
      color?: { rgb?: string; themeColor?: string };
      highlight?: string;
      fontSize?: number;
      fontFamily?: { ascii?: string; hAnsi?: string };
    };
  }) => boolean;
  /**
   * Apply a paragraph style by styleId (e.g. `'Heading1'`, `'Quote'`).
   * Direct edit, not a tracked change. Returns false if paraId is unknown.
   */
  setParagraphStyle: (options: { paraId: string; styleId: string }) => boolean;
  /**
   * Insert a page or section break after the paragraph identified by `paraId`.
   * `'page'` adds a page break; `'sectionNextPage'` / `'sectionContinuous'`
   * start a new section on a new page / the same page. Direct edit, not a
   * tracked change. Returns false if paraId is unknown.
   */
  insertBreak: (options: {
    paraId: string;
    type: 'page' | 'sectionNextPage' | 'sectionContinuous';
  }) => boolean;
  /**
   * Read the contents of a single page. 1-indexed; returns null if the page
   * does not exist. Each paragraph is returned with its stable paraId so the
   * agent can comment on or modify it without an extra round-trip.
   */
  getPageContent: (pageNumber: number) => {
    pageNumber: number;
    text: string;
    paragraphs: Array<{ paraId: string; text: string; styleId?: string }>;
  } | null;
  /** Read the user's current cursor / selection — what's highlighted right now. */
  getSelectionInfo: () => {
    paraId: string | null;
    selectedText: string;
    paragraphText: string;
    before: string;
    after: string;
  } | null;
  /** Get all comments. */
  getComments: () => Comment[];
  /** Subscribe to document changes. Fires after every committed edit. Returns unsubscribe. */
  onContentChange: (listener: (document: Document) => void) => () => void;
  /** Subscribe to selection changes (cursor moves / selection changes). Returns unsubscribe. */
  onSelectionChange: (listener: (selection: SelectionState | null) => void) => () => void;
}

/**
 * Editor internal state
 */
interface EditorState {
  isLoading: boolean;
  parseError: string | null;
  zoom: number;
  /** Current selection formatting for toolbar */
  selectionFormatting: SelectionFormatting;
  /** Paragraph indent data for ruler */
  paragraphIndentLeft: number;
  paragraphIndentRight: number;
  paragraphFirstLineIndent: number;
  paragraphHangingIndent: boolean;
  paragraphTabs: import('@betteroffice/docx/types/document').TabStop[] | null;
  /** Table context for showing the table toolbar. */
  pmTableContext: TableContextInfo | null;
  /** Image context when cursor is on an image node */
  pmImageContext: {
    pos: number;
    wrapType: string;
    displayMode: string;
    cssFloat: string | null;
    transform: string | null;
    alt: string | null;
    borderWidth: number | null;
    borderColor: string | null;
    borderStyle: string | null;
    width: number | null;
    height: number | null;
  } | null;
}

export type { EditorMode } from './DocxEditor/internals/editing-modes';
import type { EditorMode } from './DocxEditor/internals/editing-modes';

function displayRangeToYrsRange(
  editor: PagedEditorRef,
  from: number,
  to: number
): YrsStoryRange | null {
  const start = editor.displayPositionToYrsLoc(from);
  const end = editor.displayPositionToYrsLoc(to);
  if (!start || !end || start.story !== end.story) return null;
  return {
    story: start.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
}

function yrsStoryOffset(session: YrsSession, loc: YrsLoc): number {
  return session.locateParagraph(loc.story, loc.paraId).start + loc.offset;
}

// ============================================================================
// MAIN COMPONENT
// ============================================================================

// `injectReplyRangeMarkers` + `injectTCReplyRangeMarkers` live in
// `@betteroffice/docx/docx` so React + Vue share the same
// pre-serialization range-marker injection.

import { getInitialSectionProperties } from './DocxEditor/internals/documentSetup';
import {
  EMPTY_ANCHOR_POSITIONS,
  createComment,
  createCommentIdAllocator,
} from './DocxEditor/commentFactories';

/**
 * DocxEditor - Complete DOCX editor component
 */
export const DocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function DocxEditor(
  {
    documentBuffer,
    document: initialDocument,
    onSave,
    collaboration,
    onOpen,
    author = 'User',
    onChange,
    onSelectionChange,
    onError,
    onFontsLoaded: onFontsLoadedCallback,
    colorMode = 'light',
    theme,
    showToolbar = true,
    showFileOpen = true,
    showHelpMenu = true,
    showZoomControl = true,
    showMarginGuides: _showMarginGuides = false,
    marginGuideColor: _marginGuideColor,
    showRuler = false,
    rulerUnit = 'inch',
    initialZoom = 1.0,
    readOnly: readOnlyProp = false,
    disableFindReplaceShortcuts = false,
    toolbarExtra,
    className = '',
    style,
    placeholder,
    loadingIndicator,
    showOutline: showOutlineProp = false,
    showOutlineButton = true,
    fontFamilies,
    fonts,
    watermarkPresets,
    printOptions: _printOptions,
    onPrint,
    onCopy: _onCopy,
    onCut: _onCut,
    onPaste: _onPaste,
    mode: modeProp,
    onModeChange,
    onCommentAdd,
    onCommentResolve,
    onCommentDelete,
    onCommentReply,
    comments: commentsProp,
    onCommentsChange,
    commentsSidebarOpen,
    onCommentsSidebarOpenChange,
    onRenderedDomContextReady,
    pluginOverlays,
    pluginSidebarItems,
    pluginRenderedDomContext,
    renderLogo,
    documentName,
    onDocumentNameChange,
    documentNameEditable = true,
    renderTitleBarRight,
    i18n,
    measurementFontProvider,
  },
  ref
) {
  // Host slot the Rust measure source (mounted deep in PagedEditor) fills with
  // the merged doc-wide font chains; the canvas display-list build reads it to
  // gate GlyphRun emission. Null until Rust measurement warms its first chains.
  const rustFontChainsProviderRef = useRef<RustFontChainsProvider | null>(null);
  // Assigned by CanvasA11yLiveRegion, called by useSelectionTracker.
  const canvasA11yNotifyRef = useRef<(() => void) | null>(null);

  // State
  const [state, setState] = useState<EditorState>({
    isLoading: !!documentBuffer,
    parseError: null,
    zoom: initialZoom,
    selectionFormatting: {},
    paragraphIndentLeft: 0,
    paragraphIndentRight: 0,
    paragraphFirstLineIndent: 0,
    paragraphHangingIndent: false,
    paragraphTabs: null,
    pmTableContext: null,
    pmImageContext: null,
  });

  const isDark = useIsDark(colorMode);

  // Header/footer editing state.
  const [hfEditPosition, setHfEditPosition] = useState<'header' | 'footer' | null>(null);
  const [hfEditIsFirstPage, setHfEditIsFirstPage] = useState(false);
  const [hfEditPageIndex, setHfEditPageIndex] = useState(0);

  // Controlled by `commentsSidebarOpen` when provided, else editor-owned; the
  // setter routes through `onCommentsSidebarOpenChange`. See useControllableBoolean.
  const [showCommentsSidebar, setShowCommentsSidebar] = useControllableBoolean(
    commentsSidebarOpen,
    onCommentsSidebarOpenChange
  );
  // Auto-open the sidebar the first time a comment / tracked change
  // appears so users see the card without manually toggling. Latches so
  // a subsequent close stays closed; reset on doc reload.
  const sidebarAutoOpenedRef = useRef(false);
  const [expandedSidebarItem, setExpandedSidebarItem] = useState<string | null>(null);
  // PagedEditor ref declared early so comment management can read the live
  // Yrs session before the tracked-changes effect drives `setComments`.
  const pagedEditorRef = useRef<PagedEditorRef>(null);

  const {
    comments,
    setComments,
    isAddingComment,
    setIsAddingComment,
    isAddingCommentRef,
    commentSelectionRange,
    setCommentSelectionRange,
    addCommentYPosition,
    setAddCommentYPosition,
    floatingCommentBtn,
    setFloatingCommentBtn,
    cleanOrphanedCommentsTimerRef,
    cleanOrphanedComments,
  } = useCommentManagement({
    commentsProp,
    onCommentDelete,
    onCommentsChange,
    pagedEditorRef,
  });

  const resolvedCommentIds = useMemo(() => {
    const ids = new Set<number>();
    for (const c of comments) {
      if (c.done && c.parentId == null) ids.add(c.id);
    }
    return ids;
  }, [comments]);

  // Exclude expanded resolved comment from hide-set so its text gets highlighted
  const resolvedIdsForRender = useMemo(() => {
    if (!expandedSidebarItem?.startsWith('comment-')) return resolvedCommentIds;
    const expandedId = parseInt(expandedSidebarItem.slice(8), 10);
    if (isNaN(expandedId) || !resolvedCommentIds.has(expandedId)) return resolvedCommentIds;
    const ids = new Set(resolvedCommentIds);
    ids.delete(expandedId);
    return ids;
  }, [resolvedCommentIds, expandedSidebarItem]);

  // Canvas renderer plumbing. `resolvedIdsForRender` reaches the Rust
  // display-list build so the canvas drops the comment wash of resolved
  // threads (and re-tints the one whose sidebar card is expanded).
  const canvasRenderer = useCanvasRenderer(rustFontChainsProviderRef, resolvedIdsForRender);
  useEffect(() => {
    if (canvasRenderer.error) onError?.(canvasRenderer.error);
  }, [canvasRenderer.error, onError]);

  const [yrsTrackedChangesResult, setYrsTrackedChangesResult] = useState<TrackedChangesResult>(
    () => ({
      entries: [],
      commentToRevision: new Map(),
    })
  );

  const { entries: trackedChanges, commentToRevision } = yrsTrackedChangesResult;

  const [anchorPositions, setAnchorPositions] =
    useState<Map<string, number>>(EMPTY_ANCHOR_POSITIONS);
  // No separate state needed — pluginRenderedDomContext comes from PluginHost

  const [editingModeInternal, setEditingModeInternal] = useState<EditorMode>(modeProp ?? 'editing');
  const editingMode = modeProp ?? editingModeInternal;
  const setEditingMode = (mode: EditorMode) => {
    if (!modeProp) setEditingModeInternal(mode);
    onModeChange?.(mode);
  };
  // 'viewing' mode acts as read-only
  const readOnly = readOnlyProp || editingMode === 'viewing';

  // Bridge / agent event subscribers — fan-out from the existing onChange and
  // onSelectionChange paths so multiple listeners (host app, MCP server, etc.)
  // can observe edits without competing for the single React prop.
  const contentChangeSubscribersRef = useRef(new Set<(doc: Document) => void>());
  const selectionChangeSubscribersRef = useRef(new Set<(s: SelectionState | null) => void>());

  // History hook for undo/redo - start with null document
  const history = useDocumentHistory<Document | null>(initialDocument || null, {
    maxEntries: 100,
    groupingInterval: 500,
    enableKeyboardShortcuts: true,
  });
  const [yrsHistoryState, setYrsHistoryState] = useState({ canUndo: false, canRedo: false });
  const handleYrsHistoryChange = useCallback((canUndo: boolean, canRedo: boolean) => {
    setYrsHistoryState((previous) =>
      previous.canUndo === canUndo && previous.canRedo === canRedo
        ? previous
        : { canUndo, canRedo }
    );
  }, []);

  // Refs (pagedEditorRef is declared earlier — useCommentManagement needs it)
  const containerRef = useRef<HTMLDivElement>(null);
  const editorContentRef = useRef<HTMLDivElement>(null);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const [documentFonts, setDocumentFonts] = useState<FontOption[]>([]);
  const {
    showOutline,
    setShowOutline,
    showOutlineRef,
    outlineHeadings,
    setHeadingInfos,
    refreshHeadings,
    toolbarHeight,
    toolbarRefCallback,
    editorScrollLeft,
  } = useOutlineSidebar({
    showOutlineProp,
    pagedEditorRef,
    scrollContainerRef,
    isLoading: state.isLoading,
  });
  // Keep history.state accessible in stable callbacks without stale closures
  const historyStateRef = useRef(history.state);
  historyStateRef.current = history.state;
  // Track current border color/width for border presets (like Google Docs)
  const borderSpecRef = useRef({ style: 'single', size: 4, color: { rgb: '000000' } });
  // Cache style resolver to avoid recreating on every selection change
  const styleResolverCacheRef = useRef<{
    styles: unknown;
    resolver: ReturnType<typeof createStyleResolver>;
  } | null>(null);
  const getCachedStyleResolver = useCallback(
    (styles: Parameters<typeof createStyleResolver>[0]) => {
      const cached = styleResolverCacheRef.current;
      if (cached && cached.styles === styles) {
        return cached.resolver;
      }
      const resolver = createStyleResolver(styles);
      styleResolverCacheRef.current = { styles, resolver };
      return resolver;
    },
    []
  );

  const { focusActiveEditor, undoActiveEditor, redoActiveEditor } = useActiveEditor({
    pagedEditorRef,
  });

  // Find/Replace hook
  const findReplace = useFindReplace();

  // Hyperlink dialog hook
  const hyperlinkDialog = useHyperlinkDialog();

  // Lifted out of useDocumentLoader / useCommentLifecycle so `resetForNewDocument`
  // (declared next) can clear both on every fresh load.
  const commentsLoadedRef = useRef(false);
  const trackedChangesLoadedRef = useRef(false);

  // One comment/revision ID allocator per editor instance (monotonic, no reuse).
  // Seeded above the loaded doc's max ID on load; shared by every comment/
  // tracked-change allocation in this component and its hooks.
  const commentIdAllocatorRef = useRef(createCommentIdAllocator());

  const { resetForNewDocument } = useResetEditorState({
    commentsLoadedRef,
    trackedChangesLoadedRef,
    setComments,
    setHeadingInfos,
    setShowCommentsSidebar,
    setIsAddingComment,
    setCommentSelectionRange,
    setAddCommentYPosition,
    setFloatingCommentBtn,
    setHfEditPosition,
    setHfEditIsFirstPage,
    setAnchorPositions,
    clearFindReplaceMatches: useCallback(() => findReplace.setMatches([], 0), [findReplace]),
    cleanOrphanedCommentsTimerRef,
  });

  const { loadParsedDocument, loadBuffer, yrsSeedDocument } = useDocumentLoader({
    documentBuffer,
    initialDocument,
    externalContent: false,
    history,
    pagedEditorRef,
    setLoadingState: useCallback((s: { isLoading: boolean; parseError: string | null }) => {
      setState((prev) => ({ ...prev, isLoading: s.isLoading, parseError: s.parseError }));
    }, []),
    setComments,
    setShowCommentsSidebar,
    onError,
    resetForNewDocument,
    commentsLoadedRef,
    commentIdAllocator: commentIdAllocatorRef.current,
    setDocumentFonts,
  });

  const {
    imageInputRef,
    docxInputRef,
    handleSave,
    handleDirectPrint,
    handleDownloadDocument,
    handleOpenDocument,
    handleDocxFileChange,
    handleInsertImageClick,
    handleImageFileChange,
  } = useFileIO({
    pagedEditorRef,
    displayList: canvasRenderer.displayList,
    resolveImage: canvasRenderer.resolveImage,
    comments,
    documentName,
    onSave,
    onOpen,
    onError,
    onPrint,
    onDocumentNameChange,
    loadBuffer,
    focusActiveEditor,
  });

  // Auto-open the sidebar once if the loaded document already has tracked changes.
  useCommentLifecycle({
    commentToRevision,
    setComments,
    isLoading: state.isLoading,
    trackedChangesCount: trackedChanges.length,
    setShowCommentsSidebar,
    trackedChangesLoadedRef,
  });

  useFontLifecycle(fonts, onFontsLoadedCallback, onError);

  const pushDocument = useCallback(
    (document: Document) => {
      history.push(document);
      return document;
    },
    [history]
  );

  // Handle document change
  const handleDocumentChange = useCallback(
    (newDocument: Document) => {
      pushDocument(newDocument);
      onChange?.(newDocument);
      // Fan out to bridge subscribers (errors in one don't break the others).
      for (const cb of contentChangeSubscribersRef.current) {
        try {
          cb(newDocument);
        } catch (e) {
          console.error('contentChange subscriber threw:', e);
        }
      }
      // Update outline headings if sidebar is open
      if (showOutlineRef.current) refreshHeadings();
      // Clean up orphaned comments (debounced — avoid yanking comments mid-edit)
      if (cleanOrphanedCommentsTimerRef.current) {
        clearTimeout(cleanOrphanedCommentsTimerRef.current);
      }
      cleanOrphanedCommentsTimerRef.current = setTimeout(cleanOrphanedComments, 300);
    },
    [onChange, pushDocument, cleanOrphanedComments, refreshHeadings, showOutlineRef]
  );

  // Recompute the floating "add comment" button position from the current Yrs
  // selection + page/container geometry. Called from handleSelectionChange and
  // from the geometry-change effects below (resize, zoom), because PagedEditor's
  // onSelectionChange no longer fires on mere overlay redraws after the
  // state-identity dedup in #268.
  const { recomputeFloatingCommentBtn } = useFloatingCommentBtn({
    pagedEditorRef,
    scrollContainerRef,
    editorContentRef,
    isAddingCommentRef,
    setFloatingCommentBtn,
    readOnly,
    isLoading: state.isLoading,
    zoom: state.zoom,
    canvasHostRef: canvasRenderer.canvasHostRef,
    displayListQueries: canvasRenderer.queries,
  });

  const { handleYrsSelectionChange } = useSelectionTracker({
    borderSpecRef,
    theme,
    historyStateRef,
    getCachedStyleResolver,
    setFloatingCommentBtn,
    applySelectionDelta: useCallback(
      (delta: SelectionStateDelta) =>
        setState((prev) => {
          const unchanged = Object.entries(delta).every(([key, next]) => {
            const current = prev[key as keyof EditorState];
            return current === next || JSON.stringify(current) === JSON.stringify(next);
          });
          return unchanged ? prev : { ...prev, ...delta };
        }),
      []
    ),
    recomputeFloatingCommentBtn,
    onSelectionChange,
    selectionChangeSubscribersRef,
    canvasA11yNotifyRef,
  });

  // Table selection hook
  const tableSelection = useTableSelection({
    document: history.state,
    onChange: handleDocumentChange,
    onSelectionChange: (_context) => {
      // Could notify parent of table selection changes
    },
  });

  useKeyboardShortcuts({
    pagedEditorRef,
    disableFindReplaceShortcuts,
    showFileOpen,
    onOpenDocument: handleOpenDocument,
    findReplace,
    hyperlinkDialog,
    tableSelection,
  });

  // Handle table insert from toolbar
  // Toggle document outline sidebar
  const handleToggleOutline = useCallback(() => {
    setShowOutline((prev) => {
      if (!prev) refreshHeadings();
      return !prev;
    });
  }, [refreshHeadings, setShowOutline]);

  // Navigate to a heading from the outline
  const handleHeadingInfoClick = useCallback((pmPos: number) => {
    pagedEditorRef.current?.scrollToPosition(pmPos);
    // Also set selection to the heading
    pagedEditorRef.current?.setSelection(pmPos + 1);
    pagedEditorRef.current?.focus();
  }, []);

  // Handle shape insertion
  // Handle image wrap type change
  const {
    imagePositionOpen,
    setImagePositionOpen,
    imagePropsOpen,
    setImagePropsOpen,
    footnotePropsOpen,
    setFootnotePropsOpen,
    handleImageWrapType,
    handleImageTransform,
    handleApplyImagePosition,
    handleOpenImageProperties,
    handleApplyImageProperties,
    handleApplyFootnoteProperties,
  } = useImageActions({
    document: history.state,
    pmImageContext: state.pmImageContext,
    displayListQueries: canvasRenderer.queries,
    pagedEditorRef,
    focusActiveEditor,
    pushDocument,
  });

  const {
    tablePropsOpen,
    setTablePropsOpen,
    currentTableProperties,
    handleTablePropertiesApply,
    splitCellDialogState,
    openSplitCellDialog,
    handleTableAction,
    handleSplitCellDialogClose,
    handleSplitCellDialogApply,
  } = useTableDialogs({
    pagedEditorRef,
    borderSpecRef,
  });

  const {
    handleFormat,
    handleInsertTable,
    handleInsertPageBreak,
    handleInsertSectionBreakNextPage,
    handleInsertSectionBreakContinuous,
    handleInsertTOC,
  } = useFormattingActions({
    focusActiveEditor,
    pagedEditorRef,
    hyperlinkDialog,
  });

  const handleZoomChange = useCallback((zoom: number) => {
    setState((prev) => ({ ...prev, zoom }));
  }, []);

  const {
    hyperlinkPopupData,
    handleHyperlinkSubmit,
    handleHyperlinkRemove,
    handleHyperlinkClick,
    handleHyperlinkPopupNavigate,
    handleHyperlinkPopupCopy,
    handleHyperlinkPopupEdit,
    handleHyperlinkPopupRemove,
    handleHyperlinkPopupClose,
  } = useHyperlinkActions({
    hyperlinkDialog,
    pagedEditorRef,
    focusActiveEditor,
  });

  const {
    contextMenu,
    imageContextMenu,
    handleEditorContextMenu,
    handleContextMenu,
    handleContextMenuClose,
    handleImageWrapApply,
    imageContextMenuTextActions,
    contextMenuItems,
    handleContextMenuAction,
  } = useContextMenus({
    pagedEditorRef,
    focusActiveEditor,
    openSplitCellDialog,
    editorContentRef,
    displayListQueries: canvasRenderer.queries,
    interactionPageHostRef: canvasRenderer.canvasHostRef,
    i18n,
    onAddComment: useCallback(
      ({ from, to, yPos }: { from: number; to: number; yPos: number | null }) => {
        setCommentSelectionRange({ from, to });
        setAddCommentYPosition(yPos);
        setShowCommentsSidebar(true);
        setIsAddingComment(true);
        setFloatingCommentBtn(null);
      },
      []
    ),
  });

  // Handle margin changes from rulers
  const {
    showPageSetup,
    setShowPageSetup,
    handleOpenPageSetup,
    handleLeftMarginChange,
    handleRightMarginChange,
    handleTopMarginChange,
    handleBottomMarginChange,
    handlePageSetupApply,
    handleIndentLeftChange,
    handleIndentRightChange,
    handleFirstLineIndentChange,
    handleTabStopRemove,
  } = usePageSetupControls({
    document: history.state,
    readOnly,
    handleDocumentChange,
    pagedEditorRef,
  });

  const {
    showWatermark,
    setShowWatermark,
    handleOpenWatermark,
    currentWatermark,
    handleWatermarkApply,
  } = useWatermarkControls({
    readOnly,
    document: history.state,
    pushDocument,
  });

  const { scrollPageInfo, setScrollPageInfo } = useScrollPageInfo({
    scrollContainerRef,
    pagedEditorRef,
  });

  // Handle save
  // Handle error from editor
  const handleEditorError = useCallback(
    (error: Error) => {
      onError?.(error);
    },
    [onError]
  );

  const {
    findResultRef,
    handleFind,
    handleFindNext,
    handleFindPrevious,
    handleReplace,
    handleReplaceAll,
  } = useFindReplaceBridge({
    pagedEditorRef,
    findReplace,
  });

  // Canvas-mode find highlights. The bridge stores the live display range on every
  // match (`YrsFindMatch`), so the matches held in `findReplace.state` carry the
  // display positions the display-list `range_rects` query needs. Memoized off the
  // reactive matches array so the overlay effect only re-runs when the result
  // set actually changes. Only resolves a portal target while the canvas paints
  // (queries != null); the DOM-painter path is untouched.
  const canvasFindMatches = useMemo(
    () =>
      (findReplace.state.matches as YrsFindMatch[]).map((m) => ({
        displayFrom: m.displayFrom,
        displayTo: m.displayTo,
      })),
    [findReplace.state.matches]
  );
  const canvasFindOverlayTarget = useCanvasOverlayTarget(
    canvasRenderer.queries != null,
    editorContentRef
  );

  // Canvas-mode brighten for the expanded sidebar card (G3). On the DOM-painter
  // path DocxEditorShell injects a `<style>` matching the parked (invisible)
  // painter; on canvas we resolve the focused comment/tracked-change to its body
  // display range and draw a tinted overlay over its display-list geometry instead.
  // aligned. HF-scoped items (`hfRid`) are not body-`range_rects`-resolvable, so
  // they fall through (no canvas brighten) — a noted follow-up.
  const canvasBrightenRange = useMemo<CanvasBrightenRange | null>(() => {
    if (!canvasRenderer.queries || !expandedSidebarItem) return null;
    if (expandedSidebarItem.startsWith('comment-')) {
      const id = parseInt(expandedSidebarItem.slice('comment-'.length), 10);
      if (!Number.isFinite(id)) return null;
      const session = pagedEditorRef.current?.getYrsSession();
      if (!session) return null;
      try {
        const anchors = session.resolveComment(String(id));
        const projection = createYrsSidebarProjection(session);
        const start = anchors
          .map((anchor) => projection.storyOffsetToDisplayPoint(anchor.story, anchor.start))
          .filter((point): point is NonNullable<typeof point> => point != null)
          .sort((a, b) => a.position - b.position)[0];
        const end = anchors
          .map((anchor) => projection.storyOffsetToDisplayPoint(anchor.story, anchor.end))
          .filter((point): point is NonNullable<typeof point> => point != null)
          .sort((a, b) => b.position - a.position)[0];
        return start && end && !start.hfRid
          ? { from: start.position, to: end.position, variant: 'comment' }
          : null;
      } catch {
        return null;
      }
    }
    if (expandedSidebarItem.startsWith('tc-')) {
      const revId = expandedSidebarItem.split('-')[1];
      const tc = trackedChanges.find((c) => String(c.revisionId) === revId);
      if (!tc || (tc as { hfRid?: string }).hfRid) return null;
      const isDeletion = /deletion|deleted/i.test(tc.type);
      return { from: tc.from, to: tc.to, variant: isDeletion ? 'deletion' : 'insertion' };
    }
    return null;
  }, [canvasRenderer.queries, expandedSidebarItem, trackedChanges]);

  // Expose ref methods
  useDocxEditorRefApi({
    ref,
    document: history.state,
    historyStateRef,
    pagedEditorRef,
    handleSave,
    handleDirectPrint,
    zoom: state.zoom,
    setZoom: (zoom: number) => setState((prev) => ({ ...prev, zoom })),
    scrollPageInfo,
    loadParsedDocument,
    loadBuffer,
    comments,
    setComments,
    setShowCommentsSidebar,
    contentChangeSubscribersRef,
    selectionChangeSubscribersRef,
    getCachedStyleResolver,
    commentIdAllocator: commentIdAllocatorRef.current,
  });

  const initialSectionProperties = useMemo(
    () => getInitialSectionProperties(history.state),
    [history.state]
  );
  const finalSectionProperties = history.state?.package.document?.finalSectionProperties;

  const {
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    handleHeaderFooterDoubleClick,
    handleBodyClick,
    handleRemoveHeaderFooter,
  } = useHeaderFooterEditing({
    document: history.state,
    pushDocument,
    initialSectionProperties,
    finalSectionProperties,
    hfEditPosition,
    setHfEditPosition,
    hfEditIsFirstPage,
    setHfEditIsFirstPage,
    setHfEditPageIndex,
  });

  // Container styles - using overflow: auto so sticky toolbar works
  const containerStyle: CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
    backgroundColor: 'var(--doc-bg)',
    ...style,
  };

  const mainContentStyle: CSSProperties = {
    display: 'flex',
    flex: 1,
    minHeight: 0, // Allow flex item to shrink below content size
    minWidth: 0, // Allow flex item to shrink below content width on narrow viewports
    flexDirection: 'row',
  };

  // --- Unified sidebar items ---
  const refreshTrackedChanges = (session: YrsSession): void => {
    setYrsTrackedChangesResult(
      extractTrackedChangesFromYrs(session.listRevisions(), createYrsSidebarProjection(session))
    );
  };
  const commentCallbacksRef = useRef<CommentCallbacks>({});
  commentCallbacksRef.current = {
    onCommentReply: (id, text) => {
      const reply = createComment(commentIdAllocatorRef.current, text, author, id);
      const parent = comments.find((c) => c.id === id);
      setComments((prev) => [...prev, reply]);
      if (parent) onCommentReply?.(reply, parent);
    },
    onCommentResolve: (id) => {
      const target = comments.find((c) => c.id === id);
      setComments((prev) => prev.map((c) => (c.id === id ? { ...c, done: true } : c)));
      // Collapse the card to its checkmark marker immediately.
      if (expandedSidebarItem === `comment-${id}`) {
        setExpandedSidebarItem(null);
      }
      if (target) onCommentResolve?.({ ...target, done: true });
    },
    onCommentUnresolve: (id) => {
      setComments((prev) => prev.map((c) => (c.id === id ? { ...c, done: undefined } : c)));
    },
    onCommentDelete: (id) => {
      const target = comments.find((c) => c.id === id);
      setComments((prev) => prev.filter((c) => c.id !== id && c.parentId !== id));
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (session) {
        try {
          session.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
          editor?.syncYrsInputState(true);
        } catch {
          // The anchor may already have disappeared with its content.
        }
      }
      if (target) onCommentDelete?.(target);
    },
    onAddComment: (addText) => {
      const comment = createComment(commentIdAllocatorRef.current, addText, author);
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (editor && session && commentSelectionRange) {
        const { from, to } = commentSelectionRange;
        const start = editor.displayPositionToYrsLoc(from);
        const end = editor.displayPositionToYrsLoc(to);
        if (start && end && start.story === end.story) {
          session.applyRawOps(start.story, [
            {
              op: 'setComment',
              id: String(comment.id),
              ranges: [[yrsStoryOffset(session, start), yrsStoryOffset(session, end)]],
              author,
              date: comment.date,
              body: comment.content,
            },
          ]);
          editor.syncYrsInputState(true);
        }
      }
      setComments((prev) => [...prev, comment]);
      setIsAddingComment(false);
      setCommentSelectionRange(null);
      setAddCommentYPosition(null);
      onCommentAdd?.(comment);
    },
    onCancelAddComment: () => {
      setIsAddingComment(false);
      setCommentSelectionRange(null);
      setAddCommentYPosition(null);
    },
    onAcceptChange: (from, to) => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      const range = editor ? displayRangeToYrsRange(editor, from, to) : null;
      if (!session || !range) return;
      session.acceptChange(range);
      editor?.syncYrsInputState(true);
      refreshTrackedChanges(session);
    },
    onRejectChange: (from, to) => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      const range = editor ? displayRangeToYrsRange(editor, from, to) : null;
      if (!session || !range) return;
      session.rejectChange(range);
      editor?.syncYrsInputState(true);
      refreshTrackedChanges(session);
    },
    onAcceptChangeById: (revisionId) => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      const revision = session
        ?.listRevisions()
        .find((candidate) => yrsIdToNumericId(candidate.revisionId) === revisionId);
      if (revision && session) {
        session.acceptChange({ revisionId: revision.revisionId });
        editor?.syncYrsInputState(true);
        refreshTrackedChanges(session);
      }
    },
    onRejectChangeById: (revisionId) => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      const revision = session
        ?.listRevisions()
        .find((candidate) => yrsIdToNumericId(candidate.revisionId) === revisionId);
      if (revision && session) {
        session.rejectChange({ revisionId: revision.revisionId });
        editor?.syncYrsInputState(true);
        refreshTrackedChanges(session);
      }
    },
    onTrackedChangeReply: (revisionId, text) => {
      setComments((prev) => [
        ...prev,
        createComment(commentIdAllocatorRef.current, text, author, revisionId),
      ]);
    },
  };

  // Stable callbacks wrapper that delegates to ref (avoids recreating items on every render)
  const stableCallbacks = useMemo<CommentCallbacks>(
    () => ({
      onCommentReply: (...args) => commentCallbacksRef.current.onCommentReply?.(...args),
      onCommentResolve: (...args) => commentCallbacksRef.current.onCommentResolve?.(...args),
      onCommentUnresolve: (...args) => commentCallbacksRef.current.onCommentUnresolve?.(...args),
      onCommentDelete: (...args) => commentCallbacksRef.current.onCommentDelete?.(...args),
      onAddComment: (...args) => commentCallbacksRef.current.onAddComment?.(...args),
      onCancelAddComment: (...args) => commentCallbacksRef.current.onCancelAddComment?.(...args),
      onAcceptChange: (...args) => commentCallbacksRef.current.onAcceptChange?.(...args),
      onRejectChange: (...args) => commentCallbacksRef.current.onRejectChange?.(...args),
      onAcceptChangeById: (...args) => commentCallbacksRef.current.onAcceptChangeById?.(...args),
      onRejectChangeById: (...args) => commentCallbacksRef.current.onRejectChangeById?.(...args),
      onTrackedChangeReply: (...args) =>
        commentCallbacksRef.current.onTrackedChangeReply?.(...args),
    }),
    []
  );

  const commentSidebarItems = useCommentSidebarItems({
    comments,
    trackedChanges,
    callbacks: stableCallbacks,
    showResolved: showCommentsSidebar,
    isAddingComment: showCommentsSidebar ? isAddingComment : false,
    addCommentYPosition,
  });

  const allSidebarItems = useMemo(() => {
    const items: ReactSidebarItem[] = [];
    if (showCommentsSidebar) items.push(...commentSidebarItems);
    if (pluginSidebarItems) items.push(...pluginSidebarItems);
    return items;
  }, [showCommentsSidebar, commentSidebarItems, pluginSidebarItems]);

  // Build a map from insertion revisionIds to sidebar item IDs for replacement tracked changes.
  // This allows clicking the insertion part of a replacement to activate the same sidebar card.
  const revisionIdAliases = useMemo(() => {
    const map = new Map<string, string>();
    trackedChanges.forEach((change, idx) => {
      if (change.type === 'replacement' && change.insertionRevisionId != null) {
        map.set(String(change.insertionRevisionId), `tc-${change.revisionId}-${idx}`);
      }
    });
    return map;
  }, [trackedChanges]);

  const sidebarOpen = allSidebarItems.length > 0;
  // Reserve 2× the left-edge allowance so the centered page clears whatever
  // outline UI is showing, without forcing a shift on wide viewports.
  const outlineLeftAllowance =
    (showOutline
      ? OUTLINE_RESERVED_SPACE
      : showOutlineButton
        ? OUTLINE_BUTTON_RESERVED_SPACE
        : 20) +
    // The outline toggle/panel inset past the vertical ruler when it's shown,
    // so the page must clear that extra width too.
    (showRuler && (showOutline || showOutlineButton) ? RULER_WIDTH : 0);
  // Reserve against the WIDEST page in the doc, not the portrait default: pages
  // center via `alignItems:center`, so a landscape section (wider than
  // DEFAULT_PAGE_WIDTH) gets a smaller side margin and, with the old default,
  // slid left under the outline toggle/panel. Taking the max across all section
  // widths also covers mixed-orientation docs.
  const docBody = history.state?.package?.document;
  const sectionPageWidths = [
    docBody?.finalSectionProperties?.pageWidth,
    ...(docBody?.sections?.map((s) => s.properties?.pageWidth) ?? []),
  ].filter((w): w is number => typeof w === 'number' && w > 0);
  const maxPageWidthPx = sectionPageWidths.length
    ? Math.round(Math.max(...sectionPageWidths) / 15)
    : DEFAULT_PAGE_WIDTH;

  const minLayoutWidth =
    2 * outlineLeftAllowance + maxPageWidthPx + (sidebarOpen ? SIDEBAR_DOCUMENT_SHIFT * 2 : 0);

  const liveYrsStory = (() => {
    try {
      return pagedEditorRef.current?.getYrsSession()?.selection()?.head.story;
    } catch {
      return undefined;
    }
  })();
  const toolbarTableContext =
    liveYrsStory && /:t\d+:r\d+c\d+/.test(liveYrsStory)
      ? state.pmTableContext?.isInTable
        ? state.pmTableContext
        : { isInTable: true, canSplitCell: true }
      : state.pmTableContext;

  // pageWidthPx — the final section's width — positions the sidebar / comment
  // margin markers against the page most content lives under.
  const sectionPropsPageWidth = docBody?.finalSectionProperties?.pageWidth;
  const pageWidthPx = sectionPropsPageWidth
    ? Math.round(sectionPropsPageWidth / 15)
    : DEFAULT_PAGE_WIDTH;

  // PagedEditor selection callback: resolve sticky comment/revision coverage
  // from Yrs so the matching sidebar card opens as the caret moves.
  const handlePagedSelectionChange = useCallback(() => {
    // Body selection transitions arrive here even when the derived toolbar
    // context is unchanged. Notify the canvas live region from the authoritative
    // selection event so range/caret announcements are never lost to toolbar
    // state deduplication.
    canvasA11yNotifyRef.current?.();
    const session = pagedEditorRef.current?.getYrsSession();
    const head = session?.selection()?.head;
    if (!session || !head) return;
    const offset = yrsStoryOffset(session, head);
    let cursorSidebarItem: string | null = null;
    for (const comment of comments) {
      if (comment.parentId != null || resolvedCommentIds.has(comment.id)) continue;
      try {
        if (
          session
            .resolveComment(String(comment.id))
            .some((anchor) => anchor.story === head.story && anchor.start <= offset && offset <= anchor.end)
        ) {
          cursorSidebarItem = `comment-${comment.id}`;
          break;
        }
      } catch {
        // Ignore comments whose anchor disappeared between selection events.
      }
    }
    if (!cursorSidebarItem) {
      for (const revision of session.listRevisions()) {
        if (revision.range.story !== head.story) continue;
        const start = session.locateParagraph(head.story, revision.range.start.paraId).start + revision.range.start.offset;
        const end = session.locateParagraph(head.story, revision.range.end.paraId).start + revision.range.end.offset;
        if (start <= offset && offset <= end) {
          const revId = String(yrsIdToNumericId(revision.revisionId));
          const prefix = `tc-${revId}-`;
          let match = commentSidebarItems.find((item) => item.id.startsWith(prefix));
          if (!match) {
            const aliasedId = revisionIdAliases.get(revId);
            if (aliasedId) match = commentSidebarItems.find((item) => item.id === aliasedId);
          }
          if (match) cursorSidebarItem = match.id;
          break;
        }
      }
    }
    if (cursorSidebarItem) {
      setShowCommentsSidebar(true);
    }
    setExpandedSidebarItem(cursorSidebarItem);
  }, [comments, resolvedCommentIds, commentSidebarItems, revisionIdAliases, setShowCommentsSidebar]);

  const handleYrsToolbarSelectionChange = useCallback(
    (selection: YrsToolbarSelection) => {
      handleYrsSelectionChange(selection);
    },
    [handleYrsSelectionChange]
  );

  // Auto-open the sidebar the first time a comment or tracked-change card
  // is produced — covers the case where the user inserts an empty tracked
  // table: no cursor anchor exists yet (no inline marks at cursor), so the
  // cursor-driven open above doesn't fire. Latches via a ref so a later
  // manual close stays closed.
  useEffect(() => {
    if (sidebarAutoOpenedRef.current) return;
    if (commentSidebarItems.length === 0) return;
    sidebarAutoOpenedRef.current = true;
    setShowCommentsSidebar(true);
  }, [commentSidebarItems]);

  const editorContainerStyle: CSSProperties = {
    flex: 1,
    minHeight: 0,
    minWidth: 0, // Allow flex item to shrink below content width on narrow viewports
    overflow: 'auto', // Sole scroll container — PagedEditor sizes to content
    position: 'relative',
    overflowAnchor: 'none',
  };

  // Render loading state
  if (state.isLoading) {
    return (
      <div
        className={cn('oox-root docx-editor docx-editor-loading', isDark && 'dark', className)}
        style={containerStyle}
        data-testid="docx-editor"
      >
        {loadingIndicator || <DefaultLoadingIndicator />}
      </div>
    );
  }

  // Render error state
  if (state.parseError) {
    return (
      <div
        className={cn('oox-root docx-editor docx-editor-error', isDark && 'dark', className)}
        style={containerStyle}
        data-testid="docx-editor"
      >
        <ParseError message={state.parseError} />
      </div>
    );
  }

  // Render placeholder when no document
  if (!history.state) {
    return (
      <div
        className={cn('oox-root docx-editor docx-editor-empty', isDark && 'dark', className)}
        style={containerStyle}
        data-testid="docx-editor"
      >
        {placeholder || <DefaultPlaceholder />}
      </div>
    );
  }

  const handleScrollContainerMouseDown = (e: React.MouseEvent) => {
    // Click in the grey gutter around the page → collapse any expanded sidebar
    // card. Clicks on the doc body already collapse via the cursor-mark
    // detector; clicks inside the sidebar are user interactions with the card.
    const target = e.target as HTMLElement;
    if (
      // Accepts both renderers' page hosts (`.paged-editor__pages` painter,
      // `.canvas-pages` canvas) so canvas body clicks don't wrongly collapse.
      isWithinPageArea(target) ||
      target.closest('.docx-unified-sidebar') ||
      target.closest('.docx-comment-margin-markers')
    ) {
      return;
    }
    setExpandedSidebarItem(null);
  };

  const handleEditorBgMouseDown = (e: React.MouseEvent) => {
    // Focus editor when clicking on the background area (not the editor itself).
    // mouseDown for immediate response before focus can be lost.
    if (e.target === e.currentTarget) {
      e.preventDefault();
      pagedEditorRef.current?.focus();
    }
  };

  return (
    <>
      <DocxEditorShell
        i18n={i18n}
        isDark={isDark}
        onEditorError={handleEditorError}
        containerRef={containerRef}
        scrollContainerRef={scrollContainerRef}
        editorContentRef={editorContentRef}
        className={className}
        containerStyle={containerStyle}
        mainContentStyle={mainContentStyle}
        editorContainerStyle={editorContainerStyle}
        showRuler={showRuler}
        readOnlyProp={readOnlyProp}
        showOutline={showOutline}
        showOutlineButton={showOutlineButton}
        sidebarOpen={sidebarOpen}
        minLayoutWidth={minLayoutWidth}
        toolbarHeight={toolbarHeight}
        editorScrollLeft={editorScrollLeft}
        expandedSidebarItem={expandedSidebarItem}
        trackedChanges={trackedChanges}
        onScrollContainerMouseDown={handleScrollContainerMouseDown}
        onEditorBgMouseDown={handleEditorBgMouseDown}
        onEditorContextMenu={handleEditorContextMenu}
        horizontalRulerProps={{
          sectionProps: history.state?.package.document?.finalSectionProperties,
          zoom: state.zoom,
          unit: rulerUnit,
          editable: !readOnly,
          onLeftMarginChange: handleLeftMarginChange,
          onRightMarginChange: handleRightMarginChange,
          indentLeft: state.paragraphIndentLeft,
          indentRight: state.paragraphIndentRight,
          onIndentLeftChange: handleIndentLeftChange,
          onIndentRightChange: handleIndentRightChange,
          firstLineIndent: state.paragraphFirstLineIndent,
          hangingIndent: state.paragraphHangingIndent,
          onFirstLineIndentChange: handleFirstLineIndentChange,
          tabStops: state.paragraphTabs,
          onTabStopRemove: handleTabStopRemove,
        }}
        verticalRulerProps={{
          sectionProps: initialSectionProperties,
          zoom: state.zoom,
          unit: rulerUnit,
          editable: !readOnly,
          onTopMarginChange: handleTopMarginChange,
          onBottomMarginChange: handleBottomMarginChange,
        }}
        outlineProps={{
          headings: outlineHeadings,
          onHeadingClick: handleHeadingInfoClick,
          onClose: () => setShowOutline(false),
          topOffset: toolbarHeight,
          scrollLeft: editorScrollLeft,
        }}
        onToggleOutline={handleToggleOutline}
        scrollPageInfo={scrollPageInfo}
        toolbar={
          showToolbar && !readOnlyProp ? (
            <DocxEditorToolbar
              toolbarRefCallback={toolbarRefCallback}
              document={history.state}
              theme={theme}
              authoritativeCanUndo={yrsHistoryState.canUndo}
              authoritativeCanRedo={yrsHistoryState.canRedo}
              selectionFormatting={state.selectionFormatting}
              tableContext={toolbarTableContext}
              imageContext={state.pmImageContext}
              readOnly={readOnly}
              editingMode={editingMode}
              setEditingMode={setEditingMode}
              setShowCommentsSidebar={setShowCommentsSidebar}
              setExpandedSidebarItem={setExpandedSidebarItem}
              showCommentsSidebar={showCommentsSidebar}
              renderLogo={renderLogo}
              documentName={documentName}
              onDocumentNameChange={onDocumentNameChange}
              documentNameEditable={documentNameEditable}
              renderTitleBarRight={renderTitleBarRight}
              toolbarExtra={toolbarExtra}
              fontFamilies={fontFamilies}
              documentFonts={documentFonts}
              zoom={state.zoom}
              showZoomControl={showZoomControl}
              onFormat={handleFormat}
              onUndo={undoActiveEditor}
              onRedo={redoActiveEditor}
              onPrint={handleDirectPrint}
              showFileOpen={showFileOpen}
              showHelpMenu={showHelpMenu}
              onOpen={handleOpenDocument}
              onSave={handleDownloadDocument}
              onZoomChange={handleZoomChange}
              onRefocusEditor={focusActiveEditor}
              onInsertTable={handleInsertTable}
              onInsertImage={handleInsertImageClick}
              onInsertPageBreak={handleInsertPageBreak}
              onInsertSectionBreakNextPage={handleInsertSectionBreakNextPage}
              onInsertSectionBreakContinuous={handleInsertSectionBreakContinuous}
              onInsertTOC={handleInsertTOC}
              onImageWrapType={handleImageWrapType}
              onImageTransform={handleImageTransform}
              onOpenImageProperties={handleOpenImageProperties}
              onPageSetup={handleOpenPageSetup}
              onWatermark={handleOpenWatermark}
              onTableAction={handleTableAction}
            />
          ) : null
        }
        pagedArea={
          <CanvasPagedArea
            renderer={canvasRenderer}
            a11y={{
              getYrsSession: () => pagedEditorRef.current?.getYrsSession(),
              notifyRef: canvasA11yNotifyRef,
            }}
            sidebarOpen={sidebarOpen}
            zoom={state.zoom}
            interactive={!readOnly}
          >
            <DocxEditorPagedArea
              yrsSeedDocument={yrsSeedDocument}
              collaboration={collaboration}
              pagedEditorRef={pagedEditorRef}
              scrollContainerRef={scrollContainerRef}
              editorContentRef={editorContentRef}
              document={history.state}
              theme={theme}
              initialSectionProperties={initialSectionProperties}
              finalSectionProperties={finalSectionProperties}
              headerContent={headerContent}
              footerContent={footerContent}
              firstPageHeaderContent={firstPageHeaderContent}
              firstPageFooterContent={firstPageFooterContent}
              hfEditPosition={hfEditPosition}
              setHfEditPosition={setHfEditPosition}
              hfEditIsFirstPage={hfEditIsFirstPage}
              hfEditPageIndex={hfEditPageIndex}
              onHeaderFooterDoubleClick={handleHeaderFooterDoubleClick}
              onRemoveHeaderFooter={handleRemoveHeaderFooter}
              onBodyClick={handleBodyClick}
              zoom={state.zoom}
              readOnly={readOnly}
              isSuggesting={editingMode === 'suggesting'}
              author={author}
              measurementFontProvider={measurementFontProvider}
              rustFontChainsProviderRef={rustFontChainsProviderRef}
              onDocumentChange={handleDocumentChange}
              onYrsHistoryChange={handleYrsHistoryChange}
              onPagedSelectionChange={handlePagedSelectionChange}
              onYrsSelectionChange={handleYrsToolbarSelectionChange}
              onRenderedDomContextReady={onRenderedDomContextReady}
              pluginOverlays={pluginOverlays}
              onHyperlinkClick={handleHyperlinkClick}
              hyperlinkPopupData={hyperlinkPopupData}
              onHyperlinkPopupNavigate={handleHyperlinkPopupNavigate}
              onHyperlinkPopupCopy={handleHyperlinkPopupCopy}
              onHyperlinkPopupEdit={handleHyperlinkPopupEdit}
              onHyperlinkPopupRemove={handleHyperlinkPopupRemove}
              onHyperlinkPopupClose={handleHyperlinkPopupClose}
              onContextMenu={handleContextMenu}
              sidebarOpen={sidebarOpen}
              sidebarItems={allSidebarItems}
              anchorPositions={anchorPositions}
              onAnchorPositionsChange={setAnchorPositions}
              onYrsTrackedChangesChange={setYrsTrackedChangesResult}
              pluginRenderedDomContext={pluginRenderedDomContext}
              pageWidthPx={pageWidthPx}
              expandedSidebarItem={expandedSidebarItem}
              setExpandedSidebarItem={setExpandedSidebarItem}
              comments={comments}
              resolvedCommentIds={resolvedCommentIds}
              resolvedIdsForRender={resolvedIdsForRender}
              setShowCommentsSidebar={setShowCommentsSidebar}
              onTotalPagesChange={(totalPages) => {
                setScrollPageInfo((prev) =>
                  prev.totalPages === totalPages ? prev : { ...prev, totalPages }
                );
              }}
              onLayoutComputed={canvasRenderer.onLayoutComputed}
              applyResidentInput={canvasRenderer.applyInput}
              applyResidentDelete={canvasRenderer.applyDelete}
              displayListQueries={canvasRenderer.queries}
              canvasDisplayList={canvasRenderer.displayList}
              displayListFrameEpoch={canvasRenderer.frame?.frameEpoch ?? null}
              residentCaret={canvasRenderer.caret}
              residentCaretAuthoritative={canvasRenderer.authoritativeCaretActive}
              paintedCaretActive={canvasRenderer.paintedCaretActive}
              onCaretInput={canvasRenderer.notifyCaretInput}
              onCaretInputDispatched={canvasRenderer.notifyCaretInputDispatched}
              onCaretInterrupt={canvasRenderer.notifyCaretInterrupt}
              canvasHostRef={canvasRenderer.canvasHostRef}
              floatingCommentBtn={floatingCommentBtn}
              isAddingComment={isAddingComment}
              setCommentSelectionRange={setCommentSelectionRange}
              setAddCommentYPosition={setAddCommentYPosition}
              setIsAddingComment={setIsAddingComment}
              setFloatingCommentBtn={setFloatingCommentBtn}
            />
            {!readOnly && (
              <ContentControlWidgets
                containerRef={containerRef}
                applyYrsValue={(pmPos, value, embedId) =>
                  pagedEditorRef.current?.applyYrsCommand({
                    type: 'contentControlValue',
                    pmPos,
                    ...(embedId ? { embedId } : {}),
                    value,
                  }) ?? false
                }
              />
            )}
          </CanvasPagedArea>
        }
        overlays={
          <DocxEditorOverlays
            contextMenu={contextMenu}
            contextMenuItems={contextMenuItems}
            onContextMenuAction={handleContextMenuAction}
            onContextMenuClose={handleContextMenuClose}
            imageContextMenu={imageContextMenu}
            onImageWrapApply={handleImageWrapApply}
            imageContextMenuTextActions={imageContextMenuTextActions}
            onOpenImageProperties={handleOpenImageProperties}
            readOnly={readOnly}
          />
        }
        dialogs={
          <DocxEditorDialogs
            findReplace={findReplace}
            findResultRef={findResultRef}
            onFind={handleFind}
            onFindNext={handleFindNext}
            onFindPrevious={handleFindPrevious}
            onReplace={handleReplace}
            onReplaceAll={handleReplaceAll}
            hyperlinkDialog={hyperlinkDialog}
            onHyperlinkSubmit={handleHyperlinkSubmit}
            onHyperlinkRemove={handleHyperlinkRemove}
            tablePropsOpen={tablePropsOpen}
            onTablePropsClose={() => setTablePropsOpen(false)}
            tableProperties={currentTableProperties}
            onTablePropertiesApply={handleTablePropertiesApply}
            splitCellDialogState={splitCellDialogState}
            onSplitCellDialogClose={handleSplitCellDialogClose}
            onSplitCellDialogApply={handleSplitCellDialogApply}
            imagePositionOpen={imagePositionOpen}
            onImagePositionClose={() => setImagePositionOpen(false)}
            onApplyImagePosition={handleApplyImagePosition}
            imagePropsOpen={imagePropsOpen}
            onImagePropsClose={() => setImagePropsOpen(false)}
            onApplyImageProperties={handleApplyImageProperties}
            pmImageContext={state.pmImageContext}
            showPageSetup={showPageSetup}
            onPageSetupClose={() => setShowPageSetup(false)}
            onPageSetupApply={handlePageSetupApply}
            showWatermark={showWatermark}
            onWatermarkClose={() => setShowWatermark(false)}
            onWatermarkApply={handleWatermarkApply}
            currentWatermark={currentWatermark}
            watermarkPresets={watermarkPresets}
            document={history.state}
            footnotePropsOpen={footnotePropsOpen}
            onFootnotePropsClose={() => setFootnotePropsOpen(false)}
            onApplyFootnoteProperties={handleApplyFootnoteProperties}
          />
        }
        fileInputs={
          <>
            <input
              ref={imageInputRef}
              type="file"
              accept="image/*"
              style={{ display: 'none' }}
              onChange={handleImageFileChange}
            />
            <input
              ref={docxInputRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              style={{ display: 'none' }}
              onChange={handleDocxFileChange}
            />
          </>
        }
      />
      {/* Canvas-mode find-match highlights, portaled onto the visible canvas
        pages. On the DOM-painter path the target stays null (rendered nothing)
        and the current match shows as an ordinary PM selection as before. */}
      {canvasFindOverlayTarget && canvasRenderer.queries ? (
        <CanvasFindHighlightOverlay
          matches={canvasFindMatches}
          currentIndex={findReplace.state.currentIndex}
          overlayTarget={canvasFindOverlayTarget}
          canvasHostRef={canvasRenderer.canvasHostRef}
          displayListQueries={canvasRenderer.queries}
          sidebarOpen={sidebarOpen}
          zoom={state.zoom}
        />
      ) : null}
      {/* Canvas-mode brighten for the expanded sidebar card (G3). On the
        DOM-painter path the target stays null and DocxEditorShell's injected
        <style> handles the brighten as before. */}
      {canvasFindOverlayTarget && canvasRenderer.queries ? (
        <CanvasSidebarBrightenOverlay
          range={canvasBrightenRange}
          overlayTarget={canvasFindOverlayTarget}
          canvasHostRef={canvasRenderer.canvasHostRef}
          displayListQueries={canvasRenderer.queries}
          sidebarOpen={sidebarOpen}
          zoom={state.zoom}
        />
      ) : null}
    </>
  );
});

// ============================================================================
// EXPORTS
// ============================================================================

export default DocxEditor;
