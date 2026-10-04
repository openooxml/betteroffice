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

import { useRef, useCallback, useState, useEffect, useLayoutEffect, useMemo, forwardRef } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { Document, Theme } from '@betteroffice/docx/types/document';
import type {
  DocxContentControlQuery,
  DocxContentControlsOptions,
  DocxContentControlsResult,
  DocxEditRequest,
  DocxEditResult,
  DocxExportResult,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxLayoutMap,
  DocxPageExportOptions,
  DocxParagraphAnchor,
  DocxParagraphAnchorResult,
  DocxParagraphIdentitySnapshot,
  DocxPagedStructuredContent,
  DocxProposalRequest,
  DocxProposalResult,
  DocxProposalSnapshot,
  DocxProposalStateRequest,
  DocxProposalWithdrawRequest,
  DocxReadParagraphsRequest,
  DocxReadParagraphsResult,
  DocxValidationResult,
  YrsLoc,
  YrsSession,
  YrsStoryRange,
} from '@betteroffice/docx/yrs';
import type { BundledFontProvider } from '@betteroffice/docx/layout';
import {
  createYrsSidebarProjection,
  extractTrackedChangesFromYrs,
  loadRustDisplayListQueryEngine,
  yrsIdToNumericId,
  type TrackedChangesResult,
} from '@betteroffice/docx/layout/render';

import { cn } from '../lib/utils';
import type {
  DocxEditorCollaborationOptions,
  DocxPointPosition,
  SelectionState,
  TableContextInfo,
} from './DocxEditor/types';
import {
  isPresented,
  onPresented,
  onReplayFailed,
  presentedWorkerVersion,
  workerFrameVersionOf,
} from './DocxEditor/internals/layoutProvenance';
import { SupersededPreviewError } from './DocxEditor/internals/supersededPreview';
import { useOutlineSidebar } from './DocxEditor/hooks/useOutlineSidebar';
import { useKeyboardShortcuts } from './DocxEditor/hooks/useKeyboardShortcuts';
import { useFileIO } from './DocxEditor/hooks/useFileIO';
import { usePageSetupControls } from './DocxEditor/hooks/usePageSetupControls';
import { useWatermarkControls } from './DocxEditor/hooks/useWatermarkControls';
import { useHyperlinkActions } from './DocxEditor/hooks/useHyperlinkActions';
import { useFindReplaceBridge, type YrsFindMatch } from './DocxEditor/hooks/useFindReplaceBridge';
import {
  useHostSearch,
  type DocxSearchOptions,
  type DocxSearchState,
} from './DocxEditor/hooks/useHostSearch';
import { CanvasFindHighlightOverlay } from './DocxEditor/overlays/CanvasFindHighlightOverlay';
import {
  CanvasSidebarBrightenOverlay,
  type CanvasBrightenRange,
} from './DocxEditor/overlays/CanvasSidebarBrightenOverlay';
import { useCanvasOverlayTarget } from './DocxEditor/internals/useCanvasOverlayTarget';
import { isWithinPageArea } from './DocxEditor/internals/pageAreaRouting';
import { requestWorkerOpenReplica, workerOpenReplicaPending } from './DocxEditor/internals/workerOpenReplica';
import { registeredWorkerProposalAuthority } from './DocxEditor/internals/workerProposalAuthority';
import { isWorkerViewer } from './DocxEditor/internals/workerViewer';
import { warnDeprecatedViewerMember } from './DocxEditor/internals/deprecatedViewerMembers';
import type { ViewerCommentRanges } from './DocxEditor/internals/viewerSidebarReads';
import { useViewerSession, viewerReadsWorker } from './DocxEditor/internals/viewerSession';
import type { ViewerSelectionChange } from './DocxEditor/internals/viewerSelectionController';
import { pagePressNeedsReplica } from './DocxEditor/internals/replicaTriggers';
import { useImageActions } from './DocxEditor/hooks/useImageActions';
import { useDocxEditorRefApi } from './DocxEditor/hooks/useDocxEditorRefApi';
import {
  useMemoryPressure,
  type DocxMemoryBudget,
  type DocxMemoryPressure,
  type DocxMemoryStats,
} from './DocxEditor/memoryStats';
import { commandOutcome, useDocxCommandBinding } from './DocxEditor/hooks/useDocxCommands';
import type {
  PagedEditorCommandBridge,
  PagedEditorImageHandle,
  PagedEditorSelectedImage,
} from './DocxEditor/hooks/usePagedEditorRefApi';
import { DocxCommandAdmissionError } from '../commands/createDocxCommandStore';
import { DocxCommandProvider } from '../commands/DocxCommandProvider';
import type { DocxCommandStore } from '../commands/types';
import { EditorChromeContext, type EditorChrome } from './EditorToolbarContext';
import { useControllableBoolean } from './DocxEditor/hooks/useControllableBoolean';
import { useTableDialogs } from './DocxEditor/hooks/useTableDialogs';
import { useHeaderFooterEditing } from './DocxEditor/hooks/useHeaderFooterEditing';
import type { PartEditTarget } from './DocxEditor/partEdit';
import { useDocumentLoader } from './DocxEditor/hooks/useDocumentLoader';
import { useCompatibilityWarm, useYrsCoreSession } from './DocxEditor/hooks/useYrsCoreSession';
import { useHostProposalRevisions } from './DocxEditor/hooks/useHostProposalRevisions';
import {
  useDocxEnginePrewarm,
  useDocxEnginePrewarmOnBytes,
} from './DocxEditor/hooks/useDocxEnginePrewarm';
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
import {
  commitLegacyDocumentChange,
  commitYrsDocumentChange,
  LEGACY_PROJECTION_DELAY_MS,
} from './DocxEditor/documentChangeCommit';
import type { FontOption } from './ui/FontPicker';
import { OUTLINE_BUTTON_RESERVED_SPACE, OUTLINE_RESERVED_SPACE } from './DocumentOutline';
import { RULER_WIDTH } from './ui/VerticalRuler';
import { SIDEBAR_DOCUMENT_SHIFT } from './sidebar/constants';
import { useCommentSidebarItems, type CommentCallbacks } from '../hooks/useCommentSidebarItems';
import type { ReactSidebarItem } from '../plugin-api/types';
import type { DocxEditorPluginProps } from '../plugins/types';
import { useDocxPluginHost } from '../plugins/useDocxPluginHost';
import { PluginOverlays } from '../plugins/PluginOverlays';
import { PluginDock } from '../plugins/PluginPanels';
import { mergeSidebarItems } from '../plugins/PluginSidebarItems';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Translations } from '@betteroffice/docx-i18n';
import { type PrintOptions } from './ui/PrintPreview';
// Dialog hooks and utilities (static imports — lightweight, no UI)
import { useFindReplace } from './dialogs/FindReplaceDialog';
import { useHyperlinkDialog } from './dialogs/HyperlinkDialog';
import { DefaultLoadingIndicator, DefaultPlaceholder, ParseError } from './DocxEditorHelpers';
import { type DocxInput } from '@betteroffice/docx/utils';
import type { FontDefinition, ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import { useFontLifecycle, useFontLoadScope } from '../hooks/useFontLifecycle';
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

export type { DocxParagraphMatch } from '@betteroffice/docx/yrs';
import type { DocxParagraphMatch } from '@betteroffice/docx/yrs';

export interface DocxSelectionInfo {
  paraId: string | null;
  selectedText: string;
  paragraphText: string;
  before: string;
  after: string;
}

export interface DocxCommentInsertion {
  paraId: string;
  text: string;
  author: string;
  search?: string;
}

export interface DocxDocumentChange {
  version: string;
}

export type { DocxEditorCollaborationOptions, DocxPointPosition } from './DocxEditor/types';

/**
 * DocxEditor props
 */
export interface DocxEditorProps extends DocxEditorPluginProps {
  /** Document data — ArrayBuffer, Uint8Array, Blob, or File */
  documentBuffer?: DocxInput | null;
  /**
   * Preload the editing engine on mount. Off by default.
   * @experimental
   */
  experimentalPrewarm?: boolean;
  /** Pre-parsed document (alternative to documentBuffer) */
  document?: Document | null;
  /** Callback when document is saved */
  onSave?: (buffer: ArrayBuffer) => void;
  /** Owns File > Save and Cmd/Ctrl+S; return true to continue the built-in save. */
  onSaveRequest?: () => boolean | void | Promise<boolean | void>;
  /** Whether Save also downloads a copy. Defaults to true. */
  downloadOnSave?: boolean;
  /** Configure the Yrs collaboration replica used by the editor. */
  collaboration?: DocxEditorCollaborationOptions;
  /**
   * Open DOCX files in the resident worker. Off by default. A read-only editor without
   * collaboration then loads its main-thread copy of the document only when something needs it.
   * While a read-only document's host proposals are held in the worker, synchronous ref members
   * that need the main-thread document throw `DocxReplicaNotReadyError`; await `flushPendingInput()` first.
   * Display lists are built for visible pages and a small margin instead of the whole document.
   * @experimental
   */
  experimentalWorkerOpen?: boolean;
  /**
   * Opens images as `media:{n}` tokens read from the document file instead of
   * `data:` URLs, keeping them out of the document state and its updates.
   * Every client of a shared room must open the same file on a version that
   * reads them. Read when a document opens. Off by default.
   */
  mediaTokens?: boolean;
  /**
   * Callback when a DOCX file is selected through `File > Open` or Cmd/Ctrl+O.
   * Pass it to route the picked file through your own import pipeline. Omit it
   * to keep the built-in local document load behavior.
   */
  onOpen?: (file: File) => void | Promise<void>;
  /** Author name used for comments and track changes */
  author?: string;
  /** @deprecated Use {@link onDocumentChange}. Worker viewers do not fire this callback. */
  onChange?: (document: Document) => void;
  /** Receives the version after a committed edit or a changed document is presented by the worker. */
  onDocumentChange?: (change: DocxDocumentChange) => void;
  /** Callback when selection changes */
  onSelectionChange?: (state: SelectionState | null) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
  /**
   * Called when the fullest wasm memory, on the main thread or in the resident
   * worker, crosses a `memoryBudget` level or drops back below it. Silent while
   * memory stays under the warning level.
   */
  onMemoryPressure?: (pressure: DocxMemoryPressure) => void;
  /**
   * Levels for `onMemoryPressure`, 75% and 90% of 4 GiB by default, and an
   * optional limit on the resident worker's allocations.
   */
  memoryBudget?: DocxMemoryBudget;
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
  /** Whether to show hidden (vanished) text in the layout (default: false) */
  showHiddenText?: boolean;
  /** Whether the editor is read-only. When true, hides toolbar and rulers */
  readOnly?: boolean;
  /**
   * Experimental: paint a display-only preview of a document's first pages
   * before the whole document is opened, then hand them over to it. The editor
   * is read-only and plugins wait until the full document is open. Ignored
   * with collaboration. Default false.
   */
  previewFirstPage?: boolean;
  /**
   * Lets the ref's proposal methods run while the editor is read-only or viewing. Typing,
   * `applyEdits`, commands and plugin writes stay blocked. Default: false.
   */
  allowHostProposals?: boolean;
  /** Lists host proposals as sidebar cards while `allowHostProposals` is set; otherwise they stay out of the sidebar. Default: false. */
  showHostProposalsInSidebar?: boolean;
  /**
   * When true, the editor does not intercept Cmd/Ctrl+F or Cmd/Ctrl+H.
   * This lets the browser or host app handle native find/history shortcuts.
   */
  disableFindReplaceShortcuts?: boolean;
  /** Custom toolbar actions */
  toolbarExtra?: ReactNode;
  /**
   * Replaces the built-in chrome: omit it for the default toolbar, pass
   * `null` for none, or pass chrome composed from the toolbar parts. Supplied
   * chrome also renders when `readOnly` is set; `showToolbar={false}` hides both.
   * @experimental
   */
  toolbar?: ReactNode;
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
   * comment threads sync across peers — the document only carries the
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
  /** Receives the editor's rendered-DOM context whenever a new frame or zoom rebuilds it. */
  onRenderedDomContextReady?: (context: RenderedDomContext) => void;
  /**
   * Called once per document, when its first pages are painted on screen, by the first-page
   * preview or the full document, whichever shows first. The document may still be opening.
   */
  onFirstPagePainted?: () => void;
  /**
   * Unmanaged overlay content, drawn under managed plugin overlays.
   * @deprecated Contribute an `overlay` through `plugins` instead.
   */
  pluginOverlays?: ReactNode;
  /**
   * Unmanaged sidebar items, merged with comments and managed plugin items.
   * @deprecated Contribute sidebar items through `plugins` instead.
   */
  pluginSidebarItems?: ReactSidebarItem[];
  /**
   * Geometry for `pluginSidebarItems`; ignored while `plugins` are installed.
   * @deprecated The editor supplies its own geometry.
   */
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
   * Font provider for this editor, overriding whatever `configureDefaultFonts`
   * set globally. With neither, measurement falls back to the browser and may
   * not paginate like Word. Call `configureDefaultFonts` before editors load;
   * existing registries retain their provider, so it is not per editor or tenant.
   */
  measurementFontProvider?: BundledFontProvider;
}

/**
 * DocxEditor ref interface
 */
export interface DocxEditorRef {
  /**
   * The editor's commands, shared by built-in and host chrome.
   * @experimental
   */
  readonly commands: DocxCommandStore;
  /** @deprecated Use {@link readParagraphs} or {@link exportStructuredWithPages}. Throws DocxAsyncOnlyError in worker viewers. */
  getDocument: () => Document | null;
  /** @deprecated The paged editor is internal; use the editor ref's members. Returns null in viewer sessions. */
  getEditorRef: () => PagedEditorRef | null;
  /** Commits accepted input and selection; waits for active IME composition. */
  flushPendingInput: () => Promise<void>;
  /**
   * Flushes pending input, then reads paragraph texts with the version they were read at. Build
   * edit targets and `expectVersion` from this result.
   */
  readParagraphs: (request: DocxReadParagraphsRequest) => Promise<DocxReadParagraphsResult>;
  /** Reads paragraph identities from the current document. */
  getParagraphIdentities: () => Promise<DocxParagraphIdentitySnapshot>;
  /** Resolves paragraph anchors in input order at the current version. */
  resolveParagraphAnchors: (anchors: readonly DocxParagraphAnchor[]) => Promise<{
    version: string;
    results: DocxParagraphAnchorResult[];
  }>;
  /** Flushes pending input, then searches exactly and case-sensitively within one scope. */
  findText: (request: DocxFindTextRequest) => Promise<DocxFindTextResult>;
  /** Flushes pending input, then checks an edit batch without changing anything. */
  validateEdits: (request: DocxEditRequest) => Promise<DocxValidationResult>;
  /**
   * Flushes pending input, then applies every step or none against `expectVersion`. A refusal
   * is returned as data and never rolls back the flushed typing. Read-only editors refuse with
   * `read-only`, and suggesting mode requires `suggest` on every step. Throws when the document
   * is replaced while input is flushing.
   */
  applyEdits: (request: DocxEditRequest) => Promise<DocxEditResult>;
  /**
   * Flushes pending input, then proposes a round of tracked changes grouped by proposal id; see
   * `YrsSession.proposeChanges`. Read-only editors refuse with `read-only` unless
   * `allowHostProposals` is set. Never saves or opens the comments sidebar. Throws when the
   * document is replaced while input is flushing.
   */
  proposeChanges: (request: DocxProposalRequest) => Promise<DocxProposalResult>;
  /**
   * Flushes pending input, then sets how proposals render; see `YrsSession.setProposalStates`.
   * Gated like {@link proposeChanges}.
   */
  setProposalStates: (request: DocxProposalStateRequest) => Promise<DocxProposalResult>;
  /**
   * Flushes pending input, then withdraws proposals, settling each as its decision previews it;
   * see `YrsSession.withdrawProposals`. Gated like {@link proposeChanges}.
   */
  withdrawProposals: (request: DocxProposalWithdrawRequest) => Promise<DocxProposalResult>;
  /** Flushes pending input, then reads the proposals of the loaded document. */
  getProposals: () => Promise<DocxProposalSnapshot>;
  /**
   * Flushes pending input, then lists the document's content controls with the version they were
   * read at. Fill text controls with `setContentControlText` steps through {@link applyEdits}.
   */
  listContentControls: (options?: DocxContentControlsOptions) => Promise<DocxContentControlsResult>;
  /** Flushes pending input, then returns the content controls matching `query` exactly. */
  findContentControls: (
    query: DocxContentControlQuery,
    options?: DocxContentControlsOptions
  ) => Promise<DocxContentControlsResult>;
  /**
   * Flushes pending input, then exports the document as structured content with the page map
   * of the editor's authoritative layout of that version: the physical page, displayed page
   * label and body, header, footer or note occurrence showing each block and inline. The layout
   * must have been computed from the editor's current fonts, measurement defaults, render
   * environment and pagination options; the editor lays the document out, or waits for its
   * fonts, when it was not. The references describe that layout, which a later edit may
   * supersede before it is painted. Refuses as data when no such layout is ready in time, and
   * never lays out again when `expectLayoutVersion` names a layout. Throws when the document is
   * replaced meanwhile. In a viewer session, reads the worker's layout and waits for the layout
   * the editor runs on its own instead of laying out again.
   */
  exportStructuredWithPages: (
    options: DocxPageExportOptions
  ) => Promise<DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>>;
  /** @deprecated Use {@link readPositionAtPoint}. Worker viewers return a cached answer or null while reading it. */
  getPositionAtPoint: (clientX: number, clientY: number) => DocxPointPosition | null;
  /** Reads the text under client coordinates after pending input commits. Worker viewers resolve it against the presented frame, retrying superseded reads. */
  readPositionAtPoint: (clientX: number, clientY: number) => Promise<DocxPointPosition | null>;
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
  /**
   * The document's page count, or 0 until it is laid out in full: a large document paints its
   * first pages before the rest is laid out. See {@link whenLayoutComplete}.
   */
  getTotalPages: () => number;
  /** The editor's wasm memories on the main thread and in its resident worker. */
  getMemoryStats: () => DocxMemoryStats;
  /**
   * Resolves with the page count once the whole document, as it is now, is laid out and its
   * pages are ready to paint. Waits for the layout the editor runs on its own and never asks for
   * one. Rejects when rendering fails, or after `options.timeoutMs` when given.
   * With `experimentalWorkerOpen`, resolves once layout is complete and visible pages are built;
   * pages away from the viewport build when shown.
   * @example const pages = await ref.current?.whenLayoutComplete({ timeoutMs: 60_000 })
   */
  whenLayoutComplete: (options?: { timeoutMs?: number }) => Promise<number>;
  /**
   * Scroll the paginated view so the given page is in view.
   * Page numbers are 1-indexed (matches `getCurrentPage` / `getTotalPages`).
   * No-op for out-of-range or non-integer values. While only the first pages
   * are laid out, a later page waits for the rest.
   * @example ref.current?.scrollToPage(2)
   */
  scrollToPage: (pageNumber: number) => void;
  /** @deprecated Use {@link scrollToParagraph}. Worker viewers start the async navigation and return true. */
  scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  /**
   * Scroll the paginated view to a specific display position.
   * For Word `w14:paraId` use
   * `scrollToParaId` instead.
   * @example ref.current?.scrollToPosition(42)
   */
  scrollToPosition: (displayPosition: number) => void;
  /** @deprecated Use {@link scrollToComment}. Worker viewers start the async navigation and return true. */
  scrollToCommentId: (commentId: number) => boolean;
  /** @deprecated Use {@link scrollToChange}. Worker viewers start the async navigation and return true. */
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
  /** @deprecated Use {@link insertComment}. Returns null in viewer sessions. */
  addComment: (options: DocxCommentInsertion) => number | null;
  /** @deprecated Use {@link insertCommentReply}. Returns null in viewer sessions. */
  replyToComment: (commentId: number, text: string, author: string) => number | null;
  /** Resolve (mark as done) a comment. Does nothing in viewer sessions. */
  resolveComment: (commentId: number) => void;
  /** @deprecated In a viewer session it queues the change through the worker and returns true; use {@link proposeChanges} for the result. */
  proposeChange: (options: {
    paraId: string;
    search: string;
    replaceWith: string;
    author: string;
  }) => boolean;
  /** @deprecated Use {@link findParagraphs}. Throws DocxAsyncOnlyError in worker viewers. */
  findInDocument: (
    query: string,
    options?: { caseSensitive?: boolean; limit?: number }
  ) => Array<{ paraId: string; match: string; before: string; after: string }>;
  /** @deprecated Use {@link commands}; they act on the selection. Returns false in viewer sessions. */
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
  /** @deprecated Use {@link applyEdits} with a `setParagraphStyle` step. Returns false in viewer sessions. */
  setParagraphStyle: (options: { paraId: string; styleId: string }) => boolean;
  /** @deprecated Use {@link commands}; they act on the selection. Returns false in viewer sessions. */
  insertBreak: (options: {
    paraId: string;
    type: 'page' | 'sectionNextPage' | 'sectionContinuous';
  }) => boolean;
  /** @deprecated Use {@link exportStructuredWithPages}. Throws DocxAsyncOnlyError in worker viewers. */
  getPageContent: (pageNumber: number) => {
    pageNumber: number;
    text: string;
    paragraphs: Array<{ paraId: string; text: string; styleId?: string }>;
  } | null;
  /** @deprecated Use {@link readSelectionInfo}. Worker viewers return null. */
  getSelectionInfo: () => DocxSelectionInfo | null;
  /** Get all comments. */
  getComments: () => Comment[];
  /** @deprecated Use {@link onDocumentChange}. Worker viewers do not fire these listeners. */
  onContentChange: (listener: (document: Document) => void) => () => void;
  /** Subscribe to selection changes (cursor moves / selection changes). Returns unsubscribe. */
  onSelectionChange: (listener: (selection: SelectionState | null) => void) => () => void;
  /**
   * Find `query` in the document body, tables included, for a host's own find UI. Flushes
   * pending input, highlights every match, makes the first match on or after the page in view
   * current and scrolls it to the middle of the view. Moves neither the selection nor focus and
   * works read-only. Case-insensitive unless `options.caseSensitive`; an empty query clears.
   * The search runs again when the document changes, keeping the current match.
   */
  search: (query: string, options?: DocxSearchOptions) => Promise<DocxSearchState>;
  /** Make the next match current, wrapping, and scroll to it. Null without a search. */
  searchNext: () => DocxSearchState | null;
  /** Make the previous match current, wrapping, and scroll to it. Null without a search. */
  searchPrevious: () => DocxSearchState | null;
  /** Make match `index` (zero-based, wrapping) current and scroll to it. Null without a search. */
  searchGoTo: (index: number) => DocxSearchState | null;
  /** Remove the search and its highlights. */
  clearSearch: () => void;
  /** The current search, or null. */
  getSearchState: () => DocxSearchState | null;
  /**
   * Subscribe to search changes: a search, a new current match, a re-run after a document change,
   * and clearing (null). Returns unsubscribe.
   */
  onSearchChange: (listener: (state: DocxSearchState | null) => void) => () => void;
  /** Reads the current caret or selection, returning null when there is no selection. */
  readSelectionInfo: () => Promise<DocxSelectionInfo | null>;
  /** Finds paragraphs containing one unique occurrence of the query. */
  findParagraphs: (query: string, options?: { caseSensitive?: boolean; limit?: number }) => Promise<DocxParagraphMatch[]>;
  /** Selects and reveals a paragraph, optionally flashing its text. */
  scrollToParagraph: (paraId: string, options?: ScrollToParaIdOptions) => Promise<boolean>;
  /** Selects and reveals a comment's anchored range, returning false when it no longer exists. */
  scrollToComment: (commentId: number) => Promise<boolean>;
  /** Selects and reveals a revision's range, returning false when it no longer exists. */
  scrollToChange: (revisionId: number) => Promise<boolean>;
  /** Inserts an anchored comment after pending input commits. Returns null in viewer sessions. */
  insertComment: (options: DocxCommentInsertion) => Promise<number | null>;
  /** Adds a reply to an existing comment. Returns null in viewer sessions. */
  insertCommentReply: (commentId: number, text: string, author: string) => Promise<number | null>;
  /** Subscribes to committed document versions and returns an unsubscribe function. */
  onDocumentChange: (listener: (change: DocxDocumentChange) => void) => () => void;
}

/**
 * Editor internal state
 */
interface EditorState {
  isLoading: boolean;
  parseError: string | null;
  zoom: number;
  /** Paragraph indent data for ruler */
  paragraphIndentLeft: number;
  paragraphIndentRight: number;
  paragraphFirstLineIndent: number;
  paragraphHangingIndent: boolean;
  paragraphTabs: import('@betteroffice/docx/types/document').TabStop[] | null;
  /** Table context for showing the table toolbar. */
  pmTableContext: TableContextInfo | null;
}

/** The image an image dialog edits, captured when it opened. */
interface ImageDialogTarget {
  pos: number;
  handle: PagedEditorImageHandle | null;
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
}

function imageDialogTarget(
  image: PagedEditorSelectedImage,
  handle: PagedEditorImageHandle | null
): ImageDialogTarget {
  const { attrs } = image;
  const text = (value: unknown) => (typeof value === 'string' ? value : null);
  const number = (value: unknown) =>
    typeof value === 'number' && Number.isFinite(value) ? value : null;
  return {
    pos: image.pos,
    handle,
    wrapType: text(attrs.wrapType) ?? 'inline',
    displayMode: text(attrs.displayMode) ?? 'inline',
    cssFloat: text(attrs.cssFloat),
    transform: text(attrs.transform),
    alt: text(attrs.alt),
    borderWidth: number(attrs.borderWidth),
    borderColor: text(attrs.borderColor),
    borderStyle: text(attrs.borderStyle),
    width: number(attrs.width),
    height: number(attrs.height),
  };
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

/** Sidebar anchor keys of host proposals' revisions, which never open the sidebar themselves. */
function proposalAnchorKeys(session: YrsSession | null): Set<string> {
  const keys = new Set<string>();
  for (const proposal of session?.getProposals().proposals ?? []) {
    for (const revisionId of proposal.revisionIds) {
      keys.add(`revision-${yrsIdToNumericId(revisionId)}`);
    }
  }
  return keys;
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
    experimentalPrewarm = false,
    document: initialDocument,
    onSave,
    onSaveRequest,
    downloadOnSave = true,
    collaboration,
    experimentalWorkerOpen = false,
    mediaTokens,
    onOpen,
    author = 'User',
    onChange,
    onDocumentChange,
    onSelectionChange,
    onError,
    onMemoryPressure,
    memoryBudget,
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
    showHiddenText = false,
    readOnly: readOnlyProp = false,
    previewFirstPage = false,
    allowHostProposals = false,
    showHostProposalsInSidebar = false,
    disableFindReplaceShortcuts = false,
    toolbarExtra,
    toolbar,
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
    onFirstPagePainted,
    pluginOverlays,
    pluginSidebarItems,
    pluginRenderedDomContext,
    plugins,
    pluginGrants,
    onPluginError,
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
  useDocxEnginePrewarm(experimentalPrewarm);
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
    paragraphIndentLeft: 0,
    paragraphIndentRight: 0,
    paragraphFirstLineIndent: 0,
    paragraphHangingIndent: false,
    paragraphTabs: null,
    pmTableContext: null,
  });
  const [imageTarget, setImageTarget] = useState<ImageDialogTarget | null>(null);

  const isDark = useIsDark(colorMode);

  // The one non-body part open for editing — a header/footer band or a note.
  const [partEditTarget, setPartEditTarget] = useState<PartEditTarget | null>(null);

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
  const fontScope = useFontLoadScope();

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
  const handoffFromRef = useRef<YrsSession | null>(null);
  const canvasRenderer = useCanvasRenderer(
    rustFontChainsProviderRef,
    resolvedIdsForRender,
    () => pagedEditorRef.current?.relayout(),
    memoryBudget?.workerLimitBytes,
    handoffFromRef,
    experimentalWorkerOpen
  );
  // The full session failing to lay out or render as it opens fails the
  // load, which reports it. Each render error is handled once: one the
  // preview left set is not the full session's.
  const failOpeningRef = useRef<(error: Error, session?: unknown) => boolean>(() => false);
  const notifiedErrorsRef = useRef(new WeakSet<Error>());
  const notifyError = useCallback((error: Error) => {
    if (notifiedErrorsRef.current.has(error)) return;
    notifiedErrorsRef.current.add(error);
    onError?.(error);
  }, [onError]);
  const handledRenderErrorRef = useRef<Error | null>(null);
  const renderErrorEngine = canvasRenderer.errorEngine ?? undefined;
  const coreSessionRef = useRef<unknown>(null);
  // A worker-opened session the editor has not taken yet fails through its open.
  const untakenWorkerSession = useCallback(
    (session?: unknown) =>
      experimentalWorkerOpen &&
      session != null &&
      session !== coreSessionRef.current &&
      (session as YrsSession).isDisplayOnly?.() !== true,
    [experimentalWorkerOpen]
  );
  useEffect(() => {
    const error = canvasRenderer.error;
    if (!error || error === handledRenderErrorRef.current) return;
    handledRenderErrorRef.current = error;
    if (untakenWorkerSession(renderErrorEngine)) return;
    if (!failOpeningRef.current(error, renderErrorEngine)) notifyError(error);
  }, [canvasRenderer.error, renderErrorEngine, notifyError, untakenWorkerSession]);
  useMemoryPressure(onMemoryPressure, memoryBudget, canvasRenderer.workerMemory, [
    canvasRenderer.frame,
    canvasRenderer.error,
  ]);

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
  const modeReadOnly = readOnlyProp || editingMode === 'viewing';
  const commandBridgeRef = useRef<PagedEditorCommandBridge | null>(null);
  const writeModeRef = useRef<EditorMode>(editingMode);
  writeModeRef.current = modeReadOnly ? 'viewing' : editingMode;
  const allowHostProposalsRef = useRef(allowHostProposals);
  allowHostProposalsRef.current = allowHostProposals;

  // Bridge / agent event subscribers — fan-out from the existing onChange and
  // onSelectionChange paths so multiple listeners (host app, MCP server, etc.)
  // can observe edits without competing for the single React prop.
  const contentChangeSubscribersRef = useRef(new Set<(doc: Document) => void>());
  const documentChangeSubscribersRef = useRef(new Set<(change: DocxDocumentChange) => void>());
  const onDocumentChangeRef = useRef(onDocumentChange);
  onDocumentChangeRef.current = onDocumentChange;
  const [contentSubscriberCount, setContentSubscriberCount] = useState(0);
  const selectionChangeSubscribersRef = useRef(new Set<(s: SelectionState | null) => void>());
  const legacyProjectionTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // History hook for undo/redo - start with null document
  const history = useDocumentHistory<Document | null>(initialDocument || null, {
    maxEntries: 100,
    groupingInterval: 500,
  });

  // Refs (pagedEditorRef is declared earlier — useCommentManagement needs it)
  const containerRef = useRef<HTMLDivElement>(null);
  const editorContentRef = useRef<HTMLDivElement>(null);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const [documentFonts, setDocumentFonts] = useState<FontOption[]>([]);
  // The loader needs outline setters before it identifies the viewer session.
  const viewerOutlineRef = useRef(false);
  const [viewerCommentRanges, setViewerCommentRanges] = useState<ViewerCommentRanges>(new Map());
  const {
    showOutline,
    setShowOutline,
    showOutlineRef,
    outlineHeadings,
    setHeadingInfos,
    refreshHeadings,
    navigateViewerHeading,
    editorScrollLeft,
  } = useOutlineSidebar({
    showOutlineProp,
    viewerRef: viewerOutlineRef,
    viewerRead: canvasRenderer.readWorkerDocument,
    queries: canvasRenderer.queries,
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

  const { focusActiveEditor } = useActiveEditor({
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

  const beginPluginLoadRef = useRef<() => void>(() => {});
  const { resetForNewDocument: resetEditorState } = useResetEditorState({
    commentsLoadedRef,
    trackedChangesLoadedRef,
    sidebarAutoOpenedRef,
    setComments,
    setYrsTrackedChangesResult,
    setHeadingInfos,
    setShowCommentsSidebar,
    setIsAddingComment,
    setCommentSelectionRange,
    setAddCommentYPosition,
    setFloatingCommentBtn,
    setPartEditTarget,
    setAnchorPositions,
    clearFindReplaceMatches: useCallback(() => findReplace.setMatches([], 0), [findReplace]),
    cleanOrphanedCommentsTimerRef,
  });
  const { resetSettled, awaitingDocument } = canvasRenderer;
  const shownListRef = useRef(canvasRenderer.displayList);
  // Pages shown while a load is under way belong to the document it replaces.
  const replacedListsRef = useRef(new WeakSet<object>());
  if (canvasRenderer.displayList && awaitingDocument()) {
    replacedListsRef.current.add(canvasRenderer.displayList);
  }
  const firstPagePendingRef = useRef(true);
  const firstPageGenerationRef = useRef(0);
  const onFirstPagePaintedRef = useRef(onFirstPagePainted);
  useLayoutEffect(() => {
    shownListRef.current = canvasRenderer.displayList;
    onFirstPagePaintedRef.current = onFirstPagePainted;
  });
  const resetForNewDocument = useCallback(() => {
    beginPluginLoadRef.current();
    if (shownListRef.current) replacedListsRef.current.add(shownListRef.current);
    firstPagePendingRef.current = true;
    firstPageGenerationRef.current += 1;
    resetEditorState();
    setViewerCommentRanges(new Map());
    resetSettled();
  }, [resetEditorState, resetSettled]);

  const {
    loadParsedDocument,
    loadBuffer,
    yrsSeedDocument,
    yrsSeedBytes,
    yrsSeedGeneration,
    isCurrentLoad,
    acceptHostDocument,
    failHostDocument,
    reportLayoutError: reportDocumentLayoutError,
    fontAliases,
  } = useDocumentLoader({
    documentBuffer,
    initialDocument,
    externalContent: false,
    history,
    pagedEditorRef,
    setLoadingState: useCallback(
      (s: { isLoading: boolean; parseError: string | null }) => {
        setState((prev) => ({ ...prev, isLoading: s.isLoading, parseError: s.parseError }));
        // Each failed load fails the wait, also one repeating the previous message.
        if (s.parseError !== null) resetSettled(new Error(s.parseError));
      },
      [resetSettled]
    ),
    setComments,
    setShowCommentsSidebar,
    onError: notifyError,
    resetForNewDocument,
    commentsLoadedRef,
    commentIdAllocator: commentIdAllocatorRef.current,
    setDocumentFonts,
    fontScope,
  });

  // A layout error of the session a newer load replaced is not the loaded
  // document's: the old session may already be freed.
  const sessionGenerationRef = useRef<number | null>(null);
  const reportLayoutError = useCallback(
    (error: Error, session?: unknown) => {
      if (sessionGenerationRef.current !== yrsSeedGeneration) return;
      // A display-only preview's error fails no wait: the full session replaces it.
      const preview = (session as YrsSession | undefined)?.isDisplayOnly?.() === true;
      reportDocumentLayoutError(error, preview ? undefined : resetSettled);
    },
    [reportDocumentLayoutError, resetSettled, yrsSeedGeneration]
  );
  useDocxEnginePrewarmOnBytes(experimentalPrewarm, yrsSeedBytes);
  // A read-only worker-open document keeps host proposals in the worker until the replica loads.
  const workerProposals = modeReadOnly && !collaboration;
  // A viewer session holds no document here: selection, copy and point reads go to the worker.
  const viewerSession = useViewerSession(Boolean(experimentalWorkerOpen), workerProposals, yrsSeedGeneration);
  // Hit testing answers from the first painted page once the query engine has loaded.
  useEffect(() => {
    if (viewerSession) void loadRustDisplayListQueryEngine().catch(() => {});
  }, [viewerSession]);
  const workerContentChangeRef = useRef<() => void>(() => {});
  const workerRevisionsRef = useRef<() => void>(() => {});
  const yrsCore = useYrsCoreSession(
    true,
    history.state,
    yrsSeedDocument,
    yrsSeedBytes,
    yrsSeedGeneration,
    collaboration,
    {
      isCurrentLoad,
      onSession: canvasRenderer.recordSession,
      onHostDocument: acceptHostDocument,
      onError: failHostDocument,
      onReplicaError: (error, generation) => {
        if (isCurrentLoad(generation)) reportLayoutError(error);
      },
    },
    {
      previewFirstPage,
      heldEngine: canvasRenderer.layoutEngine,
      shownEngine: canvasRenderer.presentedEngine,
      workerOpen: experimentalWorkerOpen
        ? {
            openInWorker: canvasRenderer.openInWorker,
            openPreviewInWorker: canvasRenderer.openPreviewInWorker,
            workerProposals,
            refreshWorkerLayout: () => pagedEditorRef.current?.refreshWorkerLayout(),
            renderedFrame: canvasRenderer.status === 'ready' ? canvasRenderer.displayList : null,
            pendingCompletion: canvasRenderer.pendingCompletion,
            hydrateOnDemand: workerProposals,
            onWorkerContentChange: () => workerContentChangeRef.current(),
            onWorkerRevisions: () => workerRevisionsRef.current(),
          }
        : undefined,
      mediaTokens,
    }
  );
  // A viewer whose document fell back to this thread reads the copy it holds here.
  const viewerReads = viewerSession && viewerReadsWorker(canvasRenderer.queries, yrsCore.session);
  viewerOutlineRef.current = viewerReads;
  // Until the full session's pages are shown, the editor takes no input and its
  // API and commands see a document that is still loading.
  const opening = yrsCore.opening;
  failOpeningRef.current = yrsCore.failOpening;
  coreSessionRef.current = yrsCore.session;
  const hostProposalRevisions = useHostProposalRevisions(yrsCore.session);
  const hostProposalRevisionsRef = useRef(hostProposalRevisions);
  hostProposalRevisionsRef.current = hostProposalRevisions;
  const reportPagedError = useCallback(
    (error: Error, session?: unknown) => {
      if (error instanceof SupersededPreviewError || untakenWorkerSession(session)) return;
      if (!failOpeningRef.current(error, session)) reportLayoutError(error, session);
    },
    [untakenWorkerSession, reportLayoutError]
  );
  const readOnly = modeReadOnly || opening;
  if (opening) writeModeRef.current = 'viewing';
  const openingRef = useRef(opening);
  openingRef.current = opening;
  handoffFromRef.current = yrsCore.handoffFrom;
  const { notifyFramePresented } = yrsCore;
  const shownRef = useRef({
    displayList: canvasRenderer.displayList,
    engine: canvasRenderer.presentedEngine,
  });
  shownRef.current = {
    displayList: canvasRenderer.displayList,
    engine: canvasRenderer.presentedEngine,
  };
  // Layout cleanup runs in the unmount commit, before a queued frame could fire the callback.
  useLayoutEffect(
    () => () => {
      firstPageGenerationRef.current += 1;
    },
    []
  );
  useEffect(() => {
    const offPresented = onPresented((displayList, options) => {
      const shown = shownRef.current;
      if (displayList !== shown.displayList) return;
      if (shown.engine) notifyFramePresented(shown.engine);
      if (
        firstPagePendingRef.current &&
        !awaitingDocument() &&
        !replacedListsRef.current.has(displayList) &&
        shown.displayList?.pages.some((page) => page.unbuilt !== true)
      ) {
        firstPagePendingRef.current = false;
        // The callback of the document whose pages presented, not of one committed since.
        const callback = onFirstPagePaintedRef.current;
        const generation = firstPageGenerationRef.current;
        const fire = () => {
          if (generation !== firstPageGenerationRef.current) return;
          try {
            callback?.();
          } catch (error) {
            console.error('[DocxEditor] onFirstPagePainted threw', error);
          }
        };
        if (options?.worker) requestAnimationFrame(() => requestAnimationFrame(fire));
        else fire();
      }
    });
    // Pages of the opening session that fail to paint fail the load, as its render errors do.
    const offFailed = onReplayFailed((displayList, error) => {
      const shown = shownRef.current;
      if (displayList !== shown.displayList || !shown.engine) return;
      failOpeningRef.current(
        error instanceof Error ? error : new Error(String(error)),
        shown.engine
      );
    });
    return () => {
      offPresented();
      offFailed();
    };
  }, [awaitingDocument, notifyFramePresented]);
  sessionGenerationRef.current = yrsCore.sessionGeneration;
  // Content listeners project the document on every edit; warm its base once
  // the first pages are on screen so neither opening nor the first key pays.
  useCompatibilityWarm(
    yrsCore.session,
    canvasRenderer.status === 'ready' ? canvasRenderer.displayList : null,
    yrsCore.replicaReady && (Boolean(onChange) || contentSubscriberCount > 0),
    yrsCore.scheduleCompatibilityWarm,
    yrsCore.cancelCompatibilityWarm
  );

  const {
    imageInputRef,
    docxInputRef,
    handleSave,
    reservePrint,
    handleDownloadDocument,
    handleOpenDocument,
    handleDocxFileChange,
    handleInsertImageClick,
    handleImageFileChange,
  } = useFileIO({
    pagedEditorRef,
    viewerSession,
    resolveImage: canvasRenderer.resolveImage,
    shownImageResolver: canvasRenderer.imageResolverForShownFrame,
    fontFamilies: fontAliases,
    comments,
    documentName,
    onSave,
    onSaveRequest,
    downloadOnSave,
    onOpen,
    onError,
    onPrint,
    onDocumentNameChange,
    loadBuffer,
    focusActiveEditor,
  });

  const handleZoomChange = useCallback((zoom: number) => {
    setState((prev) => ({ ...prev, zoom }));
  }, []);

  const commands = useDocxCommandBinding({
    experimentalWorkerOpen,
    pagedEditorRef,
    bridgeRef: commandBridgeRef,
    isLoading: state.isLoading || opening,
    parseError: state.parseError,
    document: history.state,
    session: yrsCore.session,
    viewerSession,
    readOnly: readOnlyProp || opening,
    mode: editingMode,
    modeControlled: modeProp !== undefined,
    onModeChange,
    setEditingMode,
    sidebarOpen: showCommentsSidebar,
    sidebarControlled: commentsSidebarOpen !== undefined,
    sidebarHasSetter: onCommentsSidebarOpenChange !== undefined,
    setShowCommentsSidebar,
    setExpandedSidebarItem,
    zoom: state.zoom,
    setZoom: handleZoomChange,
    showFileOpen,
    showHelpMenu,
    partEditing: partEditTarget !== null,
    fontFamilies,
    documentFonts,
    theme: history.state?.package.theme ?? theme ?? null,
    i18n,
    isDark,
    displayListQueries: canvasRenderer.queries,
    getCachedStyleResolver,
    hyperlinkDialog,
    findReplace,
    save: handleDownloadDocument,
    reservePrint,
    renderedDisplayList: () =>
      canvasRenderer.settledDisplayList(() => pagedEditorRef.current?.relayout()),
    openDocument: handleOpenDocument,
    pickImage: handleInsertImageClick,
    tableAction: (action) => handleTableAction(action),
    openImageProperties: (image) => {
      setImageTarget(
        imageDialogTarget(image, commandBridgeRef.current?.imageHandle(image.pos) ?? null)
      );
      setImagePropsOpen(true);
    },
    openPageSetup: () => handleOpenPageSetup(),
    openWatermark: () => handleOpenWatermark(),
    refreshTrackedChanges: (session) => refreshTrackedChanges(session),
  });
  const commandController = commands.controller;
  const runBridgeCommand = useCallback(
    (command: Parameters<PagedEditorCommandBridge['command']>[0]) =>
      commandOutcome(commandBridgeRef.current?.command(command) ?? false),
    []
  );

  const getProposalAnchorKeys = useCallback(
    () => proposalAnchorKeys(pagedEditorRef.current?.getYrsSession() ?? null),
    []
  );

  // Auto-open the sidebar once if the loaded document already has tracked changes.
  useCommentLifecycle({
    commentToRevision,
    setComments,
    isLoading: state.isLoading || opening,
    trackedChanges,
    getProposalAnchorKeys,
    setShowCommentsSidebar,
    trackedChangesLoadedRef,
  });

  useFontLifecycle(fonts, onFontsLoadedCallback, onError, fontScope);

  const pushDocument = useCallback(
    (document: Document) => {
      history.push(document);
      return document;
    },
    [history]
  );

  const notifiedVersionRef = useRef<string | null>(null);
  const notifyDocumentVersion = useCallback((version: string) => {
    if (notifiedVersionRef.current === version) return;
    notifiedVersionRef.current = version;
    const change = { version };
    for (const listener of [onDocumentChangeRef.current, ...documentChangeSubscribersRef.current]) {
      try {
        listener?.(change);
      } catch (error) {
        console.error('documentChange listener threw:', error);
      }
    }
  }, []);
  const presentedVersionRef = useRef<{ generation: number | null; version: string | null }>({ generation: null, version: null });
  const observePresentedVersion = useCallback((displayList: object) => {
    if (!isWorkerViewer(pagedEditorRef.current) || !isPresented(canvasRenderer.canvasHostRef.current, displayList)) return;
    const version = workerFrameVersionOf(displayList);
    if (version === null) return;
    const previous = presentedVersionRef.current;
    presentedVersionRef.current = { generation: yrsSeedGeneration, version };
    if (previous.generation === yrsSeedGeneration && previous.version !== null && previous.version !== version) {
      notifyDocumentVersion(version);
    }
  }, [canvasRenderer.canvasHostRef, notifyDocumentVersion, yrsSeedGeneration]);
  useEffect(() => {
    if (canvasRenderer.displayList) observePresentedVersion(canvasRenderer.displayList);
    return onPresented(observePresentedVersion);
  }, [canvasRenderer.displayList, observePresentedVersion]);

  const notifyDocumentChange = useCallback(
    (document: Document) => {
      if (isWorkerViewer(pagedEditorRef.current)) {
        if (onChange) warnDeprecatedViewerMember('onChange', 'does not fire in viewer sessions', 'onDocumentChange');
        return;
      }
      onChange?.(document);
      for (const callback of contentChangeSubscribersRef.current) {
        try {
          callback(document);
        } catch (error) {
          console.error('contentChange subscriber threw:', error);
        }
      }
    },
    [onChange]
  );

  const handleContentHousekeeping = useCallback(() => {
    if (showOutlineRef.current) refreshHeadings();
    if (cleanOrphanedCommentsTimerRef.current) {
      clearTimeout(cleanOrphanedCommentsTimerRef.current);
    }
    if (!viewerReads) cleanOrphanedCommentsTimerRef.current = setTimeout(cleanOrphanedComments, 300);
  }, [cleanOrphanedComments, refreshHeadings, showOutlineRef, viewerReads]);

  const scheduleLegacyProjection = useCallback((project: () => void) => {
    if (legacyProjectionTimerRef.current !== null) {
      clearTimeout(legacyProjectionTimerRef.current);
    }
    legacyProjectionTimerRef.current = setTimeout(() => {
      legacyProjectionTimerRef.current = null;
      project();
    }, LEGACY_PROJECTION_DELAY_MS);
  }, []);

  useEffect(
    () => () => {
      if (legacyProjectionTimerRef.current !== null) {
        clearTimeout(legacyProjectionTimerRef.current);
        legacyProjectionTimerRef.current = null;
      }
    },
    []
  );

  const handleDocumentChange = useCallback(
    (newDocument: Document) => {
      commitLegacyDocumentChange(
        newDocument,
        yrsCore.documentFromYrs,
        {
          push: pushDocument,
          notify:
            onChange || contentChangeSubscribersRef.current.size > 0
              ? notifyDocumentChange
              : undefined,
        },
        scheduleLegacyProjection
      );
      if (!isWorkerViewer(pagedEditorRef.current) && yrsCore.session) notifyDocumentVersion(yrsCore.session.version());
      handleContentHousekeeping();
    },
    [
      handleContentHousekeeping,
      notifyDocumentChange,
      notifyDocumentVersion,
      yrsCore.session,
      onChange,
      pushDocument,
      scheduleLegacyProjection,
      yrsCore.documentFromYrs,
    ]
  );

  const projectYrsContentChange = useCallback(() => {
    if (isWorkerViewer(pagedEditorRef.current)) {
      if (onChange) warnDeprecatedViewerMember('onChange', 'does not fire in viewer sessions', 'onDocumentChange');
      return;
    }
    if (onChange || contentChangeSubscribersRef.current.size > 0) {
      commitYrsDocumentChange(yrsCore.documentFromYrs, {
        push: pushDocument,
        notify: notifyDocumentChange,
      });
    }
    handleContentHousekeeping();
  }, [
    handleContentHousekeeping,
    notifyDocumentChange,
    onChange,
    pushDocument,
    yrsCore.documentFromYrs,
  ]);
  const handleYrsContentChange = useCallback(() => {
    if (!isWorkerViewer(pagedEditorRef.current) && yrsCore.session) {
      notifyDocumentVersion(yrsCore.session.version());
    }
    projectYrsContentChange();
  }, [notifyDocumentVersion, projectYrsContentChange, yrsCore.session]);
  // A worker-held change reaches document listeners once the replica holds it; without them,
  // nothing needs the replica.
  workerContentChangeRef.current = () => {
    if (isWorkerViewer(pagedEditorRef.current)) {
      if (onChange) warnDeprecatedViewerMember('onChange', 'does not fire in viewer sessions', 'onDocumentChange');
      return;
    }
    const session = yrsCore.session;
    if (!session) return;
    notifyDocumentVersion(session.version());
    if (!onChange && contentChangeSubscribersRef.current.size === 0) return;
    void requestWorkerOpenReplica(session)?.then(
      () => {
        if (coreSessionRef.current === session) projectYrsContentChange();
      },
      () => {}
    );
  };

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
    partEditOpen: partEditTarget !== null,
    readOnly,
    isLoading: state.isLoading,
    zoom: state.zoom,
    canvasHostRef: canvasRenderer.canvasHostRef,
    displayListQueries: canvasRenderer.queries,
  });

  const { handleYrsSelectionChange, handleViewerSelectionChange } = useSelectionTracker({
    borderSpecRef,
    theme,
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
    commands: commandController,
    pagedEditorRef,
    containerRef,
    disableFindReplaceShortcuts,
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
    if (viewerReads) {
      navigateViewerHeading(pmPos);
      return;
    }
    pagedEditorRef.current?.scrollToPosition(pmPos);
    // Also set selection to the heading
    pagedEditorRef.current?.setSelection(pmPos + 1);
    pagedEditorRef.current?.focus();
  }, [viewerReads, navigateViewerHeading]);

  // Handle shape insertion
  // Handle image wrap type change
  const {
    imagePositionOpen,
    setImagePositionOpen,
    imagePropsOpen,
    setImagePropsOpen,
    footnotePropsOpen,
    setFootnotePropsOpen,
    handleApplyImagePosition,
    handleApplyImageProperties,
    handleApplyFootnoteProperties,
  } = useImageActions({
    document: history.state,
    pmImageContext: imageTarget,
    applyGeometry: (patch) => {
      const handle = imageTarget?.handle ?? null;
      void commands
        .complete('imageProperties', () => {
          const pos = handle ? commandBridgeRef.current?.imagePosition(handle) : null;
          if (pos == null) throw new DocxCommandAdmissionError('target-changed');
          return runBridgeCommand({ type: 'imageGeometry', pmPos: pos, patch });
        })
        .then((result) => {
          if (result.ok && result.status === 'executed') focusActiveEditor();
        });
    },
    pushDocument,
  });

  const {
    tablePropsOpen,
    setTablePropsOpen,
    currentTableProperties,
    handleTablePropertiesApply,
    splitCellDialogState,
    handleTableAction,
    handleSplitCellDialogClose,
    handleSplitCellDialogApply,
  } = useTableDialogs({
    pagedEditorRef,
    borderSpecRef,
    apply: (command) => commandBridgeRef.current?.command(command) ?? false,
    complete: (dialog, command) => {
      void commands.complete(dialog, () => runBridgeCommand(command));
    },
  });

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
    openPopup: useCallback(() => commands.open('linkPopup'), [commands]),
    applyCommand: (source, command) => {
      void commands.complete(source, () => runBridgeCommand(command));
    },
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
    runTableAction: (action) => void commandController.store.execute('tableAction', action),
    editorContentRef,
    displayListQueries: canvasRenderer.queries,
    interactionPageHostRef: canvasRenderer.canvasHostRef,
    i18n,
    partEditOpen: partEditTarget !== null,
    readOnly,
    onAddComment: useCallback(
      ({ from, to, yPos }: { from: number; to: number; yPos: number | null }) => {
        setCommentSelectionRange({ from, to });
        setAddCommentYPosition(yPos === null ? null : yPos / state.zoom);
        setShowCommentsSidebar(true);
        setIsAddingComment(true);
        setFloatingCommentBtn(null);
      },
      [state.zoom]
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
  const dialogApply = useRef({ pageSetup: handlePageSetupApply, watermark: handleWatermarkApply });
  dialogApply.current = { pageSetup: handlePageSetupApply, watermark: handleWatermarkApply };

  const { scrollPageInfo, setScrollPageInfo, readCurrentPage } = useScrollPageInfo({
    scrollContainerRef,
    pagedEditorRef,
  });
  // The error view unmounts the pages before they report none: a failed load,
  // a preview's included, keeps nothing of what they showed.
  const resetCanvasRenderer = canvasRenderer.reset;
  useEffect(() => {
    if (!state.parseError) return;
    resetCanvasRenderer();
    setScrollPageInfo((prev) => (prev.totalPages === 0 ? prev : { ...prev, totalPages: 0 }));
  }, [state.parseError, resetCanvasRenderer, setScrollPageInfo]);

  const pluginOverlayTarget = useCanvasOverlayTarget((plugins?.length ?? 0) > 0, editorContentRef);
  const pluginHost = useDocxPluginHost({
    plugins,
    pluginGrants,
    onPluginError,
    pagedEditorRef,
    writeModeRef,
    mode: editingMode,
    readOnly,
    commands: commandController,
    viewerSelection: viewerSession && !(
      canvasRenderer.queries &&
      presentedWorkerVersion(canvasRenderer.queries) === null &&
      yrsCore.session &&
      !workerOpenReplicaPending(yrsCore.session)
    ),
    session:
      yrsCore.session &&
      !opening &&
      (yrsCore.replicaReady || yrsCore.workerProposalsReady) &&
      yrsCore.sessionGeneration === yrsSeedGeneration &&
      history.state &&
      !state.isLoading &&
      !state.parseError
        ? yrsCore.session
        : null,
    loadGeneration: yrsSeedGeneration,
    queries: canvasRenderer.queries,
    viewerDocumentRead: viewerReads ? canvasRenderer.readWorkerDocument : undefined,
    layoutError: canvasRenderer.error,
    zoom: state.zoom,
    canvasHostRef: canvasRenderer.canvasHostRef,
    overlayTarget: pluginOverlayTarget,
    selectionChangeSubscribersRef,
    i18n,
    onRenderedDomContextReady,
  });
  beginPluginLoadRef.current = pluginHost.beginLoad;
  const sidebarDomContext = pluginHost.managed
    ? pluginHost.renderedDomContext
    : (pluginRenderedDomContext ?? null);
  useEffect(() => {
    if (pluginHost.managed && pluginRenderedDomContext) {
      console.warn(
        '[DocxEditor] pluginRenderedDomContext is ignored while plugins are installed; the editor supplies its own geometry.'
      );
    }
  }, [pluginHost.managed, pluginRenderedDomContext]);

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
    complete: (write) => commands.complete('replace', write),
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

  // Header/footer items have no body range rectangles.
  const canvasBrightenRange = useMemo<CanvasBrightenRange | null>(() => {
    if (!canvasRenderer.queries || !expandedSidebarItem) return null;
    if (expandedSidebarItem.startsWith('comment-')) {
      const id = parseInt(expandedSidebarItem.slice('comment-'.length), 10);
      if (!Number.isFinite(id)) return null;
      if (viewerReads) {
        const range = viewerCommentRanges.get(id);
        return range ? { ...range, variant: 'comment' } : null;
      }
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
  }, [canvasRenderer.queries, expandedSidebarItem, trackedChanges, viewerReads, viewerCommentRanges]);

  // Expose ref methods
  const hostSearch = useHostSearch({
    pagedEditorRef,
    displayListQueries: canvasRenderer.queries,
    canvasHostRef: canvasRenderer.canvasHostRef,
  });

  useDocxEditorRefApi({
    experimentalWorkerOpen,
    viewerSession,
    ref,
    document: history.state,
    documentFromYrs: yrsCore.documentFromYrs,
    historyStateRef,
    pagedEditorRef,
    handleSave,
    zoom: state.zoom,
    setZoom: (zoom: number) => setState((prev) => ({ ...prev, zoom })),
    scrollPageInfo,
    readCurrentPage,
    loadParsedDocument,
    loadBuffer,
    comments,
    setComments,
    setShowCommentsSidebar,
    contentChangeSubscribersRef,
    documentChangeSubscribersRef,
    onContentSubscribersChange: setContentSubscriberCount,
    selectionChangeSubscribersRef,
    getCachedStyleResolver,
    commentIdAllocator: commentIdAllocatorRef.current,
    commands: commandController.store,
    modeRef: writeModeRef,
    openingRef,
    allowHostProposalsRef,
    workerMemory: canvasRenderer.workerMemory,
    settledDisplayList: canvasRenderer.settledDisplayList,
    awaitingDocument,
    hostSearch: hostSearch.api,
  });

  const initialSectionProperties = useMemo(
    () => getInitialSectionProperties(history.state),
    [history.state]
  );
  const {
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    finalSectionProperties,
    handleHeaderFooterDoubleClick,
    handleBodyClick,
    handleRemoveHeaderFooter,
  } = useHeaderFooterEditing({
    document: history.state,
    pushDocument,
    partEditTarget,
    setPartEditTarget,
    readOnly,
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
        const main = () => {
          try {
            session.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
            editor?.syncYrsInputState(true);
          } catch {}
        };
        const authority = viewerReads ? registeredWorkerProposalAuthority(session) : null;
        if (authority) void authority.removeComment(String(id), main).catch(() => {});
        else main();
      }
      if (target) onCommentDelete?.(target);
    },
    onAddComment: (addText) => {
      const comment = createComment(commentIdAllocatorRef.current, addText, author);
      const editor = pagedEditorRef.current;
      const session = viewerReads ? null : editor?.getYrsSession();
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
      if (readOnly) return;
      const editor = pagedEditorRef.current;
      const session = viewerReads ? null : editor?.getYrsSession();
      const range = editor ? displayRangeToYrsRange(editor, from, to) : null;
      if (!session || !range) return;
      session.acceptChange(range);
      editor?.syncYrsInputState(true);
      refreshTrackedChanges(session);
    },
    onRejectChange: (from, to) => {
      if (readOnly) return;
      const editor = pagedEditorRef.current;
      const session = viewerReads ? null : editor?.getYrsSession();
      const range = editor ? displayRangeToYrsRange(editor, from, to) : null;
      if (!session || !range) return;
      session.rejectChange(range);
      editor?.syncYrsInputState(true);
      refreshTrackedChanges(session);
    },
    onAcceptChangeById: (revisionId) => {
      if (viewerSession) return;
      const revision = pagedEditorRef.current
        ?.getYrsSession()
        ?.listRevisions()
        .find((candidate) => yrsIdToNumericId(candidate.revisionId) === revisionId);
      if (revision) {
        void commandController.store.execute('reviewAccept', { revisionId: revision.revisionId });
      }
    },
    onRejectChangeById: (revisionId) => {
      if (viewerSession) return;
      const revision = pagedEditorRef.current
        ?.getYrsSession()
        ?.listRevisions()
        .find((candidate) => yrsIdToNumericId(candidate.revisionId) === revisionId);
      if (revision) {
        void commandController.store.execute('reviewReject', { revisionId: revision.revisionId });
      }
    },
    onTrackedChangeReply: (revisionId, text) => {
      setComments((prev) => [
        ...prev,
        createComment(commentIdAllocatorRef.current, text, author, revisionId),
      ]);
    },
  };

  // Stable callbacks wrapper that delegates to ref (avoids recreating items on every render).
  // Comments do not change while the document opens.
  const stableCallbacks = useMemo<CommentCallbacks>(
    () => ({
      onCommentReply: (...args) => {
        if (!openingRef.current) commentCallbacksRef.current.onCommentReply?.(...args);
      },
      onCommentResolve: (...args) => {
        if (!openingRef.current) commentCallbacksRef.current.onCommentResolve?.(...args);
      },
      onCommentUnresolve: (...args) => {
        if (!openingRef.current) commentCallbacksRef.current.onCommentUnresolve?.(...args);
      },
      onCommentDelete: (...args) => {
        if (!openingRef.current) commentCallbacksRef.current.onCommentDelete?.(...args);
      },
      onAddComment: (...args) => {
        if (!openingRef.current) commentCallbacksRef.current.onAddComment?.(...args);
      },
      onCancelAddComment: (...args) => commentCallbacksRef.current.onCancelAddComment?.(...args),
      onAcceptChange: (...args) => commentCallbacksRef.current.onAcceptChange?.(...args),
      onRejectChange: (...args) => commentCallbacksRef.current.onRejectChange?.(...args),
      onAcceptChangeById: (...args) => commentCallbacksRef.current.onAcceptChangeById?.(...args),
      onRejectChangeById: (...args) => commentCallbacksRef.current.onRejectChangeById?.(...args),
      onTrackedChangeReply: (...args) => {
        if (!openingRef.current) commentCallbacksRef.current.onTrackedChangeReply?.(...args);
      },
    }),
    []
  );

  const sidebarTrackedChanges = useMemo(
    () =>
      allowHostProposals && !showHostProposalsInSidebar
        ? trackedChanges.filter(
            (change) =>
              !hostProposalRevisions.has(`revision-${change.revisionId}`) &&
              (change.insertionRevisionId == null ||
                !hostProposalRevisions.has(`revision-${change.insertionRevisionId}`))
          )
        : trackedChanges,
    [trackedChanges, allowHostProposals, showHostProposalsInSidebar, hostProposalRevisions]
  );

  const commentSidebarItems = useCommentSidebarItems({
    comments,
    trackedChanges: sidebarTrackedChanges,
    callbacks: stableCallbacks,
    showResolved: showCommentsSidebar,
    isAddingComment: showCommentsSidebar ? isAddingComment : false,
    addCommentYPosition,
  });

  const allSidebarItems = useMemo(
    () =>
      mergeSidebarItems(
        showCommentsSidebar ? commentSidebarItems : [],
        pluginSidebarItems ?? [],
        pluginHost.sidebarItems
      ),
    [showCommentsSidebar, commentSidebarItems, pluginSidebarItems, pluginHost.sidebarItems]
  );

  useEffect(() => {
    if (
      expandedSidebarItem?.startsWith('plugin:') &&
      !allSidebarItems.some((item) => item.id === expandedSidebarItem)
    ) {
      setExpandedSidebarItem(null);
    }
  }, [allSidebarItems, expandedSidebarItem]);

  // Map insertion revisionIds to the sidebar card id prefix of replacement tracked changes,
  // so clicking the insertion part of a replacement activates the same card.
  const revisionIdAliases = useMemo(() => {
    const map = new Map<string, string>();
    trackedChanges.forEach((change) => {
      if (change.type === 'replacement' && change.insertionRevisionId != null) {
        map.set(String(change.insertionRevisionId), `tc-${change.revisionId}-`);
      }
    });
    return map;
  }, [trackedChanges]);

  // An opening document's comment cards arrive with the full document: keep their space meanwhile.
  const sidebarOpen =
    allSidebarItems.some((item) => !item.hidden) || (opening && showCommentsSidebar);

  const requestReplica = yrsCore.requestReplica;
  const replicaPending = Boolean(
    experimentalWorkerOpen && yrsCore.hydrateOnDemand && yrsCore.session && !yrsCore.replicaReady
  );
  const replicaWanted =
    (!(experimentalWorkerOpen && workerProposals) &&
      ((plugins?.length ?? 0) > 0 || Boolean(onRenderedDomContextReady))) ||
    (!viewerReads && (showCommentsSidebar || sidebarOpen || showOutline));
  useEffect(() => {
    if (replicaPending && replicaWanted) requestReplica();
  }, [replicaPending, replicaWanted, requestReplica, yrsCore.session]);
  // An outline opened before the replica loaded reads its headings once it has.
  const replicaReady = yrsCore.replicaReady;
  useEffect(() => {
    if (!viewerReads && experimentalWorkerOpen && replicaReady && showOutlineRef.current) refreshHeadings();
  }, [experimentalWorkerOpen, replicaReady, refreshHeadings, showOutlineRef, viewerReads]);
  // A tap asks through its gesture, the input for itself.
  useEffect(() => {
    const content = editorContentRef.current;
    if (!replicaPending || !content || viewerSession) return;
    const onPointer = (event: PointerEvent) => {
      if (pagePressNeedsReplica(event)) requestReplica();
    };
    content.addEventListener('pointerdown', onPointer, true);
    return () => content.removeEventListener('pointerdown', onPointer, true);
  }, [replicaPending, requestReplica, viewerSession]);

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

  // pageWidthPx — the final section's width — positions the sidebar / comment
  // margin markers against the page most content lives under.
  const sectionPropsPageWidth = docBody?.finalSectionProperties?.pageWidth;
  const pageWidthPx = sectionPropsPageWidth
    ? Math.round(sectionPropsPageWidth / 15)
    : DEFAULT_PAGE_WIDTH;

  const handlePagedViewerSelectionChange = useCallback((selection: ViewerSelectionChange) => {
    pluginHost.publishViewerSelection(selection);
    handleViewerSelectionChange(selection);
  }, [handleViewerSelectionChange, pluginHost.publishViewerSelection]);

  // PagedEditor selection callback: resolve sticky comment/revision coverage
  // from Yrs so the matching sidebar card opens as the caret moves.
  const handlePagedSelectionChange = useCallback(() => {
    // Body selection transitions arrive here even when the derived toolbar
    // context is unchanged. Notify the canvas live region from the authoritative
    // selection event so range/caret announcements are never lost to toolbar
    // state deduplication.
    canvasA11yNotifyRef.current?.();
    pluginHost.publishSelection();
    if (viewerReads) return;
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
        const revId = String(yrsIdToNumericId(revision.revisionId));
        if (
          allowHostProposalsRef.current &&
          hostProposalRevisionsRef.current.has(`revision-${revId}`)
        ) {
          continue;
        }
        const start = session.locateParagraph(head.story, revision.range.start.paraId).start + revision.range.start.offset;
        const end = session.locateParagraph(head.story, revision.range.end.paraId).start + revision.range.end.offset;
        if (start <= offset && offset <= end) {
          const prefix = `tc-${revId}-`;
          let match = commentSidebarItems.find((item) => item.id.startsWith(prefix));
          if (!match) {
            const aliasedPrefix = revisionIdAliases.get(revId);
            if (aliasedPrefix) {
              match = commentSidebarItems.find((item) => item.id.startsWith(aliasedPrefix));
            }
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
  }, [
    viewerReads,
    comments,
    resolvedCommentIds,
    commentSidebarItems,
    revisionIdAliases,
    setShowCommentsSidebar,
    pluginHost.publishSelection,
  ]);

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
    if (sidebarAutoOpenedRef.current || commentSidebarItems.length === 0) return;
    const proposed = getProposalAnchorKeys();
    if (commentSidebarItems.every((item) => proposed.has(item.anchorKey ?? ''))) return;
    sidebarAutoOpenedRef.current = true;
    setShowCommentsSidebar(true);
  }, [commentSidebarItems, getProposalAnchorKeys, setShowCommentsSidebar]);
  // A document whose worker found tracked changes asks for the sidebar, as a loaded one does; a
  // host that keeps the sidebar closed keeps the replica unloaded.
  workerRevisionsRef.current = () => setShowCommentsSidebar(true);

  const editorContainerStyle: CSSProperties = {
    flex: 1,
    minHeight: 0,
    minWidth: 0, // Allow flex item to shrink below content width on narrow viewports
    overflow: 'auto', // Sole scroll container — PagedEditor sizes to content
    position: 'relative',
    overflowAnchor: 'none',
  };

  const chromeContext = useMemo<EditorChrome>(
    () => ({ showZoomControl, toolbarExtra }),
    [showZoomControl, toolbarExtra]
  );
  const chrome = !showToolbar ? null : toolbar !== undefined ? (
    toolbar !== null && (
      <div className="z-50 flex flex-col gap-0 flex-shrink-0">
        {toolbar}
      </div>
    )
  ) : readOnlyProp ? null : (
    <DocxEditorToolbar
      renderLogo={renderLogo}
      documentName={documentName}
      onDocumentNameChange={onDocumentNameChange}
      documentNameEditable={documentNameEditable}
      renderTitleBarRight={renderTitleBarRight}
    />
  );

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
    <DocxCommandProvider commands={commandController.store}>
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
          editable: !readOnly && yrsCore.replicaReady,
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
          editable: !readOnly && yrsCore.replicaReady,
          onTopMarginChange: handleTopMarginChange,
          onBottomMarginChange: handleBottomMarginChange,
        }}
        outlineProps={{
          headings: outlineHeadings,
          onHeadingClick: handleHeadingInfoClick,
          onClose: () => setShowOutline(false),
          scrollLeft: editorScrollLeft,
        }}
        onToggleOutline={handleToggleOutline}
        scrollPageInfo={scrollPageInfo}
        renderDock={(placement, available) => (
          <PluginDock
            host={pluginHost.host}
            placement={placement}
            activations={pluginHost.activations.filter(
              (activation) => activation.plugin.panel?.placement === placement
            )}
            available={available}
          />
        )}
        toolbar={
          chrome && (
            <EditorChromeContext.Provider value={chromeContext}>{chrome}</EditorChromeContext.Provider>
          )
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
            interactive={!readOnly && yrsCore.replicaReady}
            fontFamilies={fontAliases}
          >
            <DocxEditorPagedArea
              commandBridgeRef={commandBridgeRef}
              yrsCore={yrsCore}
              onError={reportPagedError}
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
              partEditTarget={partEditTarget}
              setPartEditTarget={setPartEditTarget}
              onHeaderFooterDoubleClick={handleHeaderFooterDoubleClick}
              onNoteClick={setPartEditTarget}
              onRemoveHeaderFooter={handleRemoveHeaderFooter}
              onBodyClick={handleBodyClick}
              zoom={state.zoom}
              readOnly={readOnly}
              viewerDocumentRead={viewerSession ? canvasRenderer.readWorkerDocument : undefined}
              showHiddenText={showHiddenText}
              isSuggesting={editingMode === 'suggesting'}
              author={author}
              measurementFontProvider={measurementFontProvider}
              rustFontChainsProviderRef={rustFontChainsProviderRef}
              onYrsContentChange={handleYrsContentChange}
              onPagedSelectionChange={handlePagedSelectionChange}
              onYrsSelectionChange={handleYrsToolbarSelectionChange}
              onViewerSelectionChange={handlePagedViewerSelectionChange}
              onRenderedDomContextReady={
                pluginHost.managed || onRenderedDomContextReady
                  ? pluginHost.onRenderedDomContext
                  : undefined
              }
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
              onViewerCommentRangesChange={setViewerCommentRanges}
              viewerSidebarActive={showCommentsSidebar || sidebarOpen}
              pluginRenderedDomContext={sidebarDomContext}
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
              layoutInWorker={canvasRenderer.layoutInWorker}
              fontRequirementsInWorker={
                experimentalWorkerOpen ? canvasRenderer.fontRequirementsInWorker : undefined
              }
              applyResidentInput={canvasRenderer.applyInput}
              applyResidentDelete={canvasRenderer.applyDelete}
              displayListQueries={canvasRenderer.queries}
              resolveDisplayListQueries={canvasRenderer.resolveQueries}
              pageNavigation={canvasRenderer.pageNavigation}
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
            {!readOnly && yrsCore.replicaReady && (
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
            <PluginOverlays
              host={pluginHost.host}
              activations={pluginHost.activations}
              target={pluginOverlayTarget}
              layerRef={pluginHost.overlayLayerRef}
              heldGeometry={pluginHost.heldGeometry}
            />
          </CanvasPagedArea>
        }
        loadingIndicator={
          canvasRenderer.status === 'loading' && (loadingIndicator || <DefaultLoadingIndicator />)
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
            onOpenImageProperties={() => void commandController.store.execute('imageProperties', null)}
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
            pmImageContext={imageTarget}
            showPageSetup={showPageSetup}
            onPageSetupClose={() => setShowPageSetup(false)}
            onPageSetupApply={(properties) => {
              void commands.complete('pageSetup', () => {
                dialogApply.current.pageSetup(properties);
                return commandOutcome(true);
              });
            }}
            showWatermark={showWatermark}
            onWatermarkClose={() => setShowWatermark(false)}
            onWatermarkApply={(watermark) => {
              void commands.complete('watermark', () => {
                dialogApply.current.watermark(watermark);
                return commandOutcome(true);
              });
            }}
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
        and the current match shows as an ordinary selection as before. */}
      {canvasFindOverlayTarget && canvasRenderer.queries ? (
        <CanvasFindHighlightOverlay
          matches={hostSearch.highlight?.matches ?? canvasFindMatches}
          currentIndex={hostSearch.highlight?.current ?? findReplace.state.currentIndex}
          overlayTarget={canvasFindOverlayTarget}
          canvasHostRef={canvasRenderer.canvasHostRef}
          displayListQueries={canvasRenderer.queries}
          sidebarOpen={sidebarOpen}
          zoom={state.zoom}
        />
      ) : null}
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
    </DocxCommandProvider>
  );
});

// ============================================================================
// EXPORTS
// ============================================================================

export default DocxEditor;
