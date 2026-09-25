/**
 * `<XlsxEditor />` — the editor shell. A dpr-aware canvas paints the grid; DOM
 * overlays (selection marquee, active-cell outline, in-cell editor) sit above it,
 * positioned from the same display-list geometry the painter uses. A top bar
 * holds the name box, formula bar, and save/undo/redo; an offscreen `role=grid`
 * mirror serves screen readers. All compute lives in `@betteroffice/xlsx`; this
 * layer is framework glue — keyboard/mouse wiring, focus flow, and DOM chrome.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { SetStateAction } from 'react';
import {
  buildA11yGrid,
  cellAtPoint,
  cellRect,
  chartRegionAtPoint,
  extendTo,
  fromTsv,
  hyperlinkAtCell,
  initWasm,
  isPngExportAvailable,
  isProposalsAvailable,
  moveFocus,
  normalizeRange,
  openWorkbook,
  paintDisplayList,
  parseHyperlinkLocation,
  rangeRect,
  selectionAt,
  selectionKeyReducer,
  safeExternalHyperlink,
  toTsv,
} from '@betteroffice/xlsx';
import type {
  CellAddr,
  CellEdit,
  CellInputEdit,
  CellRange,
  CapturedFormat,
  ChartRegion,
  DisplayList,
  DrawCmd,
  Direction,
  EditResult,
  MergedRange,
  Proposal,
  Selection,
  SelectionLimits,
  SheetInfo,
  WorkbookHandle,
  XlsxEditRefusal,
  XlsxEditRequest,
  XlsxEditResult,
  XlsxFindRequest,
  XlsxFindResult,
  XlsxReadRequest,
  XlsxReadResult,
  XlsxValidationResult,
} from '@betteroffice/xlsx';
import type {
  AwarenessPeer,
  CollaborationAwareness,
  CollaborationReplica,
} from '@betteroffice/xlsx/collaboration';
import type { Translations } from '@betteroffice/xlsx-i18n';
import { LocaleProvider, useTranslation } from './i18n';
import {
  createXlsxCommandController,
  XlsxCommandAdmissionError,
} from './commands/createXlsxCommandStore';
import { commandForEvent } from './commands/descriptors';
import { useXlsxCommand } from './commands/hooks';
import {
  createInputCoordinator,
  type InputCoordinatorHooks,
  type InputDraft,
} from './commands/inputCoordinator';
import type { XlsxCommandStore } from './commands/types';
import { useCommandShortcuts } from './commands/useCommandShortcuts';
import { useXlsxCommandBinding, type XlsxEditorBridge } from './commands/useXlsxCommands';
import { XlsxCommandContext } from './commands/XlsxCommandProvider';
import { EditorToolbar } from './components/EditorToolbar';
import { EditorChromeContext } from './components/EditorToolbarContext';
import type { BorderStyle } from './components/Toolbar';
import { FormulaBarContext, type FormulaBarBinding } from './components/toolbar/FormulaBar';
import { ToolbarCommandButton } from './components/toolbar/ToolbarCommand';
import { ToolbarIcon } from './components/ui/ToolbarIcon';
import { ToolbarButtonBase, ToolbarGroup } from './components/ui/ToolbarPrimitives';
import {
  expandRangeToMergedCells,
  PresenceStrip,
  RemoteSelections,
} from './presence/Presence';
import { ProposalsPanel } from './proposals/ProposalsPanel';

/**
 * The imperative surface handed to {@link XlsxEditorProps.onReady}: the open
 * workbook handle plus a `refreshProposals` to re-read the pending list after an
 * external caller (e.g. a demo agent) stages proposals on the same handle.
 *
 * The version, read and edit-batch methods run in order with the input accepted
 * before them, as commands do: cell and formula entries, chart moves, pastes and
 * cuts land first, and an IME composition ends with its text written. They reject
 * with the command failure `code`: `input-failed` while a refused entry waits for
 * correction or earlier input failed, `gesture-active` during a chart drag, and
 * `document-replaced` when the workbook is replaced meanwhile.
 */
export interface XlsxEditorApi {
  /**
   * Closes the open cell entry and clears the selection. The entry is written
   * at once, or queued behind input still waiting to be written.
   */
  clearSelection: () => void;
  /**
   * The editor's command store, shared by its toolbar and host chrome. It
   * gates the editor's own UI; `handle` stays unrestricted host authority.
   */
  readonly commands: XlsxCommandStore;
  focus: () => void;
  handle: WorkbookHandle;
  refreshProposals: () => void;
  /**
   * Writes the open cell entry and returns the workbook bytes at once. Throws
   * {@link XlsxSaveRefusedError} rather than return bytes without accepted
   * input: `input-pending` while earlier entries or a paste still wait to be
   * written (await `commands.execute('save', null)` instead), `input-failed`
   * while an entry the workbook refused waits for correction.
   */
  save: () => Uint8Array;
  /**
   * Scrolls the focus cell into view. Commands run after it, even in the same
   * handler, act on this selection and sheet. The open cell entry is written
   * first, or queued behind input still waiting to be written, so it may not
   * have landed when this returns.
   */
  selectCells: (sheet: number, selection: Selection) => boolean;
  version: () => Promise<string>;
  readCells: (request: XlsxReadRequest) => Promise<XlsxReadResult>;
  findText: (request: XlsxFindRequest) => Promise<XlsxFindResult>;
  /** Refuses with `read-only` while the editor is read-only. */
  validateEdits: (request: XlsxEditRequest) => Promise<XlsxValidationResult>;
  /**
   * Applies the batch against the caller's `expectVersion` after committing pending input,
   * so input that lands first refuses it with `stale-version`. Refuses with `read-only`
   * while the editor is read-only; an applied batch calls `onChange` once.
   */
  applyEdits: (request: XlsxEditRequest) => Promise<XlsxEditResult>;
}

function readOnlyRefusal(handle: WorkbookHandle): XlsxEditRefusal {
  return {
    ok: false,
    version: handle.version(),
    failure: { code: 'read-only', message: 'The editor is read-only' },
  };
}

/** Why a synchronous {@link XlsxEditorApi.save} did not return bytes. */
export class XlsxSaveRefusedError extends Error {
  constructor(readonly code: 'input-pending' | 'input-failed') {
    super(
      code === 'input-pending'
        ? 'Accepted input is still being written; await commands.execute("save", null)'
        : 'A cell entry could not be written; correct or discard it first'
    );
    this.name = 'XlsxSaveRefusedError';
  }
}

export interface XlsxEditorCollaborationOptions {
  /** Peer-unique Yrs client ID. Generated securely when omitted. */
  clientId?: number;
  /** Shared Yrs state applied before the replica is exposed. */
  initialUpdate?: Uint8Array;
  /** Receive the editor-owned collaboration replica. */
  onReplica?: (replica: CollaborationReplica | null) => void;
  /** Presence-capable provider connected to the editor-owned replica. */
  provider?: CollaborationAwareness | null;
}

/**
 * Props for {@link XlsxEditor}.
 */
export interface XlsxEditorProps {
  /** Raw .xlsx bytes to open. When omitted the shell paints a demo frame. */
  file?: Uint8Array;
  /** Download name for the save button; falls back to `workbook.xlsx`. */
  fileName?: string;
  /** Receive saved bytes instead of triggering a browser download. */
  onSave?: (bytes: Uint8Array) => void;
  /** Called after a user edit changes the workbook. */
  onChange?: () => void;
  /** Open a network-ready Yrs replica and repaint when peer updates arrive. */
  collaboration?: XlsxEditorCollaborationOptions;
  i18n?: Translations;
  /**
   * Called when a workbook opens, with a handle to stage agent proposals and a
   * way to refresh the panel afterward. Enables demo/host agents without
   * exposing the wasm object through the render tree. A returned cleanup runs
   * before the workbook is replaced or disposed.
   */
  onReady?: (api: XlsxEditorApi) => void | (() => void);
  className?: string;
  /** Blocks user edits; navigation and selection remain available. */
  readOnly?: boolean;
  /**
   * Replaces the default toolbar: omitted keeps it (hidden when read-only),
   * `null` removes it, and supplied chrome renders instead, also when
   * read-only. Compose it from `EditorToolbar mode="commands"` parts.
   */
  toolbar?: React.ReactNode;
  /** `false` hides the toolbar region, whatever `toolbar` is. */
  showToolbar?: boolean;
}

/** the open in-cell editor: which cell it targets and its current draft text. */
interface EditState {
  row: number;
  col: number;
  value: string;
}

const COL_W = 96;
const ROW_H = 24;
const BRAND = '#217346';
const DEFAULT_XLSX_TOOLBAR_HEIGHT = 87;
const MAX_OVERLAY_MERGED_RANGES = 1024;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const CHART_NUDGE_PX = 1;
const CHART_NUDGE_MULTIPLIER = 10;
// how long a run of arrow presses may stay local before it lands as one edit.
const CHART_NUDGE_SETTLE_MS = 250;
// how long print waits for the canvas to show the latest change before failing.
const PAINT_SETTLE_MS = 1000;
const CHART_NUDGE_KEYS: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0],
  ArrowRight: [1, 0],
  ArrowUp: [0, -1],
  ArrowDown: [0, 1],
};

// a placeholder grid frame so the shell paints something real when no file is
// open. real files render through the wasm display list instead.
function buildDemoDisplayList(width: number, height: number, cellText: string): DisplayList {
  const commands: DrawCmd[] = [
    { op: 'fillRect', x: 0, y: 0, w: width, h: height, color: '#ffffff' },
  ];
  const cols = Math.ceil(width / COL_W);
  const rows = Math.ceil(height / ROW_H);
  for (let c = 0; c <= cols; c++) {
    commands.push({
      op: 'line',
      x1: c * COL_W,
      y1: 0,
      x2: c * COL_W,
      y2: height,
      width: 1,
      color: '#e0e0e0',
    });
  }
  for (let r = 0; r <= rows; r++) {
    commands.push({
      op: 'line',
      x1: 0,
      y1: r * ROW_H,
      x2: width,
      y2: r * ROW_H,
      width: 1,
      color: '#e0e0e0',
    });
  }
  commands.push({ op: 'fillRect', x: 0, y: 0, w: width, h: ROW_H, color: '#f3f3f3' });
  commands.push({ op: 'fillRect', x: 0, y: 0, w: COL_W, h: height, color: '#f3f3f3' });
  commands.push({
    op: 'text',
    x: COL_W + 8,
    y: ROW_H + 18,
    text: cellText,
    fontSize: 14,
    color: '#202020',
    clip: { x: COL_W, y: ROW_H, w: COL_W * 3, h: ROW_H },
    align: 'left',
  });
  return { width, height, commands };
}

// median of the gaps between consecutive offsets, or a fallback when the window
// has no tracks. the median ignores outliers like a single very wide column, so
// the extent-to-count estimate below is not skewed by one atypical track.
function medianTrack(offsets: number[] | undefined, fallback: number): number {
  if (!offsets || offsets.length < 2) return fallback;
  const gaps: number[] = [];
  for (let i = 1; i < offsets.length; i++) gaps.push(offsets[i] - offsets[i - 1]);
  gaps.sort((a, b) => a - b);
  const mid = gaps[gaps.length >> 1];
  return mid > 0 ? mid : fallback;
}

// derive nav bounds from the scrollable extent: rows/cols estimated from the
// content size over a representative (median) track size, rowsPerPage from the
// viewport. a slack of one keeps the row/col just past the used edge reachable.
function deriveLimits(
  dl: DisplayList | null,
  info: SheetInfo,
  viewportHeight: number
): SelectionLimits {
  const rowH = medianTrack(dl?.grid?.rowOffsets, ROW_H);
  const colW = medianTrack(dl?.grid?.colOffsets, COL_W);
  const rows = Math.max(1, Math.round(info.contentHeight / rowH)) + 1;
  const cols = Math.max(1, Math.round(info.contentWidth / colW)) + 1;
  const rowsPerPage = Math.max(1, Math.floor(viewportHeight / rowH));
  return { rows, cols, rowsPerPage };
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

// the png download name derived from the workbook name: swap .xlsx for .png.
function pngName(fileName: string | undefined): string {
  return `${(fileName ?? 'workbook.xlsx').replace(/\.xlsx$/i, '')}.png`;
}

/** The workbook, sheet and range a clipboard command read when it was invoked. */
interface CellTarget {
  handle: WorkbookHandle;
  sheet: number;
  range: CellRange;
}

interface PaintMark {
  generation: number;
  mutation: number;
  /** Sheet and zoom. */
  view: string;
}

/** Whether a paint shows everything `wanted` asks for. */
function covers(painted: PaintMark, wanted: PaintMark): boolean {
  return (
    painted.generation === wanted.generation &&
    painted.view === wanted.view &&
    painted.mutation >= wanted.mutation
  );
}

function scaledRect(rect: { x: number; y: number; w: number; h: number }, zoom: number) {
  return {
    x: rect.x * zoom,
    y: rect.y * zoom,
    w: rect.w * zoom,
    h: rect.h * zoom,
  };
}

const visuallyHidden: React.CSSProperties = {
  position: 'absolute',
  width: 1,
  height: 1,
  margin: -1,
  padding: 0,
  border: 0,
  overflow: 'hidden',
  clip: 'rect(0 0 0 0)',
  whiteSpace: 'nowrap',
};

const xlsxToolbarStyles: Record<string, React.CSSProperties> = {
  shell: {
    flex: '0 0 auto',
    minHeight: DEFAULT_XLSX_TOOLBAR_HEIGHT,
    padding: '4px 0 5px',
    borderBottom: '1px solid #e2e8f0',
    background: '#ffffff',
    color: '#0f172a',
    boxSizing: 'border-box',
  },
  rail: {
    display: 'flex',
    alignItems: 'center',
    minHeight: 32,
    margin: '0 8px',
    padding: '2px 8px',
    borderRadius: 4,
    background: '#ffffff',
    boxSizing: 'border-box',
    overflowX: 'auto',
    overflowY: 'hidden',
  },
  group: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 1,
    padding: '0 6px',
    borderRight: '1px solid rgba(226, 232, 240, 0.9)',
    flex: '0 0 auto',
  },
  host: {
    flex: '0 0 auto',
  },
  proposals: {
    marginLeft: 'auto',
    paddingLeft: 6,
    flex: '0 0 auto',
  },
  count: {
    display: 'inline-grid',
    placeItems: 'center',
    minWidth: 15,
    height: 15,
    marginLeft: -3,
    padding: '0 4px',
    borderRadius: 8,
    background: '#0f172a',
    color: '#ffffff',
    fontSize: 10,
    fontWeight: 700,
    lineHeight: 1,
    boxSizing: 'border-box',
  },
};

/**
 * React state whose latest value the command authority reads at once: the
 * setter updates the ref synchronously and schedules the render.
 */
function useSyncedState<T>(initial: T) {
  const [value, setValue] = useState(initial);
  const ref = useRef(value);
  const set = useCallback((next: SetStateAction<T>) => {
    ref.current = typeof next === 'function' ? (next as (previous: T) => T)(ref.current) : next;
    setValue(ref.current);
  }, []);
  return [value, set, ref] as const;
}

/**
 * The xlsx editor React component.
 */
export function XlsxEditor(props: XlsxEditorProps) {
  return (
    <LocaleProvider i18n={props.i18n}>
      <XlsxEditorContent {...props} />
    </LocaleProvider>
  );
}

function XlsxEditorContent({
  file,
  fileName,
  onSave,
  onChange,
  collaboration,
  onReady,
  className,
  readOnly = false,
  toolbar,
  showToolbar = true,
  i18n,
}: XlsxEditorProps) {
  const { t } = useTranslation();
  const collaborationEnabled = collaboration !== undefined;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const collaborationOnReplica = collaboration?.onReplica;
  const collaborationProvider = collaboration?.provider;
  const [toolbarElement, setToolbarElement] = useState<HTMLDivElement | null>(null);
  const [commandController] = useState(createXlsxCommandController);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const handleRef = useRef<WorkbookHandle | null>(null);
  const frameRef = useRef<DisplayList | null>(null);
  // the exact frame on screen, stored with the zoom it was painted at. scroll
  // repaints are rAF-coalesced and a mutation republishes the frame, so hit
  // testing reads this snapshot rather than the live scroll offset or the
  // current model — either would answer for pixels that are not on screen.
  const paintedRef = useRef<{ frame: DisplayList; zoom: number } | null>(null);
  const rafRef = useRef<number | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const editorInputRef = useRef<HTMLInputElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const draggingRef = useRef(false);
  const clickStartRef = useRef<CellAddr | null>(null);
  // an in-flight chart drag: which chart, and where the pointer went down.
  const chartDragRef = useRef<{ id: string; clientX: number; clientY: number } | null>(null);
  // an arrow-key burst not yet landed: the chart and the logical px accumulated.
  const nudgeRef = useRef<{ id: string; dx: number; dy: number } | null>(null);
  const nudgeTimerRef = useRef<number | null>(null);
  const suppressBlurRef = useRef(false);
  const pendingSheetViewRef = useRef(false);
  const flushNudgeRef = useRef<() => void>(() => {});
  const settlePendingEditsRef = useRef<() => boolean>(() => true);
  const hasRejectedRef = useRef<() => boolean>(() => false);
  // bumped when a workbook opens or closes; drafts and commands from an older
  // document never write into the next one.
  const generationRef = useRef(0);
  const mutationRef = useRef(0);
  const inputHooksRef = useRef<Omit<InputCoordinatorHooks, 'generation'> | null>(null);
  const [coordinator] = useState(() =>
    createInputCoordinator({
      generation: () => generationRef.current,
      seal: () => inputHooksRef.current?.seal() ?? {},
      sync: () => inputHooksRef.current?.sync(),
      write: (draft) => inputHooksRef.current?.write(draft) ?? false,
      close: (draft) => inputHooksRef.current?.close(draft),
      restore: (draft) => inputHooksRef.current?.restore(draft) ?? false,
    })
  );
  // an IME composition in the cell editor or formula bar, settled when it ends.
  const compositionRef = useRef<{
    source: InputDraft['source'];
    done: Promise<boolean>;
    settle: (ended: boolean) => void;
  } | null>(null);
  const suppressFormulaBlurRef = useRef(false);
  // what the canvas last painted: the document, its changes and the view.
  const paintMarkRef = useRef<PaintMark | null>(null);
  const paintWaitersRef = useRef<{ mark: PaintMark; resolve: (painted: boolean) => void }[]>([]);
  // latest onReady, read (not depended on) by the open effect so a changing
  // callback identity never reopens the workbook.
  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  const readOnlyRef = useRef(readOnly);
  readOnlyRef.current = readOnly;

  const [sheetInfo, setSheetInfo, sheetInfoRef] = useSyncedState<SheetInfo | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [renderError, setRenderError] = useState<string | null>(null);
  const [frame, setFrame] = useState<DisplayList | null>(null);
  const [selection, setSelection, selectionRef] = useSyncedState<Selection | null>(null);
  const [editing, setEditing] = useState<EditState | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [toolbarHeight, setToolbarHeight] = useState(DEFAULT_XLSX_TOOLBAR_HEIGHT);
  const [zoom, setZoom, zoomRef] = useSyncedState(1);
  const [revision, setRevision] = useState(0);
  const [dragging, setDragging] = useState(false);
  // the selected chart, and the live pointer offset while it is dragged.
  // `movable` rides along so the arrow keys never depend on a frame lookup.
  const [selectedChart, setSelectedChart, selectedChartRef] = useSyncedState<{
    id: string;
    movable: boolean;
  } | null>(
    null
  );
  const [chartDragOffset, setChartDragOffset] = useState<{ x: number; y: number } | null>(null);
  // logical-px preview of an arrow burst that has not landed yet.
  const [nudgeOffset, setNudgeOffset] = useState<{ x: number; y: number } | null>(null);
  const [visibleMergedRanges, setVisibleMergedRanges] = useState<readonly MergedRange[]>([]);
  // the border style and color the next border preset applies.
  const borderStyleChoiceRef = useRef<BorderStyle | undefined>(undefined);
  const borderColorChoiceRef = useRef<string | undefined>(undefined);
  const [capturedFormat, setCapturedFormat, capturedFormatRef] =
    useSyncedState<CapturedFormat | null>(null);
  const paintSourceRef = useRef<string | null>(null);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen, proposalsPanelOpenRef] =
    useSyncedState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);
  const [awarenessPeers, setAwarenessPeers] = useState<readonly AwarenessPeer[]>([]);
  // a1 lists keyed by proposal id: cells that drifted since a proposal was
  // staged, surfaced when accepting it throws a StaleProposalError.
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});

  const activeSheet = sheetInfo?.activeSheet ?? 0;
  const activeSheetRef = useMemo(
    () => ({
      get current() {
        return sheetInfoRef.current?.activeSheet ?? 0;
      },
    }),
    [sheetInfoRef]
  );

  const draftFor = useCallback(
    (source: InputDraft['source'], row: number, col: number, value: string): InputDraft => ({
      generation: generationRef.current,
      sheet: activeSheetRef.current,
      row,
      col,
      value,
      source,
    }),
    []
  );

  // puts a draft back into its own input at its own cell.
  const showDraft = useCallback((draft: InputDraft) => {
    setSelection(selectionAt({ row: draft.row, col: draft.col }));
    if (draft.source === 'cell') setEditing({ row: draft.row, col: draft.col, value: draft.value });
    else setFormulaDraft(draft.value);
  }, []);

  // a rejected draft waits for correction on its own sheet, once no other
  // draft is open there.
  const showRejected = useCallback(
    (sheet: number) => {
      if (coordinator.draft) return;
      const entry = coordinator.rejected.find(
        (candidate) => candidate.sheet === sheet && candidate.generation === generationRef.current
      );
      if (!entry) return;
      coordinator.setDraft(entry);
      showDraft(entry);
    },
    [coordinator, showDraft]
  );

  const dropDrafts = useCallback(() => {
    coordinator.setDraft(null);
    setEditing(null);
    setFormulaDraft(null);
  }, [coordinator]);

  const clearSelection = useCallback(() => {
    if (!settlePendingEditsRef.current()) return;
    setSelection(null);
    setSelectedChart(null);
    dropDrafts();
    setCapturedFormat(null);
    paintSourceRef.current = null;
    commandController.refresh();
  }, [dropDrafts, commandController, setSelection, setSelectedChart, setCapturedFormat]);

  const selectCells = useCallback(
    (sheet: number, nextSelection: Selection): boolean => {
      const handle = handleRef.current;
      if (!handle || !Number.isInteger(sheet)) return false;
      try {
        const info = handle.sheetInfo();
        if (sheet < 0 || sheet >= info.sheetNames.length) return false;
        handle.cellPosition(sheet, nextSelection.anchor.row, nextSelection.anchor.col);
        const position = handle.cellPosition(
          sheet,
          nextSelection.focus.row,
          nextSelection.focus.col
        );
        if (!settlePendingEditsRef.current()) return false;
        handle.setActiveSheet(sheet);
        setSheetInfo(handle.sheetInfo());
        setSelection({
          anchor: { ...nextSelection.anchor },
          focus: { ...nextSelection.focus },
        });
        setSelectedChart(null);
        dropDrafts();
        setCapturedFormat(null);
        paintSourceRef.current = null;
        requestAnimationFrame(() => {
          const scroll = scrollRef.current;
          if (!scroll) return;
          scroll.scrollLeft = position.x * zoomRef.current;
          scroll.scrollTop = position.y * zoomRef.current;
        });
        setError(null);
        commandController.refresh();
        return true;
      } catch {
        return false;
      }
    },
    []
  );

  useEffect(() => {
    if (!readOnly) return;
    dropDrafts();
    setCapturedFormat(null);
    paintSourceRef.current = null;
    chartDragRef.current = null;
    setChartDragOffset(null);
    nudgeRef.current = null;
    setNudgeOffset(null);
    if (nudgeTimerRef.current != null) clearTimeout(nudgeTimerRef.current);
    nudgeTimerRef.current = null;
  }, [readOnly]);

  // whether the embedded core was built with png export (raster cargo feature).
  // stable for the module's lifetime, so the export control can pre-disable.
  const pngExportAvailable = useMemo(() => isPngExportAvailable(), []);

  // whether the embedded core exposes the proposals api; gates all proposal
  // chrome so the editor degrades cleanly against an older module.
  const proposalsAvailable = useMemo(() => isProposalsAvailable(), []);

  useEffect(() => {
    if (!toolbarElement) {
      setToolbarHeight(0);
      return;
    }
    const updateHeight = () => setToolbarHeight(toolbarElement.offsetHeight);
    updateHeight();
    const observer = new ResizeObserver(updateHeight);
    observer.observe(toolbarElement);
    return () => observer.disconnect();
  }, [toolbarElement]);

  // runs a host read or batch in the input queue, after the input accepted before it.
  const afterInput = useCallback(
    <T,>(opened: WorkbookHandle, operation: () => T): Promise<T> => {
      if (handleRef.current !== opened) {
        return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
      }
      return coordinator.runAfterPendingInput(() => {
        if (handleRef.current !== opened) throw new XlsxCommandAdmissionError('document-replaced');
        return operation();
      });
    },
    [coordinator]
  );

  // re-read the pending proposal list and queue a repaint — ghosts paint into
  // the engine frame, so every lifecycle change (propose/accept/reject) must
  // republish it. safe to call against an old core (the loader returns an
  // empty list).
  const refreshProposals = useCallback(() => {
    const handle = handleRef.current;
    if (!handle) {
      setProposals([]);
      return;
    }
    try {
      setProposals(handle.listProposals());
    } catch {
      setProposals([]);
    }
    setRevision((r) => r + 1);
  }, []);

  // open the workbook when the file changes; dispose it on change/unmount and
  // reset all editing state so a dropped file starts clean.
  useEffect(() => {
    generationRef.current += 1;
    dropDrafts();
    setSelectedChart(null);
    // a burst belongs to the document it was typed on: its timer would fire
    // against whatever workbook `handleRef` holds by then.
    nudgeRef.current = null;
    setNudgeOffset(null);
    if (nudgeTimerRef.current != null) {
      clearTimeout(nudgeTimerRef.current);
      nudgeTimerRef.current = null;
    }
    setProposals([]);
    setStaleFor({});
    setProposalsPanelOpen(false);
    setCollaborationReplica(null);
    setVisibleMergedRanges([]);
    borderStyleChoiceRef.current = undefined;
    borderColorChoiceRef.current = undefined;
    setCapturedFormat(null);
    setRenderError(null);
    paintSourceRef.current = null;
    pendingSheetViewRef.current = false;
    if (!file) {
      handleRef.current = null;
      setSheetInfo(null);
      setSelection(null);
      setError(null);
      return;
    }
    handleRef.current = null;
    setSheetInfo(null);
    setSelection(null);
    let handle: WorkbookHandle | null = null;
    let unsubscribeUpdates = () => {};
    let cleanupReady = () => {};
    let disposed = false;
    const runReadyCleanup = () => {
      const cleanup = cleanupReady;
      cleanupReady = () => {};
      try {
        cleanup();
      } catch {}
    };
    void initWasm().then(
      () => {
        if (disposed) return;
        try {
          handle = openWorkbook(file, {
            collaborative: collaborationEnabled,
            clientId: collaborationClientId,
          });
          if (collaborationInitialUpdate) {
            handle.applyUpdate(collaborationInitialUpdate.slice());
          }
          handleRef.current = handle;
          unsubscribeUpdates = handle.onUpdate(() => {
            if (disposed || !handle) return;
            mutationRef.current += 1;
            try {
              setSheetInfo(handle.sheetInfo());
              setRevision((current) => current + 1);
              setStaleFor({});
              refreshProposals();
              setError(null);
            } catch (e) {
              setError(e instanceof Error ? e.message : String(e));
            }
          });
          pendingSheetViewRef.current = true;
          setSheetInfo(handle.sheetInfo());
          setSelection(selectionAt({ row: 0, col: 0 }));
          setCollaborationReplica(handle);
          setError(null);
          refreshProposals();
          const opened = handle;
          const cleanup = onReadyRef.current?.({
            clearSelection,
            commands: commandController.store,
            handle,
            refreshProposals,
            focus: () => scrollRef.current?.focus(),
            save: () => {
              if (hasRejectedRef.current()) throw new XlsxSaveRefusedError('input-failed');
              if (coordinator.pending) throw new XlsxSaveRefusedError('input-pending');
              if (!settlePendingEditsRef.current()) throw new XlsxSaveRefusedError('input-failed');
              return opened.save();
            },
            selectCells,
            version: () => afterInput(opened, () => opened.version()),
            readCells: (request) => afterInput(opened, () => opened.readCells(request)),
            findText: (request) => afterInput(opened, () => opened.findText(request)),
            validateEdits: async (request) => {
              if (readOnlyRef.current && handleRef.current === opened) {
                return readOnlyRefusal(opened);
              }
              return afterInput(opened, () =>
                readOnlyRef.current ? readOnlyRefusal(opened) : opened.validateEdits(request)
              );
            },
            applyEdits: async (request) => {
              if (readOnlyRef.current && handleRef.current === opened) {
                return readOnlyRefusal(opened);
              }
              return afterInput(opened, () => {
                if (readOnlyRef.current) return readOnlyRefusal(opened);
                const result = opened.applyEdits(request);
                if (result.ok && result.applied) {
                  onChangeRef.current?.();
                  commandController.refresh();
                }
                return result;
              });
            },
          });
          if (typeof cleanup === 'function') cleanupReady = cleanup;
        } catch (e) {
          runReadyCleanup();
          unsubscribeUpdates();
          unsubscribeUpdates = () => {};
          handle?.dispose();
          handle = null;
          handleRef.current = null;
          setSheetInfo(null);
          setSelection(null);
          setError(e instanceof Error ? e.message : String(e));
        }
      },
      (e: unknown) => {
        if (disposed) return;
        handleRef.current = null;
        setSheetInfo(null);
        setSelection(null);
        setError(e instanceof Error ? e.message : String(e));
      }
    );
    return () => {
      disposed = true;
      runReadyCleanup();
      unsubscribeUpdates();
      handle?.dispose();
      handleRef.current = null;
    };
  }, [
    file,
    collaborationEnabled,
    collaborationClientId,
    collaborationInitialUpdate,
    clearSelection,
    afterInput,
    commandController,
    refreshProposals,
    selectCells,
  ]);

  useEffect(() => {
    if (!sheetInfo || !pendingSheetViewRef.current) return;
    const scroll = scrollRef.current;
    if (!scroll) return;
    pendingSheetViewRef.current = false;
    scroll.scrollLeft = sheetInfo.initialScrollX * zoom;
    scroll.scrollTop = sheetInfo.initialScrollY * zoom;
  }, [sheetInfo, zoom]);

  useEffect(() => {
    if (!collaborationOnReplica || !collaborationReplica) return;
    collaborationOnReplica(collaborationReplica);
    return () => collaborationOnReplica(null);
  }, [collaborationOnReplica, collaborationReplica]);

  useEffect(() => {
    setAwarenessPeers([]);
    if (!collaborationProvider) return;
    return collaborationProvider.onAwareness((peers) => setAwarenessPeers([...peers]));
  }, [collaborationProvider]);

  useEffect(() => {
    return () => collaborationProvider?.setCursor(null);
  }, [collaborationProvider]);

  useEffect(() => {
    if (!collaborationProvider) return;
    const sheet = sheetInfo?.sheetIds[activeSheet];
    if (!selection || !sheet) {
      collaborationProvider.setCursor(null);
      return;
    }
    collaborationProvider.setCursor({
      sheet,
      anchor: { ...selection.anchor },
      head: { ...selection.focus },
    });
  }, [collaborationProvider, selection, sheetInfo, activeSheet, revision]);

  // paint the current scroll window into the canvas and publish the frame for
  // overlays + a11y. reads refs so it stays identity-stable across renders.
  const doPaint = useCallback(() => {
    const scroll = scrollRef.current;
    const canvas = canvasRef.current;
    if (!scroll || !canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    const w = scroll.clientWidth;
    const h = scroll.clientHeight;
    if (w === 0 || h === 0) return;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = `${w}px`;
    canvas.style.height = `${h}px`;
    const logicalWidth = w / zoom;
    const logicalHeight = h / zoom;
    const viewport = {
      x: scroll.scrollLeft / zoom,
      y: scroll.scrollTop / zoom,
      width: logicalWidth,
      height: logicalHeight,
    };
    const handle = handleRef.current;
    let dl: DisplayList;
    if (handle) {
      try {
        dl = handle.displayList(viewport);
      } catch (paintError) {
        frameRef.current = null;
        paintedRef.current = null;
        setFrame(null);
        setRenderError(paintError instanceof Error ? paintError.message : String(paintError));
        return;
      }
    } else {
      dl = buildDemoDisplayList(logicalWidth, logicalHeight, t('editor.demoCellText'));
    }
    let nextMergedRanges: readonly MergedRange[] = [];
    const grid = dl.grid;
    const rows = (grid?.rowOffsets.length ?? 0) - 1;
    const columns = (grid?.colOffsets.length ?? 0) - 1;
    if (handle && grid && rows > 0 && columns > 0) {
      try {
        const from = handle.cell(activeSheet, grid.startRow, grid.startCol).a1;
        const to = handle.cell(
          activeSheet,
          grid.startRow + rows - 1,
          grid.startCol + columns - 1
        ).a1;
        nextMergedRanges = handle
          .mergedRanges(activeSheet, `${from}:${to}`)
          .slice(0, MAX_OVERLAY_MERGED_RANGES);
      } catch {}
    }
    paintDisplayList(ctx, dl, dpr * zoom);
    frameRef.current = dl;
    paintedRef.current = { frame: dl, zoom };
    setRenderError(null);
    setVisibleMergedRanges(nextMergedRanges);
    setFrame(dl);
    if (!handle) return;
    const painted: PaintMark = {
      generation: generationRef.current,
      mutation: mutationRef.current,
      view: `${activeSheet}:${zoom}`,
    };
    paintMarkRef.current = painted;
    const waiters = paintWaitersRef.current;
    paintWaitersRef.current = waiters.filter((waiter) => !covers(painted, waiter.mark));
    for (const waiter of waiters) if (covers(painted, waiter.mark)) waiter.resolve(true);
  }, [activeSheet, t, zoom]);

  // paint loop: repaint on scroll/resize (rAF-coalesced) and whenever the open
  // workbook, active sheet, or a mutation (revision) changes the pixels.
  useEffect(() => {
    const scroll = scrollRef.current;
    if (!scroll) return;
    const schedulePaint = () => {
      if (rafRef.current != null) return;
      rafRef.current = requestAnimationFrame(() => {
        rafRef.current = null;
        doPaint();
      });
    };
    doPaint();
    scroll.addEventListener('scroll', schedulePaint, { passive: true });
    const observer = new ResizeObserver(schedulePaint);
    observer.observe(scroll);
    return () => {
      scroll.removeEventListener('scroll', schedulePaint);
      observer.disconnect();
      if (rafRef.current != null) {
        cancelAnimationFrame(rafRef.current);
        rafRef.current = null;
      }
    };
  }, [doPaint, sheetInfo, error, revision]);

  // true once the canvas paints every committed change; false if none does in time.
  const afterPaint = useCallback((): Promise<boolean> => {
    const mark: PaintMark = {
      generation: generationRef.current,
      mutation: mutationRef.current,
      view: `${activeSheetRef.current}:${zoomRef.current}`,
    };
    const painted = paintMarkRef.current;
    if (painted && covers(painted, mark)) return Promise.resolve(true);
    return new Promise<boolean>((resolve) => {
      const waiter = {
        mark,
        resolve: (painted: boolean) => {
          clearTimeout(timer);
          resolve(painted);
        },
      };
      const timer = setTimeout(() => {
        paintWaitersRef.current = paintWaitersRef.current.filter((entry) => entry !== waiter);
        resolve(false);
      }, PAINT_SETTLE_MS);
      paintWaitersRef.current.push(waiter);
    });
  }, []);

  // read the focused cell's editable text for the name box + formula bar; reruns
  // as the selection moves or the workbook mutates.
  useEffect(() => {
    const handle = handleRef.current;
    if (!handle || !selection || !sheetInfo) {
      setFocusedCell(null);
      return;
    }
    try {
      setFocusedCell(handle.cell(activeSheet, selection.focus.row, selection.focus.col));
    } catch {
      setFocusedCell(null);
    }
  }, [selection, sheetInfo, activeSheet, revision]);

  // clear a stuck drag if the mouse is released outside the grid.
  useEffect(() => {
    const stop = () => {
      draggingRef.current = false;
      setDragging(false);
    };
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, []);

  // rebuilt from the live frame so the offscreen mirror never lags a mutation;
  // the visible window is small, so a rebuild per paint frame is cheap enough.
  const a11yGrid = useMemo(() => {
    if (!frame || !sheetInfo) return null;
    return buildA11yGrid(frame, selection, sheetInfo.sheetNames[activeSheet] ?? '', {
      gridLabel: t('a11y.gridLabel'),
      rowHeaderLabel: t('a11y.rowHeaderLabel'),
      columnHeaderLabel: t('a11y.columnHeaderLabel'),
      cellLabel: t('a11y.cellLabel'),
      cellLabelSelected: t('a11y.cellLabelSelected'),
      emptyCellLabel: t('a11y.emptyCellLabel'),
      emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
    });
  }, [frame, selection, sheetInfo, activeSheet, t]);

  // preventScroll everywhere: the sticky overlay host sits below the full-height
  // canvas in flow, so a plain focus() scrolls the grid to bring it into view.
  const focusContainer = useCallback(() => {
    scrollRef.current?.focus({ preventScroll: true });
  }, []);

  // focus the in-cell editor when it opens, without scrolling the grid.
  useEffect(() => {
    if (editing) editorInputRef.current?.focus({ preventScroll: true });
  }, [editing]);

  // fold a mutation result back into state and queue a repaint. re-reads the
  // pending proposals because structural ops and undo/redo can drop them.
  const applyResult = useCallback(
    (result: EditResult) => {
      mutationRef.current += 1;
      setSheetInfo(result.sheetInfo);
      setRevision((r) => r + 1);
      refreshProposals();
      if (result.applied) onChangeRef.current?.();
    },
    [refreshProposals]
  );

  // before a command runs, an arrow burst lands and a composition ends.
  inputHooksRef.current = {
    seal() {
      if (chartDragRef.current) return { refused: 'gesture-active' };
      flushNudgeRef.current();
      const composition = compositionRef.current;
      if (!composition) return {};
      if (composition.source === 'cell') {
        suppressBlurRef.current = true;
        editorInputRef.current?.blur();
        suppressBlurRef.current = false;
      } else {
        suppressFormulaBlurRef.current = true;
        formulaInputRef.current?.blur();
        suppressFormulaBlurRef.current = false;
      }
      return { composition: composition.done };
    },
    sync() {
      const draft = coordinator.draft;
      const input = draft?.source === 'cell' ? editorInputRef.current : formulaInputRef.current;
      if (draft && input && input.value !== draft.value) {
        coordinator.setDraft({ ...draft, value: input.value });
      }
    },
    write(draft) {
      const handle = handleRef.current;
      if (!handle) return false;
      if (readOnlyRef.current) return true;
      try {
        if (handle.cell(draft.sheet, draft.row, draft.col).input !== draft.value) {
          applyResult(handle.editCell(draft.sheet, draft.row, draft.col, draft.value));
        }
        return true;
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
        return false;
      }
    },
    close(draft) {
      if (draft.source === 'formula') {
        setFormulaDraft(null);
        return;
      }
      const focused = document.activeElement === editorInputRef.current;
      suppressBlurRef.current = true;
      editorInputRef.current?.blur();
      suppressBlurRef.current = false;
      setEditing(null);
      if (focused) focusContainer();
    },
    restore(draft) {
      if (draft.sheet !== activeSheetRef.current) return false;
      showDraft(draft);
      return true;
    },
  };

  hasRejectedRef.current = () =>
    coordinator.rejected.some((entry) => entry.generation === generationRef.current);

  settlePendingEditsRef.current = () => {
    if (!handleRef.current) return false;
    flushNudgeRef.current();
    chartDragRef.current = null;
    setChartDragOffset(null);
    setDragging(false);
    return coordinator.settle();
  };

  const selectedRangeA1 = useCallback(
    (target: Selection): string | null => {
      const handle = handleRef.current;
      if (!handle) return null;
      const range = normalizeRange(target);
      const from = handle.cell(activeSheet, range.top, range.left).a1;
      const to = handle.cell(activeSheet, range.bottom, range.right).a1;
      return `${from}:${to}`;
    },
    [activeSheet]
  );

  const limits = useCallback((): SelectionLimits => {
    return deriveLimits(
      frameRef.current,
      sheetInfo!,
      (scrollRef.current?.clientHeight ?? 0) / zoom
    );
  }, [sheetInfo, zoom]);

  // map a viewport-local pointer event to a sheet cell via the frame geometry.
  const pointToCell = useCallback(
    (clientX: number, clientY: number): CellAddr | null => {
      const canvas = canvasRef.current;
      const grid = frameRef.current?.grid;
      if (!canvas || !grid) return null;
      const rect = canvas.getBoundingClientRect();
      return cellAtPoint(grid, (clientX - rect.left) / zoom, (clientY - rect.top) / zoom);
    },
    [zoom]
  );

  // which chart a pointer event lands on. containment runs over the regions the
  // painted frame published — engine geometry, engine clipping, engine paint
  // order — so no geometry is rebuilt here and the answer cannot drift from the
  // pixels by a scroll frame or a mutation the canvas has not drawn yet.
  const pointToChart = useCallback((clientX: number, clientY: number): ChartRegion | null => {
    const canvas = canvasRef.current;
    const painted = paintedRef.current;
    if (!canvas || !painted) return null;
    const rect = canvas.getBoundingClientRect();
    return chartRegionAtPoint(
      painted.frame.charts,
      (clientX - rect.left) / painted.zoom,
      (clientY - rect.top) / painted.zoom
    );
  }, []);

  const selectedChartRegion = useMemo(
    () =>
      selectedChart
        ? (frame?.charts?.find((chart) => chart.id === selectedChart.id) ?? null)
        : null,
    [frame, selectedChart]
  );

  // a selected chart that scrolled out of the painted frame has no outline to
  // show, so drop it rather than leave an invisible selection swallowing keys.
  useEffect(() => {
    if (!selectedChart || !frame) return;
    if (frame.charts?.some((chart) => chart.id === selectedChart.id)) return;
    // land what the burst already earned before the selection goes away, and
    // drop any armed drag with it.
    flushNudgeRef.current();
    chartDragRef.current = null;
    setChartDragOffset(null);
    setSelectedChart(null);
  }, [frame, selectedChart]);

  // slide a chart through the engine's edit path, so the new anchor is
  // undoable and reaches the drawing part on save. it lands in input order.
  const moveChartBy = useCallback(
    (id: string, dx: number, dy: number) => {
      const handle = handleRef.current;
      if (!handle || readOnly || (dx === 0 && dy === 0)) return;
      const sheet = activeSheet;
      void coordinator.input(() => {
        if (handleRef.current !== handle || readOnlyRef.current) return;
        try {
          applyResult(handle.moveChart(sheet, id, dx, dy));
          return true;
        } catch (e) {
          setError(e instanceof Error ? e.message : String(e));
          return false;
        }
      });
    },
    [activeSheet, applyResult, readOnly, coordinator]
  );

  // land a run of arrow nudges as one edit. key repeat fires as fast as the os
  // pleases and every move is a whole-workbook semantic sync, so presses only
  // accumulate a local delta and preview it; the burst lands once it settles.
  const flushNudge = useCallback(() => {
    if (nudgeTimerRef.current != null) {
      clearTimeout(nudgeTimerRef.current);
      nudgeTimerRef.current = null;
    }
    const pending = nudgeRef.current;
    nudgeRef.current = null;
    setNudgeOffset(null);
    if (!pending) return;
    moveChartBy(pending.id, pending.dx, pending.dy);
  }, [moveChartBy]);

  flushNudgeRef.current = flushNudge;

  const nudgeChart = useCallback((id: string, dx: number, dy: number) => {
    const pending = nudgeRef.current;
    if (pending && pending.id !== id) flushNudgeRef.current();
    const base = nudgeRef.current ?? { id, dx: 0, dy: 0 };
    const next = { id, dx: base.dx + dx, dy: base.dy + dy };
    nudgeRef.current = next;
    setNudgeOffset({ x: next.dx, y: next.dy });
    if (nudgeTimerRef.current != null) clearTimeout(nudgeTimerRef.current);
    nudgeTimerRef.current = setTimeout(() => {
      nudgeTimerRef.current = null;
      // a burst that returns to where it started is not an edit.
      const settled = nudgeRef.current;
      nudgeRef.current = null;
      setNudgeOffset(null);
      if (settled && (settled.dx !== 0 || settled.dy !== 0)) {
        moveChartByRef.current(settled.id, settled.dx, settled.dy);
      }
    }, CHART_NUDGE_SETTLE_MS) as unknown as number;
  }, []);

  const moveChartByRef = useRef(moveChartBy);
  moveChartByRef.current = moveChartBy;

  // commit a chart drag on release, as one edit for the whole gesture. a drag
  // the window loses (focus leaves mid-gesture) is dropped, not committed
  // later against whatever the pointer has since moved over.
  useEffect(() => {
    const commit = (event: MouseEvent) => {
      const drag = chartDragRef.current;
      // any release ends the gesture; only a primary one lands it, so a
      // non-primary release cannot leave the drag armed for a later mouseup.
      chartDragRef.current = null;
      setChartDragOffset(null);
      if (!drag || event.button !== 0) return;
      moveChartBy(
        drag.id,
        (event.clientX - drag.clientX) / zoom,
        (event.clientY - drag.clientY) / zoom
      );
    };
    const cancel = () => {
      chartDragRef.current = null;
      setChartDragOffset(null);
      flushNudgeRef.current();
    };
    window.addEventListener('mouseup', commit);
    window.addEventListener('blur', cancel);
    return () => {
      window.removeEventListener('mouseup', commit);
      window.removeEventListener('blur', cancel);
    };
  }, [moveChartBy, zoom]);

  // on unmount this only stops the timer: cleanups run in hook order, so the
  // open effect above has already disposed the workbook and a flush here would
  // find no handle to move a chart on. an unfinished burst is dropped.
  useEffect(
    () => () => {
      if (nudgeTimerRef.current != null) clearTimeout(nudgeTimerRef.current);
      nudgeTimerRef.current = null;
      nudgeRef.current = null;
    },
    []
  );

  const openEditor = useCallback(
    (seed?: string) => {
      const handle = handleRef.current;
      const selection = selectionRef.current;
      if (!handle || !selection || readOnly) return;
      const { row, col } = selection.focus;
      let value = seed ?? '';
      if (seed === undefined) {
        try {
          value = handle.cell(activeSheetRef.current, row, col).input;
        } catch {
          value = '';
        }
      }
      setEditing({ row, col, value });
      coordinator.setDraft(draftFor('cell', row, col, value));
    },
    [selectionRef, activeSheetRef, readOnly, coordinator, draftFor]
  );

  // commit the open editor, optionally stepping the selection like excel. the
  // write lands in input order; a write that fails keeps the editor open.
  const commitEditor = useCallback(
    (move?: Direction) => {
      if (!handleRef.current || !editing || readOnly) return;
      const { row, col, value } = editing;
      const live = coordinator.draft;
      const draft =
        live?.source === 'cell' && live.row === row && live.col === col
          ? live
          : draftFor('cell', row, col, value);
      if (!coordinator.submit(draft)) return;
      suppressBlurRef.current = true;
      setEditing(null);
      const base = selectionAt({ row, col });
      setSelection(move ? moveFocus(base, move, { limits: limits() }) : base);
      focusContainer();
      showRejected(activeSheetRef.current);
    },
    [editing, limits, focusContainer, readOnly, coordinator, draftFor, showRejected]
  );

  const cancelEditor = useCallback(() => {
    suppressBlurRef.current = true;
    const draft = coordinator.draft;
    if (draft?.source === 'cell') coordinator.discard(draft);
    coordinator.setDraft(null);
    setEditing(null);
    focusContainer();
    showRejected(activeSheetRef.current);
  }, [focusContainer, coordinator, showRejected]);

  const captureTarget = useCallback(() => {
    const handle = handleRef.current;
    const selection = selectionRef.current;
    if (!handle || !selection) return null;
    return { handle, sheet: activeSheetRef.current, range: normalizeRange(selection) };
  }, [selectionRef, activeSheetRef]);

  // an input write clearing `target`, while its workbook is still open and writable.
  const writeClear = useCallback(
    ({ handle, sheet, range }: CellTarget) => {
      if (handleRef.current !== handle || readOnlyRef.current) return;
      const edits: CellInputEdit[] = [];
      for (let row = range.top; row <= range.bottom; row++) {
        for (let col = range.left; col <= range.right; col++) edits.push({ row, col, input: '' });
      }
      try {
        applyResult(handle.editCells(sheet, edits));
        return true;
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
        return false;
      }
    },
    [applyResult]
  );

  const clearCells = useCallback(() => {
    const target = captureTarget();
    if (target && !readOnly) void coordinator.input(() => writeClear(target));
  }, [captureTarget, writeClear, readOnly, coordinator]);

  const copyTarget = useCallback(async ({ handle, sheet, range }: CellTarget) => {
    try {
      const from = handle.cell(sheet, range.top, range.left).a1;
      const to = handle.cell(sheet, range.bottom, range.right).a1;
      const cells = handle.rangeCells(sheet, `${from}:${to}`);
      const tsv = toTsv(
        cells.map((row) => row.map((c) => ({ input: c.input, isFormula: c.isFormula })))
      );
      await navigator.clipboard.writeText(tsv);
      return true;
    } catch {
      // clipboard denied or read failed — nothing to paste, leave state as-is.
      return false;
    }
  }, []);

  const copySelection = useCallback(async () => {
    const target = captureTarget();
    if (target) await copyTarget(target);
  }, [captureTarget, copyTarget]);

  // the cut copies at once and takes its place in input order: it clears what was selected
  // when it was accepted, once the clipboard write settles, and only while that workbook is
  // still open and writable.
  const cutSelection = useCallback(() => {
    const target = captureTarget();
    if (!target) return;
    const copied = copyTarget(target);
    if (readOnly) return;
    void coordinator.input(async () => ((await copied) ? writeClear(target) : undefined));
  }, [captureTarget, copyTarget, writeClear, readOnly, coordinator]);

  // the clipboard is read while the key press still grants access; the write
  // takes its place in input order.
  const pasteSelection = useCallback(() => {
    const handle = handleRef.current;
    const selection = selectionRef.current;
    if (!handle || !selection || readOnly) return Promise.resolve();
    const reading = (async () => {
      try {
        return await navigator.clipboard.readText();
      } catch {
        return null;
      }
    })();
    const sheet = activeSheetRef.current;
    const r = normalizeRange(selection);
    return coordinator.input(async () => {
      const text = await reading;
      if (text === null || readOnlyRef.current || handleRef.current !== handle) return;
      const grid = fromTsv(text);
      if (grid.length === 0) return;
      const edits: CellInputEdit[] = [];
      let width = 1;
      grid.forEach((rowArr, dr) => {
        width = Math.max(width, rowArr.length);
        rowArr.forEach((input, dc) => edits.push({ row: r.top + dr, col: r.left + dc, input }));
      });
      try {
        applyResult(handle.editCells(sheet, edits));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
        return false;
      }
      if (selectionRef.current === selection && activeSheetRef.current === sheet) {
        setSelection({
          anchor: { row: r.top, col: r.left },
          focus: { row: r.top + grid.length - 1, col: r.left + width - 1 },
        });
      }
      return true;
    });
  }, [selectionRef, activeSheetRef, applyResult, readOnly, coordinator, setSelection]);

  useEffect(() => {
    const handle = handleRef.current;
    if (!handle || !selection || !capturedFormat || dragging || readOnly) return;
    const normalized = normalizeRange(selection);
    const key = `${activeSheet}:${normalized.top}:${normalized.left}:${normalized.bottom}:${normalized.right}`;
    if (key === paintSourceRef.current) return;
    const range = selectedRangeA1(selection);
    if (!range) return;
    const sheet = activeSheet;
    const format = capturedFormat;
    setCapturedFormat(null);
    paintSourceRef.current = null;
    void coordinator.input(() => {
      if (handleRef.current !== handle || readOnlyRef.current) return;
      try {
        applyResult(handle.applyFormat(sheet, range, format));
        focusContainer();
        return true;
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
        return false;
      }
    });
  }, [
    selection,
    capturedFormat,
    dragging,
    activeSheet,
    selectedRangeA1,
    applyResult,
    focusContainer,
    readOnly,
    coordinator,
  ]);

  const deliverSave = useCallback(
    (bytes: Uint8Array) => {
      if (onSave) onSave(bytes);
      else downloadBytes(bytes, fileName ?? 'workbook.xlsx', XLSX_MIME);
    },
    [onSave, fileName]
  );

  // render the current scroll window to png via the raster backend and download
  // it — the same display list the canvas paints, rasterized in the core.
  const exportPng = useCallback(() => {
    const handle = handleRef.current;
    const scroll = scrollRef.current;
    if (!handle || !scroll) throw new Error('the workbook is not open');
    const png = handle.renderPng({
      x: scroll.scrollLeft / zoom,
      y: scroll.scrollTop / zoom,
      width: scroll.clientWidth / zoom,
      height: scroll.clientHeight / zoom,
    });
    downloadBytes(png, pngName(fileName), 'image/png');
  }, [fileName, zoom]);

  // commit the formula bar draft to the focused cell.
  const commitFormula = useCallback(
    (move?: Direction) => {
      const draft = coordinator.draft;
      if (!handleRef.current || draft?.source !== 'formula' || readOnly) return;
      if (!coordinator.submit(draft)) return;
      setFormulaDraft(null);
      if (move) setSelection((prev) => (prev ? moveFocus(prev, move, { limits: limits() }) : prev));
      showRejected(activeSheetRef.current);
    },
    [coordinator, limits, readOnly, showRejected]
  );

  // grid-level keyboard: command shortcuts belong to the editor's dispatcher,
  // clipboard keys stay here, then the pure selection reducer.
  const onKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      const handle = handleRef.current;
      if (!handle || !selection || !sheetInfo || editing) return;
      if (commandForEvent(e)) return;
      const mod = e.metaKey || e.ctrlKey;
      const lower = e.key.toLowerCase();

      // a selected chart owns the keyboard: arrows nudge it, escape drops it,
      // and nothing else reaches the cells hidden behind it — the grid overlay
      // is suppressed while it is selected, so a delete or a keystroke there
      // would edit a target the user cannot see.
      if (selectedChart) {
        if (e.key === 'Escape') {
          // escape cancels the whole gesture, pointer or keyboard: an armed
          // drag must not land a move the user just abandoned.
          chartDragRef.current = null;
          setChartDragOffset(null);
          nudgeRef.current = null;
          setNudgeOffset(null);
          if (nudgeTimerRef.current != null) {
            clearTimeout(nudgeTimerRef.current);
            nudgeTimerRef.current = null;
          }
          setSelectedChart(null);
          e.preventDefault();
          return;
        }
        const nudge = mod ? undefined : CHART_NUDGE_KEYS[e.key];
        if (nudge) {
          const step = CHART_NUDGE_PX * (e.shiftKey ? CHART_NUDGE_MULTIPLIER : 1);
          if (!readOnly && selectedChart.movable) {
            nudgeChart(selectedChart.id, nudge[0] * step, nudge[1] * step);
          }
          e.preventDefault();
        }
        return;
      }

      if (mod) {
        if (lower === 'c') {
          void copySelection();
          e.preventDefault();
          return;
        }
        if (lower === 'v') {
          void pasteSelection();
          e.preventDefault();
          return;
        }
        if (lower === 'x') {
          cutSelection();
          e.preventDefault();
          return;
        }
      }

      const action = selectionKeyReducer(
        selection,
        {
          key: e.key,
          shiftKey: e.shiftKey,
          metaKey: e.metaKey,
          ctrlKey: e.ctrlKey,
          altKey: e.altKey,
        },
        limits()
      );
      switch (action.type) {
        case 'move':
          setSelection(action.selection);
          e.preventDefault();
          break;
        case 'startEdit':
          openEditor(action.initialInput);
          e.preventDefault();
          break;
        case 'clear':
          clearCells();
          e.preventDefault();
          break;
        case 'none':
          break;
      }
    },
    [
      selection,
      sheetInfo,
      editing,
      limits,
      copySelection,
      pasteSelection,
      cutSelection,
      openEditor,
      clearCells,
      selectedChart,
      nudgeChart,
      readOnly,
    ]
  );

  const onMouseDown = useCallback(
    (e: React.MouseEvent) => {
      // the editor input is a dom overlay above the canvas, so a press inside
      // it is the editor's own — it places a caret, and must not commit, reach
      // a chart painted under it, or move the grid.
      const dismissesEditor = editing != null;
      if (dismissesEditor && editorInputRef.current?.contains(e.target as Node)) return;
      // a release the window never saw would otherwise leave the last gesture
      // armed and commit its accumulated delta on some later, unrelated mouseup.
      chartDragRef.current = null;
      // a pointer gesture ends any arrow burst, so the two never interleave.
      flushNudgeRef.current();

      // a chart takes the press as a whole object: it finishes any open edit,
      // then selects, leaving the grid selection where the edit left it. the
      // commit is stated here rather than left to the blur `focusContainer`
      // triggers, so both branches finish the edit for the same visible reason.
      const chart = pointToChart(e.clientX, e.clientY);
      if (chart) {
        if (dismissesEditor) commitEditor();
        setSelectedChart({ id: chart.id, movable: !readOnly && chart.movable });
        // only a primary press starts a drag: a right-press opens a context
        // menu whose release this window never sees, and an armed drag would
        // then land on whatever the next unrelated click released over.
        if (!readOnly && chart.movable && e.button === 0) {
          chartDragRef.current = { id: chart.id, clientX: e.clientX, clientY: e.clientY };
        }
        setChartDragOffset(null);
        // no click start: a press on a chart is not a cell click, so it must
        // not follow a hyperlink in whatever cell sits behind it.
        clickStartRef.current = null;
        focusContainer();
        e.preventDefault();
        return;
      }

      setSelectedChart(null);
      setChartDragOffset(null);
      // an async reopen leaves the previous frame painted, so a chart here can
      // outlive `selection` for a moment. that is fine: the chart branch has
      // already returned above, and selecting a chart the user can still see is
      // what should happen. only the cell path needs a selection to extend.
      if (!selection) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) return;
      if (dismissesEditor) commitEditor();
      if (e.shiftKey) setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
      else setSelection(selectionAt(addr));
      // no click start: a press that dismissed the editor must not also follow a
      // hyperlink in the cell it selects.
      clickStartRef.current = dismissesEditor ? null : addr;
      draggingRef.current = true;
      setDragging(true);
      focusContainer();
    },
    [editing, selection, pointToCell, pointToChart, limits, focusContainer, commitEditor, readOnly]
  );

  const onMouseMove = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      const chartDrag = chartDragRef.current;
      if (chartDrag) {
        // the button came back up somewhere this window never heard about, so
        // the gesture is over; do not keep accumulating it.
        if (e.buttons === 0) {
          chartDragRef.current = null;
          setChartDragOffset(null);
          return;
        }
        e.currentTarget.style.cursor = 'move';
        // a tooltip from whatever cell was hovered before must not ride along.
        e.currentTarget.title = '';
        setChartDragOffset({
          x: e.clientX - chartDrag.clientX,
          y: e.clientY - chartDrag.clientY,
        });
        return;
      }
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) {
        if (!draggingRef.current) {
          e.currentTarget.style.cursor = 'default';
          e.currentTarget.title = '';
        }
        return;
      }
      if (!draggingRef.current) {
        const chart = pointToChart(e.clientX, e.clientY);
        if (chart) {
          // a pinned chart still selects on click, so it must not read as a cell.
          e.currentTarget.style.cursor = !readOnly && chart.movable ? 'move' : 'pointer';
          e.currentTarget.title = '';
          return;
        }
        const hyperlink = frameRef.current
          ? hyperlinkAtCell(frameRef.current, addr.row, addr.col)
          : null;
        e.currentTarget.style.cursor = hyperlink ? 'pointer' : 'default';
        e.currentTarget.title = hyperlink?.tooltip ?? '';
        return;
      }
      setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
    },
    [pointToCell, pointToChart, limits, readOnly]
  );

  const activateHyperlink = useCallback(
    (addr: CellAddr): boolean => {
      const handle = handleRef.current;
      const currentFrame = frameRef.current;
      if (!handle || !currentFrame || !sheetInfo) return false;
      const hyperlink = hyperlinkAtCell(currentFrame, addr.row, addr.col);
      if (!hyperlink) return false;
      const href = safeExternalHyperlink(hyperlink);
      if (href) {
        window.open(href, '_blank', 'noopener,noreferrer');
        return true;
      }
      if (!hyperlink.location) return true;
      const currentName = sheetInfo.sheetNames[activeSheet] ?? '';
      const destination = parseHyperlinkLocation(hyperlink.location, currentName);
      if (!destination) return true;
      const targetSheet = sheetInfo.sheetNames.findIndex(
        (name) => name.toLowerCase() === destination.sheetName.toLowerCase()
      );
      if (targetSheet < 0) return true;
      try {
        handle.setActiveSheet(targetSheet);
        const position = handle.cellPosition(targetSheet, destination.row, destination.col);
        setSheetInfo(handle.sheetInfo());
        requestAnimationFrame(() => {
          const scroll = scrollRef.current;
          if (!scroll) return;
          scroll.scrollLeft = position.x * zoom;
          scroll.scrollTop = position.y * zoom;
        });
        setSelection(selectionAt({ row: destination.row, col: destination.col }));
        dropDrafts();
        setCapturedFormat(null);
        paintSourceRef.current = null;
        setError(null);
      } catch (error) {
        setError(error instanceof Error ? error.message : String(error));
      }
      return true;
    },
    [activeSheet, sheetInfo, zoom]
  );

  const onClick = useCallback(
    (e: React.MouseEvent) => {
      if (editing) {
        clickStartRef.current = null;
        return;
      }
      const addr = pointToCell(e.clientX, e.clientY);
      const start = clickStartRef.current;
      clickStartRef.current = null;
      if (!addr || !start || addr.row !== start.row || addr.col !== start.col) return;
      activateHyperlink(addr);
    },
    [activateHyperlink, editing, pointToCell]
  );

  const onDoubleClick = useCallback(
    (e: React.MouseEvent) => {
      if (editing) return;
      if (pointToChart(e.clientX, e.clientY)) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (
        addr &&
        frameRef.current &&
        hyperlinkAtCell(frameRef.current, addr.row, addr.col)
      ) {
        return;
      }
      if (selection) openEditor();
    },
    [editing, openEditor, pointToCell, pointToChart, selection]
  );

  const onMouseLeave = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    e.currentTarget.style.cursor = 'default';
    e.currentTarget.title = '';
  }, []);

  const grid = frame?.grid;
  const renderedSelection = selection
    ? expandRangeToMergedCells(normalizeRange(selection), visibleMergedRanges)
    : null;
  const renderedFocus = selection
    ? expandRangeToMergedCells(
        {
          top: selection.focus.row,
          left: selection.focus.col,
          bottom: selection.focus.row,
          right: selection.focus.col,
        },
        visibleMergedRanges
      )
    : null;
  const selRect = grid && renderedSelection ? rangeRect(grid, renderedSelection) : null;
  const focusRect = grid && renderedFocus ? rangeRect(grid, renderedFocus) : null;
  const editRect = grid && editing ? cellRect(grid, editing.row, editing.col) : null;
  const scaledSelectionRect = selRect ? scaledRect(selRect, zoom) : null;
  const scaledFocusRect = focusRect ? scaledRect(focusRect, zoom) : null;
  const scaledEditRect = editRect ? scaledRect(editRect, zoom) : null;

  // the selected chart's outline, placed from the engine-published region and
  // offset by the live drag so the box tracks the pointer before it commits.
  const chartOutlineRect = selectedChartRegion
    ? scaledRect(selectedChartRegion.rect, zoom)
    : null;

  const spacerWidth = sheetInfo ? sheetInfo.contentWidth * zoom : undefined;
  const spacerHeight = sheetInfo ? sheetInfo.contentHeight * zoom : undefined;
  const formulaValue = formulaDraft ?? focusedCell?.input ?? '';

  // switch sheets: retarget the core, reset scroll + selection, reread info.
  const switchSheet = (index: number) => {
    const handle = handleRef.current;
    if (!handle) return;
    // the burst belongs to the sheet it was typed on.
    if (!settlePendingEditsRef.current()) return;
    try {
      handle.setActiveSheet(index);
      pendingSheetViewRef.current = true;
      setSelection(selectionAt({ row: 0, col: 0 }));
      setSelectedChart(null);
      dropDrafts();
      setCapturedFormat(null);
      paintSourceRef.current = null;
      setSheetInfo(handle.sheetInfo());
      showRejected(index);
      commandController.refresh();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  const bridge: XlsxEditorBridge = {
    handle: () => handleRef.current,
    status: () => (handleRef.current ? 'ready' : file && !error ? 'loading' : 'empty'),
    readOnly: () => readOnlyRef.current,
    collaborative: () => collaborationEnabled,
    mutation: () => mutationRef.current,
    generation: () => generationRef.current,
    view: () => ({
      sheet: activeSheetRef.current,
      selection: selectionRef.current,
      chartSelected: selectedChartRef.current !== null,
      zoom: zoomRef.current,
      capturedFormat: capturedFormatRef.current,
      borderStyle: borderStyleChoiceRef.current,
      borderColor: borderColorChoiceRef.current,
      proposals: proposalsRef.current,
      proposalsAvailable,
      proposalsPanelOpen: proposalsPanelOpenRef.current,
      pngExport: pngExportAvailable,
    }),
    coordinator,
    i18n: () => i18n,
    translate: t,
    apply: applyResult,
    fail: (cause) => setError(cause instanceof Error ? cause.message : String(cause)),
    setZoom,
    setProposalsPanelOpen,
    setCapturedFormat: (format, source) => {
      setCapturedFormat(format);
      paintSourceRef.current = source;
    },
    setBorderStyle: (style) => {
      borderStyleChoiceRef.current = style;
    },
    setBorderColor: (color) => {
      borderColorChoiceRef.current = color;
    },
    markStale: (proposalId, cells) =>
      setStaleFor(({ [proposalId]: _dropped, ...rest }) =>
        cells ? { ...rest, [proposalId]: cells } : rest
      ),
    refreshProposals,
    deliver: deliverSave,
    exportPng,
    afterPaint,
    focusGrid: focusContainer,
  };
  useXlsxCommandBinding(commandController, bridge);
  useCommandShortcuts(commandController, rootRef);

  const startComposition = (source: InputDraft['source']) => {
    let settle: (ended: boolean) => void = () => {};
    const done = new Promise<boolean>((resolve) => {
      settle = (ended) => setTimeout(() => resolve(ended), 0);
    });
    compositionRef.current = { source, done, settle };
  };
  const endComposition = () => {
    compositionRef.current?.settle(true);
    compositionRef.current = null;
  };
  // a composition whose input goes away never ends; commands waiting on it fail.
  const abandonComposition = useCallback((source: InputDraft['source']) => {
    const composition = compositionRef.current;
    if (composition?.source !== source) return;
    composition.settle(false);
    compositionRef.current = null;
  }, []);
  const attachCellInput = useCallback(
    (element: HTMLInputElement | null) => {
      editorInputRef.current = element;
      if (!element) abandonComposition('cell');
    },
    [abandonComposition]
  );
  const attachFormulaInput = useCallback(
    (element: HTMLInputElement | null) => {
      formulaInputRef.current = element;
      if (!element) abandonComposition('formula');
    },
    [abandonComposition]
  );

  const formulaWritable = !readOnly && selection !== null && selectedChart === null;
  const formulaBar: FormulaBarBinding = {
    a1: focusedCell?.a1 ?? '',
    value: formulaValue,
    disabled: !sheetInfo || !selection,
    readOnly: !formulaWritable,
    inputRef: attachFormulaInput,
    onChange: (value) => {
      const current = selectionRef.current;
      if (!formulaWritable || readOnlyRef.current || !current || selectedChartRef.current) return;
      setFormulaDraft(value);
      coordinator.setDraft(draftFor('formula', current.focus.row, current.focus.col, value));
    },
    onCommit: (move) => {
      commitFormula(move);
      focusContainer();
    },
    onCancel: () => {
      const draft = coordinator.draft;
      if (draft?.source === 'formula') {
        coordinator.discard(draft);
        coordinator.setDraft(null);
      }
      setFormulaDraft(null);
      focusContainer();
      showRejected(activeSheet);
    },
    onBlur: () => {
      if (!suppressFormulaBlurRef.current) commitFormula();
    },
    onCompositionStart: () => startComposition('formula'),
    onCompositionEnd: endComposition,
  };

  const defaultToolbar = (
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar />
      <div style={xlsxToolbarStyles.rail} role="group" aria-label={t('toolbar.formulaBarLabel')}>
        <ToolbarGroup
          style={{ ...xlsxToolbarStyles.group, paddingLeft: 0 }}
          label={t('toolbar.fileActionsLabel')}
        >
          <ToolbarCommandButton id="save">
            <ToolbarIcon name="save" size={18} />
          </ToolbarCommandButton>
          <ToolbarCommandButton id="exportPng">
            <ToolbarIcon name="image" size={18} />
          </ToolbarCommandButton>
        </ToolbarGroup>
        <EditorToolbar.FormulaBar />
        <PresenceStrip
          peers={awarenessPeers}
          sheetIds={sheetInfo?.sheetIds ?? []}
          sheetNames={sheetInfo?.sheetNames ?? []}
          activeSheet={activeSheet}
        />
        {proposalsAvailable && (
          <div style={xlsxToolbarStyles.proposals}>
            <ProposalsButton />
          </div>
        )}
      </div>
    </EditorToolbar>
  );
  const toolbarContent = !showToolbar
    ? null
    : toolbar === undefined
      ? readOnly
        ? null
        : defaultToolbar
      : toolbar;

  const editor = (
    <div
      ref={rootRef}
      className={className}
      role="application"
      aria-label={t('editor.appLabel')}
      style={{
        position: 'relative',
        display: 'flex',
        flexDirection: 'column',
        width: '100%',
        height: '100%',
        minWidth: 0,
        color: '#202124',
        background: '#ffffff',
        fontFamily: 'ui-sans-serif, system-ui, sans-serif',
      }}
    >
      {toolbarContent != null && (
        <div
          ref={setToolbarElement}
          data-testid="xlsx-toolbar"
          style={toolbar === undefined ? xlsxToolbarStyles.shell : xlsxToolbarStyles.host}
        >
          {toolbarContent}
        </div>
      )}
      {proposalsAvailable && proposalsPanelOpen && (
        <ProposalsPanel
          proposals={proposals}
          staleFor={staleFor}
          style={{
            top: toolbarHeight + 4,
            right: 8,
            width: 'min(320px, calc(100% - 16px))',
            maxHeight: `min(420px, calc(100% - ${toolbarHeight + 12}px))`,
          }}
        />
      )}

      <div
        ref={scrollRef}
        data-testid="xlsx-scroll"
        tabIndex={0}
        onKeyDown={onKeyDown}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseLeave={onMouseLeave}
        onClick={onClick}
        onDoubleClick={onDoubleClick}
        style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}
      >
        <div
          style={{
            position: 'absolute',
            top: 0,
            left: 0,
            width: spacerWidth ?? '100%',
            height: spacerHeight ?? '100%',
          }}
        />
        {/* one sticky layer pins the canvas and overlays to the viewport top-left
            so overlay children share the canvas's coordinate space — a separate
            sticky sibling would sit below the full-height canvas in flow and
            scroll-jump when a child (the in-cell editor) is focused. */}
        <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
          <canvas
            ref={canvasRef}
            style={{ display: 'block', position: 'absolute', top: 0, left: 0 }}
          />

          <div
            data-testid="xlsx-overlay-host"
            style={{
              position: 'absolute',
              top: 0,
              left: 0,
              width: 0,
              height: 0,
              pointerEvents: 'none',
            }}
          >
            {scaledSelectionRect && !selectedChart && (
              <div
                data-testid="xlsx-selection"
                style={{
                  position: 'absolute',
                  left: scaledSelectionRect.x,
                  top: scaledSelectionRect.y,
                  width: scaledSelectionRect.w,
                  height: scaledSelectionRect.h,
                  boxSizing: 'border-box',
                  border: `1px solid ${BRAND}`,
                  background: 'rgba(33, 115, 70, 0.12)',
                }}
              />
            )}
            {chartOutlineRect && (
              <div
                data-testid="xlsx-chart-selection"
                data-chart-id={selectedChart?.id}
                aria-hidden
                style={{
                  position: 'absolute',
                  left:
                    chartOutlineRect.x + (chartDragOffset?.x ?? 0) + (nudgeOffset?.x ?? 0) * zoom,
                  top:
                    chartOutlineRect.y + (chartDragOffset?.y ?? 0) + (nudgeOffset?.y ?? 0) * zoom,
                  width: chartOutlineRect.w,
                  height: chartOutlineRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                  boxShadow: '0 1px 6px rgba(0, 0, 0, 0.25)',
                  background: chartDragOffset ? 'rgba(33, 115, 70, 0.08)' : 'transparent',
                }}
              />
            )}
            {scaledFocusRect && !selectedChart && !editing && (
              <div
                style={{
                  position: 'absolute',
                  left: scaledFocusRect.x,
                  top: scaledFocusRect.y,
                  width: scaledFocusRect.w,
                  height: scaledFocusRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                }}
              />
            )}
            <RemoteSelections
              peers={awarenessPeers}
              grid={grid}
              sheetIds={sheetInfo?.sheetIds ?? []}
              activeSheet={activeSheet}
              zoom={zoom}
              mergedRanges={visibleMergedRanges}
            />
            {!readOnly && editing && scaledEditRect && (
              <input
                ref={attachCellInput}
                data-testid="xlsx-cell-editor"
                value={editing.value}
                onChange={(e) => {
                  const value = e.target.value;
                  setEditing((prev) => (prev ? { ...prev, value } : prev));
                  if (editing) coordinator.setDraft(draftFor('cell', editing.row, editing.col, value));
                }}
                onCompositionStart={() => startComposition('cell')}
                onCompositionEnd={endComposition}
                onKeyDown={(e) => {
                  if (!commandForEvent(e)) e.stopPropagation();
                  if (e.nativeEvent.isComposing || e.keyCode === 229) return;
                  if (e.key === 'Enter') {
                    commitEditor(e.shiftKey ? 'up' : 'down');
                    e.preventDefault();
                  } else if (e.key === 'Tab') {
                    commitEditor(e.shiftKey ? 'left' : 'right');
                    e.preventDefault();
                  } else if (e.key === 'Escape') {
                    cancelEditor();
                    e.preventDefault();
                  }
                }}
                onBlur={() => {
                  if (suppressBlurRef.current) {
                    suppressBlurRef.current = false;
                    return;
                  }
                  commitEditor();
                }}
                style={{
                  position: 'absolute',
                  left: scaledEditRect.x,
                  top: scaledEditRect.y,
                  width: scaledEditRect.w,
                  height: scaledEditRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                  padding: '0 3px',
                  font: `${13 * zoom}px system-ui, sans-serif`,
                  background: '#ffffff',
                  pointerEvents: 'auto',
                  outline: 'none',
                }}
              />
            )}
          </div>
        </div>
      </div>

      {a11yGrid && (
        <>
          <div style={visuallyHidden} role="grid" aria-label={a11yGrid.label}>
            <div role="row">
              <span role="columnheader" />
              {a11yGrid.columnHeaders.map((h) => (
                <span key={h.col} role="columnheader">
                  {h.label}
                </span>
              ))}
            </div>
            {a11yGrid.rows.map((r) => (
              <div key={r.row} role="row">
                <span role="rowheader">{r.header}</span>
                {r.cells.map((c) => (
                  <span key={c.col} role="gridcell" aria-selected={c.selected}>
                    {c.label}
                  </span>
                ))}
              </div>
            ))}
          </div>
          {a11yGrid.charts.map((chart, index) => (
            <div key={`${index}:${chart.label}`} style={visuallyHidden} role="img" aria-label={chart.label} />
          ))}
        </>
      )}

      {renderError && (
        <div
          data-testid="xlsx-render-error"
          role="alert"
          style={{
            position: 'absolute',
            inset: 0,
            display: 'grid',
            placeItems: 'center',
            padding: 16,
            textAlign: 'center',
            color: '#b00020',
            background: '#ffffff',
          }}
        >
          {renderError}
        </div>
      )}

      {error && (
        <div
          data-testid="xlsx-error"
          role="alert"
          style={{
            position: 'absolute',
            inset: 0,
            display: 'grid',
            placeItems: 'center',
            padding: 16,
            textAlign: 'center',
            color: '#b00020',
          }}
        >
          {t('editor.openError')}: {error}
        </div>
      )}

      {sheetInfo && sheetInfo.sheetNames.length > 0 && (
        <div
          data-testid="xlsx-sheet-tabs"
          role="tablist"
          aria-label={t('editor.sheetTabsLabel')}
          style={{
            display: 'flex',
            gap: 2,
            padding: '4px 6px',
            borderTop: '1px solid #e0e0e0',
            background: '#fafafa',
            overflowX: 'auto',
          }}
        >
          {sheetInfo.sheetNames.map((name, i) => {
            const active = i === sheetInfo.activeSheet;
            return (
              <button
                key={i}
                role="tab"
                aria-selected={active}
                onClick={() => switchSheet(i)}
                style={{
                  border: 'none',
                  padding: '4px 12px',
                  cursor: 'pointer',
                  borderBottom: active ? `2px solid ${BRAND}` : '2px solid transparent',
                  fontWeight: active ? 600 : 400,
                  background: active ? '#ffffff' : 'transparent',
                }}
              >
                {name}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );

  return (
    <XlsxCommandContext.Provider value={commandController.store}>
      <EditorChromeContext.Provider value={true}>
        <FormulaBarContext.Provider value={formulaBar}>{editor}</FormulaBarContext.Provider>
      </EditorChromeContext.Provider>
    </XlsxCommandContext.Provider>
  );
}

/** The proposals-panel toggle, with the number of pending proposals. */
function ProposalsButton() {
  const command = useXlsxCommand('proposalsPanel');
  const count = command.state.value ?? 0;
  return (
    <ToolbarButtonBase
      testId="xlsx-proposals-button"
      onClick={() => void command.execute()}
      disabled={!command.state.enabled}
      description={command.state.enabled ? undefined : command.state.disabledReason.message}
      active={command.state.active === true}
      toggle
      ariaExpanded={command.state.active === true}
      title={command.label}
      style={{ width: count > 0 ? 42 : 28 }}
    >
      <ToolbarIcon name="proposals" size={18} />
      {count > 0 && (
        <span data-testid="xlsx-proposals-count" style={xlsxToolbarStyles.count}>
          {count}
        </span>
      )}
    </ToolbarButtonBase>
  );
}
