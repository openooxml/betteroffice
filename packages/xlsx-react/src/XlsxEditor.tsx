/**
 * `<XlsxEditor />` — the editor shell. A dpr-aware canvas paints the grid; DOM
 * overlays (selection marquee, active-cell outline, in-cell editor) sit above it,
 * positioned from the same display-list geometry the painter uses. A top bar
 * holds the name box, formula bar, and save/undo/redo; an offscreen `role=grid`
 * mirror serves screen readers. All compute lives in `@betteroffice/xlsx`; this
 * layer is framework glue — keyboard/mouse wiring, focus flow, and DOM chrome.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildA11yGrid,
  cellAtPoint,
  cellRect,
  extendTo,
  fromTsv,
  initWasm,
  isPngExportAvailable,
  isProposalsAvailable,
  moveFocus,
  normalizeRange,
  openWorkbook,
  paintDisplayList,
  rangeRect,
  selectionAt,
  selectionKeyReducer,
  StaleProposalError,
  toTsv,
} from '@betteroffice/xlsx';
import type {
  CellAddr,
  CellEdit,
  CellInputEdit,
  DisplayList,
  DrawCmd,
  Direction,
  EditResult,
  Proposal,
  Selection,
  SelectionLimits,
  SheetInfo,
  WorkbookHandle,
} from '@betteroffice/xlsx';
import type { CollaborationReplica } from '@betteroffice/xlsx/collaboration';
import type { Translations } from '@betteroffice/xlsx-i18n';
import { LocaleProvider, useTranslation } from './i18n';
import { ProposalDecoration } from './proposals/ProposalDecoration';
import { ProposalsPanel } from './proposals/ProposalsPanel';
import { proposalColor } from './proposals/palette';

/**
 * The imperative surface handed to {@link XlsxEditorProps.onReady}: the open
 * workbook handle plus a `refreshProposals` to re-read the pending list after an
 * external caller (e.g. a demo agent) stages proposals on the same handle.
 */
export interface XlsxEditorApi {
  handle: WorkbookHandle;
  refreshProposals: () => void;
}

export interface XlsxEditorCollaborationOptions {
  /** Peer-unique Yrs client ID. Generated securely when omitted. */
  clientId?: number;
  /** Shared Yrs state applied before the replica is exposed. */
  initialUpdate?: Uint8Array;
  /** Receive the editor-owned collaboration replica. */
  onReplica?: (replica: CollaborationReplica | null) => void;
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
const DEFAULT_XLSX_TOOLBAR_HEIGHT = 46;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

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
    minHeight: 36,
    margin: '0 8px',
    padding: '4px 8px',
    borderRadius: 999,
    background: '#f1f5f9',
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
  formulaGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: '1 1 320px',
    minWidth: 240,
    padding: '0 6px',
    borderRight: '1px solid rgba(226, 232, 240, 0.9)',
  },
  nameBox: {
    appearance: 'none',
    width: 64,
    height: 28,
    flex: '0 0 auto',
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    background: '#f8fafc',
    color: '#0f172a',
    font: '600 12px ui-monospace, SFMono-Regular, Menlo, monospace',
    textAlign: 'center',
    outlineColor: '#2563eb',
  },
  formulaMark: {
    display: 'grid',
    placeItems: 'center',
    width: 20,
    height: 28,
    flex: '0 0 auto',
    color: '#64748b',
    font: 'italic 700 12px Georgia, serif',
    userSelect: 'none',
  },
  formulaInput: {
    appearance: 'none',
    flex: '1 1 260px',
    minWidth: 140,
    height: 28,
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    padding: '0 8px',
    background: '#ffffff',
    color: '#0f172a',
    font: '13px ui-sans-serif, system-ui, sans-serif',
    outlineColor: '#2563eb',
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

function xlsxToolbarButton(enabled: boolean, active = false): React.CSSProperties {
  return {
    appearance: 'none',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 2,
    width: 28,
    height: 28,
    padding: 0,
    border: 0,
    borderRadius: 6,
    background: active ? '#0f172a' : 'transparent',
    color: enabled ? (active ? '#ffffff' : '#64748b') : '#94a3b8',
    cursor: enabled ? 'pointer' : 'default',
    opacity: enabled ? 1 : 0.32,
  };
}

type XlsxToolbarIconName = 'save' | 'image' | 'undo' | 'redo' | 'proposals';

function XlsxToolbarIcon({ name }: { name: XlsxToolbarIconName }) {
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
      {name === 'save' && <path d="M12 3v11m0 0 4-4m-4 4-4-4M5 19h14" />}
      {name === 'image' && (
        <>
          <rect x="3" y="5" width="18" height="14" rx="2" />
          <circle cx="8.5" cy="9.5" r="1.25" />
          <path d="m6 17 4.5-4.5 3.25 3.25L16 13.5l2 2" />
        </>
      )}
      {name === 'undo' && <path d="m9 7-5 5 5 5M5 12h9a6 6 0 0 1 6 6" />}
      {name === 'redo' && <path d="m15 7 5 5-5 5m4-5h-9a6 6 0 0 0-6 6" />}
      {name === 'proposals' && (
        <>
          <path d="m12 3 1.25 3.75L17 8l-3.75 1.25L12 13l-1.25-3.75L7 8l3.75-1.25L12 3Z" />
          <path d="m18 14 .75 2.25L21 17l-2.25.75L18 20l-.75-2.25L15 17l2.25-.75L18 14Z" />
        </>
      )}
    </svg>
  );
}

/**
 * The xlsx editor React component.
 */
export function XlsxEditor({
  i18n,
  ...props
}: XlsxEditorProps) {
  return (
    <LocaleProvider i18n={i18n}>
      <XlsxEditorContent {...props} />
    </LocaleProvider>
  );
}

function XlsxEditorContent({
  file,
  fileName,
  onSave,
  collaboration,
  onReady,
  className,
}: Omit<XlsxEditorProps, 'i18n'>) {
  const { t } = useTranslation();
  const collaborationEnabled = collaboration !== undefined;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const collaborationOnReplica = collaboration?.onReplica;
  const toolbarRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const handleRef = useRef<WorkbookHandle | null>(null);
  const frameRef = useRef<DisplayList | null>(null);
  const rafRef = useRef<number | null>(null);
  const editorInputRef = useRef<HTMLInputElement>(null);
  const draggingRef = useRef(false);
  const suppressBlurRef = useRef(false);
  // latest onReady, read (not depended on) by the open effect so a changing
  // callback identity never reopens the workbook.
  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;

  const [sheetInfo, setSheetInfo] = useState<SheetInfo | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [frame, setFrame] = useState<DisplayList | null>(null);
  const [selection, setSelection] = useState<Selection | null>(null);
  const [editing, setEditing] = useState<EditState | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [toolbarHeight, setToolbarHeight] = useState(DEFAULT_XLSX_TOOLBAR_HEIGHT);
  const [revision, setRevision] = useState(0);
  const [proposals, setProposals] = useState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen] = useState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);
  // a1 lists keyed by proposal id: cells that drifted since a proposal was
  // staged, surfaced when accepting it throws a StaleProposalError.
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});

  const activeSheet = sheetInfo?.activeSheet ?? 0;

  // whether the embedded core was built with png export (raster cargo feature).
  // stable for the module's lifetime, so the export control can pre-disable.
  const pngExportAvailable = useMemo(() => isPngExportAvailable(), []);

  // whether the embedded core exposes the proposals api; gates all proposal
  // chrome so the editor degrades cleanly against an older module.
  const proposalsAvailable = useMemo(() => isProposalsAvailable(), []);

  useEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    const updateHeight = () => setToolbarHeight(toolbar.offsetHeight);
    updateHeight();
    const observer = new ResizeObserver(updateHeight);
    observer.observe(toolbar);
    return () => observer.disconnect();
  }, []);

  // re-read the pending proposal list from the handle. safe to call against an
  // old core (the loader returns an empty list).
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
  }, []);

  // open the workbook when the file changes; dispose it on change/unmount and
  // reset all editing state so a dropped file starts clean.
  useEffect(() => {
    setEditing(null);
    setFormulaDraft(null);
    setProposals([]);
    setStaleFor({});
    setProposalsPanelOpen(false);
    setCollaborationReplica(null);
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
          setSheetInfo(handle.sheetInfo());
          setSelection(selectionAt({ row: 0, col: 0 }));
          setCollaborationReplica(handle);
          setError(null);
          refreshProposals();
          const cleanup = onReadyRef.current?.({ handle, refreshProposals });
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
    refreshProposals,
  ]);

  useEffect(() => {
    if (!collaborationOnReplica || !collaborationReplica) return;
    collaborationOnReplica(collaborationReplica);
    return () => collaborationOnReplica(null);
  }, [collaborationOnReplica, collaborationReplica]);

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
    const handle = handleRef.current;
    let dl: DisplayList;
    if (handle) {
      try {
        dl = handle.displayList({ x: scroll.scrollLeft, y: scroll.scrollTop, width: w, height: h });
      } catch {
        return;
      }
    } else {
      dl = buildDemoDisplayList(w, h, t('editor.demoCellText'));
    }
    paintDisplayList(ctx, dl, dpr);
    frameRef.current = dl;
    setFrame(dl);
  }, [t]);

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

  // fold a mutation result back into state and queue a repaint.
  const applyResult = useCallback((result: EditResult) => {
    setSheetInfo(result.sheetInfo);
    setRevision((r) => r + 1);
  }, []);

  const limits = useCallback((): SelectionLimits => {
    return deriveLimits(frameRef.current, sheetInfo!, scrollRef.current?.clientHeight ?? 0);
  }, [sheetInfo]);

  // map a viewport-local pointer event to a sheet cell via the frame geometry.
  const pointToCell = useCallback((clientX: number, clientY: number): CellAddr | null => {
    const canvas = canvasRef.current;
    const grid = frameRef.current?.grid;
    if (!canvas || !grid) return null;
    const rect = canvas.getBoundingClientRect();
    return cellAtPoint(grid, clientX - rect.left, clientY - rect.top);
  }, []);

  const openEditor = useCallback(
    (seed?: string) => {
      const handle = handleRef.current;
      if (!handle || !selection) return;
      const { row, col } = selection.focus;
      let value = seed ?? '';
      if (seed === undefined) {
        try {
          value = handle.cell(activeSheet, row, col).input;
        } catch {
          value = '';
        }
      }
      setEditing({ row, col, value });
    },
    [selection, activeSheet]
  );

  // commit the open editor, optionally stepping the selection like excel.
  const commitEditor = useCallback(
    (move?: Direction) => {
      const handle = handleRef.current;
      if (!handle || !editing) return;
      suppressBlurRef.current = true;
      const { row, col, value } = editing;
      try {
        applyResult(handle.editCell(activeSheet, row, col, value));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
      setEditing(null);
      const base = selectionAt({ row, col });
      setSelection(move ? moveFocus(base, move, { limits: limits() }) : base);
      focusContainer();
    },
    [editing, activeSheet, applyResult, limits, focusContainer]
  );

  const cancelEditor = useCallback(() => {
    suppressBlurRef.current = true;
    setEditing(null);
    focusContainer();
  }, [focusContainer]);

  const clearCells = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    const r = normalizeRange(selection);
    const edits: CellInputEdit[] = [];
    for (let row = r.top; row <= r.bottom; row++) {
      for (let col = r.left; col <= r.right; col++) edits.push({ row, col, input: '' });
    }
    try {
      applyResult(handle.editCells(activeSheet, edits));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [selection, activeSheet, applyResult]);

  const copySelection = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    const r = normalizeRange(selection);
    try {
      const from = handle.cell(activeSheet, r.top, r.left).a1;
      const to = handle.cell(activeSheet, r.bottom, r.right).a1;
      const cells = handle.rangeCells(activeSheet, `${from}:${to}`);
      const tsv = toTsv(
        cells.map((row) => row.map((c) => ({ input: c.input, isFormula: c.isFormula })))
      );
      await navigator.clipboard.writeText(tsv);
    } catch {
      // clipboard denied or read failed — nothing to paste, leave state as-is.
    }
  }, [selection, activeSheet]);

  const cutSelection = useCallback(async () => {
    await copySelection();
    clearCells();
  }, [copySelection, clearCells]);

  const pasteSelection = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return;
    }
    const grid = fromTsv(text);
    if (grid.length === 0) return;
    const r = normalizeRange(selection);
    const edits: CellInputEdit[] = [];
    let width = 1;
    grid.forEach((rowArr, dr) => {
      width = Math.max(width, rowArr.length);
      rowArr.forEach((input, dc) => edits.push({ row: r.top + dr, col: r.left + dc, input }));
    });
    try {
      applyResult(handle.editCells(activeSheet, edits));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      return;
    }
    setSelection({
      anchor: { row: r.top, col: r.left },
      focus: { row: r.top + grid.length - 1, col: r.left + width - 1 },
    });
  }, [selection, activeSheet, applyResult]);

  const undo = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || collaborationEnabled) return;
    try {
      applyResult(handle.undo());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [applyResult, collaborationEnabled]);

  const redo = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || collaborationEnabled) return;
    try {
      applyResult(handle.redo());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [applyResult, collaborationEnabled]);

  // accept a proposal (optionally forcing past drift): apply it, drop any stale
  // warning, refresh the list, repaint, and return focus to the grid.
  const acceptProposal = useCallback(
    (id: string, force?: boolean) => {
      const handle = handleRef.current;
      if (!handle) return;
      try {
        applyResult(handle.acceptProposal(id, { force }));
        setStaleFor(({ [id]: _dropped, ...rest }) => rest);
        refreshProposals();
        focusContainer();
      } catch (e) {
        if (e instanceof StaleProposalError) setStaleFor((m) => ({ ...m, [id]: e.cells }));
        else setError(e instanceof Error ? e.message : String(e));
      }
    },
    [applyResult, refreshProposals, focusContainer]
  );

  // reject a proposal: drop it and its warning, then refresh so its decorations
  // disappear (they are driven by the pending list, not the canvas).
  const rejectProposal = useCallback(
    (id: string) => {
      const handle = handleRef.current;
      if (!handle) return;
      try {
        handle.rejectProposal(id);
        setStaleFor(({ [id]: _dropped, ...rest }) => rest);
        refreshProposals();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [refreshProposals]
  );

  const save = useCallback(() => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      const bytes = handle.save();
      if (onSave) onSave(bytes);
      else downloadBytes(bytes, fileName ?? 'workbook.xlsx', XLSX_MIME);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [onSave, fileName]);

  // render the current scroll window to png via the raster backend and download
  // it — the same display list the canvas paints, rasterized in the core.
  const exportPng = useCallback(() => {
    const handle = handleRef.current;
    const scroll = scrollRef.current;
    if (!handle || !scroll) return;
    try {
      const png = handle.renderPng({
        x: scroll.scrollLeft,
        y: scroll.scrollTop,
        width: scroll.clientWidth,
        height: scroll.clientHeight,
      });
      downloadBytes(png, pngName(fileName), 'image/png');
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [fileName]);

  // commit the formula bar draft to the focused cell.
  const commitFormula = useCallback(
    (move?: Direction) => {
      const handle = handleRef.current;
      if (!handle || !selection || formulaDraft == null) return;
      const { row, col } = selection.focus;
      try {
        applyResult(handle.editCell(activeSheet, row, col, formulaDraft));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
      setFormulaDraft(null);
      if (move) setSelection((prev) => (prev ? moveFocus(prev, move, { limits: limits() }) : prev));
    },
    [selection, formulaDraft, activeSheet, applyResult, limits]
  );

  // grid-level keyboard: chrome shortcuts first, then the pure selection reducer.
  const onKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      const handle = handleRef.current;
      if (!handle || !selection || !sheetInfo || editing) return;
      const mod = e.metaKey || e.ctrlKey;
      const lower = e.key.toLowerCase();

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
          void cutSelection();
          e.preventDefault();
          return;
        }
        if (lower === 'z') {
          e.shiftKey ? redo() : undo();
          e.preventDefault();
          return;
        }
        if (lower === 'y') {
          redo();
          e.preventDefault();
          return;
        }
        if (lower === 's') {
          save();
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
      undo,
      redo,
      save,
      openEditor,
      clearCells,
    ]
  );

  const onMouseDown = useCallback(
    (e: React.MouseEvent) => {
      if (!selection) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) return;
      if (e.shiftKey) setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
      else setSelection(selectionAt(addr));
      draggingRef.current = true;
      focusContainer();
    },
    [selection, pointToCell, limits, focusContainer]
  );

  const onMouseMove = useCallback(
    (e: React.MouseEvent) => {
      if (!draggingRef.current) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) return;
      setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
    },
    [pointToCell, limits]
  );

  const onDoubleClick = useCallback(() => {
    if (selection) openEditor();
  }, [selection, openEditor]);

  const grid = frame?.grid;
  const selRect = grid && selection ? rangeRect(grid, normalizeRange(selection)) : null;
  const focusRect =
    grid && selection ? cellRect(grid, selection.focus.row, selection.focus.col) : null;
  const editRect = grid && editing ? cellRect(grid, editing.row, editing.col) : null;

  const spacerWidth = sheetInfo ? sheetInfo.contentWidth : undefined;
  const spacerHeight = sheetInfo ? sheetInfo.contentHeight : undefined;
  const formulaValue = formulaDraft ?? focusedCell?.input ?? '';

  // switch sheets: retarget the core, reset scroll + selection, reread info.
  const switchSheet = (index: number) => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      handle.setActiveSheet(index);
      const scroll = scrollRef.current;
      if (scroll) {
        scroll.scrollLeft = 0;
        scroll.scrollTop = 0;
      }
      setSelection(selectionAt({ row: 0, col: 0 }));
      setEditing(null);
      setFormulaDraft(null);
      setSheetInfo(handle.sheetInfo());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div
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
      <div
        ref={toolbarRef}
        data-testid="xlsx-toolbar"
        style={xlsxToolbarStyles.shell}
      >
        <div
          style={xlsxToolbarStyles.rail}
          role="group"
          aria-label={t('toolbar.actionsLabel')}
        >
          <div
            style={{ ...xlsxToolbarStyles.group, paddingLeft: 0 }}
            role="group"
            aria-label={t('toolbar.fileActionsLabel')}
          >
            <button
              type="button"
              data-testid="xlsx-save"
              onClick={save}
              disabled={!sheetInfo}
              aria-label={t('toolbar.save')}
              title={t('toolbar.save')}
              style={xlsxToolbarButton(Boolean(sheetInfo))}
            >
              <XlsxToolbarIcon name="save" />
            </button>
            <button
              type="button"
              data-testid="xlsx-export-png"
              onClick={exportPng}
              disabled={!sheetInfo || !pngExportAvailable}
              aria-label={t('toolbar.exportPng')}
              title={t('toolbar.exportPng')}
              style={xlsxToolbarButton(Boolean(sheetInfo && pngExportAvailable))}
            >
              <XlsxToolbarIcon name="image" />
            </button>
          </div>
          <div
            style={xlsxToolbarStyles.group}
            role="group"
            aria-label={t('toolbar.historyLabel')}
          >
            <button
              type="button"
              data-testid="xlsx-undo"
              onClick={undo}
              disabled={!sheetInfo || collaborationEnabled}
              aria-label={t('toolbar.undo')}
              title={t('toolbar.undo')}
              style={xlsxToolbarButton(Boolean(sheetInfo && !collaborationEnabled))}
            >
              <XlsxToolbarIcon name="undo" />
            </button>
            <button
              type="button"
              data-testid="xlsx-redo"
              onClick={redo}
              disabled={!sheetInfo || collaborationEnabled}
              aria-label={t('toolbar.redo')}
              title={t('toolbar.redo')}
              style={xlsxToolbarButton(Boolean(sheetInfo && !collaborationEnabled))}
            >
              <XlsxToolbarIcon name="redo" />
            </button>
          </div>
          <div
            style={xlsxToolbarStyles.formulaGroup}
            role="group"
            aria-label={t('toolbar.formulaBarLabel')}
          >
            <input
              data-testid="xlsx-name-box"
              readOnly
              value={focusedCell?.a1 ?? ''}
              placeholder={t('toolbar.nameBoxPlaceholder')}
              aria-label={t('toolbar.nameBoxPlaceholder')}
              style={xlsxToolbarStyles.nameBox}
            />
            <span style={xlsxToolbarStyles.formulaMark} aria-hidden="true">
              fx
            </span>
            <input
              data-testid="xlsx-formula-input"
              value={formulaValue}
              placeholder={t('toolbar.formulaPlaceholder')}
              aria-label={t('toolbar.formulaPlaceholder')}
              disabled={!sheetInfo}
              onChange={(e) => setFormulaDraft(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  commitFormula(e.shiftKey ? 'up' : 'down');
                  focusContainer();
                  e.preventDefault();
                } else if (e.key === 'Escape') {
                  setFormulaDraft(null);
                  focusContainer();
                  e.preventDefault();
                }
              }}
              onBlur={() => commitFormula()}
              style={xlsxToolbarStyles.formulaInput}
            />
          </div>
          {proposalsAvailable && (
            <div style={xlsxToolbarStyles.proposals}>
              <button
                type="button"
                data-testid="xlsx-proposals-button"
                onClick={() => setProposalsPanelOpen((open) => !open)}
                disabled={!sheetInfo}
                aria-expanded={proposalsPanelOpen}
                aria-label={t('proposals.panelLabel')}
                title={t('proposals.panelLabel')}
                style={{
                  ...xlsxToolbarButton(Boolean(sheetInfo), proposalsPanelOpen),
                  width: proposals.length > 0 ? 42 : 28,
                }}
              >
                <XlsxToolbarIcon name="proposals" />
                {proposals.length > 0 && (
                  <span data-testid="xlsx-proposals-count" style={xlsxToolbarStyles.count}>
                    {proposals.length}
                  </span>
                )}
              </button>
            </div>
          )}
        </div>
      </div>
      {proposalsAvailable && proposalsPanelOpen && (
        <ProposalsPanel
          proposals={proposals}
          staleFor={staleFor}
          onAccept={acceptProposal}
          onReject={rejectProposal}
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
            {selRect && (
              <div
                data-testid="xlsx-selection"
                style={{
                  position: 'absolute',
                  left: selRect.x,
                  top: selRect.y,
                  width: selRect.w,
                  height: selRect.h,
                  boxSizing: 'border-box',
                  border: `1px solid ${BRAND}`,
                  background: 'rgba(33, 115, 70, 0.12)',
                }}
              />
            )}
            {focusRect && !editing && (
              <div
                style={{
                  position: 'absolute',
                  left: focusRect.x,
                  top: focusRect.y,
                  width: focusRect.w,
                  height: focusRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                }}
              />
            )}
            {/* ghost previews for pending proposals visible in this viewport.
                aria-hidden — the a11y grid announces real committed values only. */}
            {grid &&
              proposals.flatMap((proposal) =>
                proposal.cells.map((cell) => {
                  if (cell.sheet !== activeSheet) return null;
                  const rect = cellRect(grid, cell.row, cell.col);
                  if (!rect) return null;
                  return (
                    <ProposalDecoration
                      key={`${proposal.id}:${cell.a1}`}
                      rect={rect}
                      color={proposalColor(proposal.agentId)}
                      newText={cell.newText}
                      agentId={proposal.agentId}
                    />
                  );
                })
              )}
            {editing && editRect && (
              <input
                ref={editorInputRef}
                data-testid="xlsx-cell-editor"
                value={editing.value}
                onChange={(e) =>
                  setEditing((prev) => (prev ? { ...prev, value: e.target.value } : prev))
                }
                onKeyDown={(e) => {
                  e.stopPropagation();
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
                  left: editRect.x,
                  top: editRect.y,
                  width: editRect.w,
                  height: editRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                  padding: '0 3px',
                  font: '13px system-ui, sans-serif',
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
}
