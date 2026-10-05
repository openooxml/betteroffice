import {
  buildA11yGrid, cellAtPoint, cellRect, chartRegionAtPoint, createWorkbookRecoveryMutators,
  extendTo, fromTsv, hyperlinkAtCell, isProposalsAvailable, moveFocus, normalizeRange,
  paintDisplayList, parseHyperlinkLocation, rangeRect, safeExternalHyperlink, selectionAt,
  selectionKeyReducer, toTsv, WorkbookEditPeerFailedError,
} from '@betteroffice/xlsx';
import type {
  CapturedFormat, CellAddr, CellEdit, CellInputEdit, Direction, EditResult, Proposal,
  Selection, WorkbookHandle, WorkbookSheetView, XlsxEditResult,
} from '@betteroffice/xlsx';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, MouseEvent, SetStateAction, SyntheticEvent } from 'react';
import { flushSync } from 'react-dom';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { commandForEvent } from '../commands/descriptors';
import type { InputDraft, InputSeal } from '../commands/inputCoordinator';
import { inPluginChrome, pluginEventStore } from '../commands/pluginEvents';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { useWorkerXlsxCommands } from '../commands/useWorkerXlsxCommands';
import {
  createWorkerInputCoordinator, type WorkerInputCoordinatorHooks,
} from '../commands/workerInputCoordinator';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import { EditorToolbar } from '../components/EditorToolbar';
import { EditorChromeContext } from '../components/EditorToolbarContext';
import type { BorderStyle } from '../components/Toolbar';
import { FormulaBarContext, type FormulaBarBinding } from '../components/toolbar/FormulaBar';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { useTranslation } from '../i18n';
import type { XlsxPluginEditorAccess, XlsxRevealAlignment } from '../plugins/createPluginClients';
import { PluginOverlays } from '../plugins/PluginOverlays';
import { PluginDock, useDockArea, type DockPlacement } from '../plugins/PluginPanels';
import { useXlsxPluginHost } from '../plugins/useXlsxPluginHost';
import { ProposalsPanel } from '../proposals/ProposalsPanel';
import { deriveLimits, expandRangeToMergedCells, scaledRect } from '../viewer/sessionGeometry';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import {
  useEditableSessionWorkbook, type EditableSessionWorkbookProps, type EditableWorkbookSession,
} from './useEditableSessionWorkbook';
import { WorkerPaintSource, type WorkerPaintRequest, type WorkerPaintResult } from './WorkerPaintSource';

const BRAND = '#217346';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, border: 0,
  overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap',
};
const chartKeys: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1],
};

type WorkerCellDraft = InputDraft & { prefill?: boolean };

function address(cell: CellAddr): string {
  let column = '';
  for (let col = cell.col + 1; col > 0; col = Math.floor((col - 1) / 26)) {
    column = String.fromCharCode(65 + (col - 1) % 26) + column;
  }
  return `${column}${cell.row + 1}`;
}

function selectedRange(selection: Selection): string {
  const range = normalizeRange(selection);
  return `${address({ row: range.top, col: range.left })}:${address({ row: range.bottom, col: range.right })}`;
}

function fromPlugin(event: SyntheticEvent): boolean {
  return pluginEventStore(event.nativeEvent) !== null || inPluginChrome(event.target);
}

function useSyncedState<T>(initial: T) {
  const [value, setValue] = useState(initial);
  const ref = useRef(value);
  const set = useCallback((next: SetStateAction<T>) => {
    ref.current = typeof next === 'function' ? (next as (previous: T) => T)(ref.current) : next;
    setValue(ref.current);
  }, []);
  return [value, set, ref] as const;
}

function samePaint(painted: WorkerPaintResult | null, request: WorkerPaintRequest | null, sequence: number): boolean {
  if (!painted || !request || painted.sequence < sequence) return false;
  const previous = painted.request;
  return previous.generation === request.generation && previous.sheet === request.sheet &&
    previous.navigation === request.navigation && previous.zoom === request.zoom &&
    previous.dpr === request.dpr && previous.width === request.width && previous.height === request.height &&
    JSON.stringify(previous.viewport) === JSON.stringify(request.viewport);
}

function revealAxis(position: number, size: number, frozen: boolean, scroll: number, body: number,
  align: XlsxRevealAlignment): number | null {
  if (frozen) return null;
  if (align === 'start') return position;
  if (align === 'center') return Math.max(0, position - Math.max(0, body - size) / 2);
  if (position >= scroll && position + size <= scroll + body) return null;
  return position < scroll || size >= body ? position : Math.max(0, position + size - body);
}

export function XlsxWorkerEditor(props: EditableSessionWorkbookProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const editorInputRef = useRef<HTMLInputElement | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const runRef = useRef<EditableWorkbookSession | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [commands] = useState(createXlsxCommandController);
  const [active, setActive, activeRef] = useSyncedState(0);
  const [view, setView, viewRef] = useSyncedState<WorkbookSheetView | null>(null);
  const [selection, setSelection, selectionRef] = useSyncedState<Selection | null>(null);
  const [zoom, setZoom, zoomRef] = useSyncedState(1);
  const [editing, setEditing, editingRef] = useSyncedState<InputDraft | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const [previewDraft, setPreviewDraft] = useState<InputDraft | null>(null);
  const [painted, setPainted] = useState<WorkerPaintResult | null>(null);
  const [selectedChart, setSelectedChart, chartRef] = useSyncedState<{ id: string; movable: boolean } | null>(null);
  const [chartOffset, setChartOffset] = useState<{ x: number; y: number } | null>(null);
  const [capturedFormat, setCapturedFormat, capturedFormatRef] = useSyncedState<CapturedFormat | null>(null);
  const paintFormatSource = useRef<string | null>(null);
  const borderStyleRef = useRef<BorderStyle | undefined>(undefined);
  const borderColorRef = useRef<string | undefined>(undefined);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen, panelRef] = useSyncedState(false);
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});
  const mutationRef = useRef(0);
  const navigationRef = useRef(0);
  const sourceRef = useRef<WorkerPaintSource | null>(null);
  const dragging = useRef(false);
  const clickStart = useRef<CellAddr | null>(null);
  const chartDrag = useRef<{ id: string; clientX: number; clientY: number } | null>(null);
  const nudge = useRef<{ id: string; sheet: number; dx: number; dy: number } | null>(null);
  const nudgeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flushNudgeRef = useRef<() => void>(() => {});
  const suppressBlur = useRef(false);
  const suppressFormulaBlur = useRef(false);
  const composition = useRef<{ source: InputDraft['source']; done: Promise<boolean>; settle(ended: boolean): void } | null>(null);
  const paintWaiters = useRef<{ request: WorkerPaintRequest; sequence: number; resolve(value: boolean): void }[]>([]);
  const paintBoundaries = useRef(new Map<number, (error: unknown) => void>());
  const inputHooks = useRef<WorkerInputCoordinatorHooks | null>(null);
  const focus = useCallback(() => scrollRef.current?.focus({ preventScroll: true }), []);
  const fail = useCallback((error: unknown) => runRef.current?.fail(error), []);
  const reportInputError = useCallback((error: unknown) => {
    if (error instanceof XlsxCommandAdmissionError && (error.code === 'document-replaced' ||
      error.code === 'target-changed' || error.code === 'gesture-active')) return;
    fail(error);
  }, [fail]);
  const acceptInput = () => Boolean(runRef.current?.current && !runRef.current.failure && !propsRef.current.readOnly);

  const [coordinator] = useState(() => createWorkerInputCoordinator({
    generation: () => runRef.current?.generation ?? 0,
    capture: () => ({ sheet: activeRef.current, target: JSON.stringify(selectionRef.current) }),
    isReady: () => runRef.current?.ready ?? false,
    whenReady: () => runRef.current?.whenHydrated() ?? Promise.reject(new XlsxCommandAdmissionError('editor-unavailable')),
    requestHydration: (reason) => runRef.current?.requestHydration(reason),
    flushEdits: () => runRef.current?.editPeer?.flush() ?? Promise.reject(new XlsxCommandAdmissionError('editor-unavailable')),
    preview: (draft) => inputHooks.current!.preview(draft),
    resolveDraft: (draft) => inputHooks.current!.resolveDraft!(draft),
    seal: () => inputHooks.current?.seal() ?? {},
    sync: () => inputHooks.current?.sync(),
    write: (draft) => inputHooks.current?.write(draft) ?? false,
    onError: fail,
  }));
  const apiBridgeRef = useRef<WorkerEditorApiBridge | null>(null);
  const [apiBridge] = useState<WorkerEditorApiBridge>(() => ({
    coordinator: () => coordinator,
    readOnly: () => propsRef.current.readOnly ?? false,
    clearSelection: () => apiBridgeRef.current?.clearSelection(),
    focus,
    refreshProposals: () => apiBridgeRef.current?.refreshProposals(),
    recoverInput: () => apiBridgeRef.current!.recoverInput(),
    selectCells: (...args) => apiBridgeRef.current?.selectCells(...args) ?? false,
    selectCellsAsync: (...args) => apiBridgeRef.current?.selectCellsAsync(...args) ?? Promise.resolve(false),
    apply: (result) => apiBridgeRef.current?.apply(result),
  }));
  const { run, error, loading } = useEditableSessionWorkbook(props, commands.store, apiBridge);
  runRef.current = run;

  const capture = useCallback((): WorkerPaintRequest | null => {
    const owner = runRef.current;
    const scroll = scrollRef.current;
    const zoom = zoomRef.current;
    if (!owner?.current || !viewRef.current || !scroll || !canvasRef.current ||
      !scroll.clientWidth || !scroll.clientHeight) return null;
    return {
      generation: owner.generation, sheet: activeRef.current, navigation: navigationRef.current,
      zoom, dpr: window.devicePixelRatio || 1, width: scroll.clientWidth, height: scroll.clientHeight,
      viewport: { x: scroll.scrollLeft / zoom, y: scroll.scrollTop / zoom,
        width: scroll.clientWidth / zoom, height: scroll.clientHeight / zoom },
    };
  }, []);

  const boundary = useCallback((): Promise<void> => {
    const owner = runRef.current;
    if (!owner?.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (owner.recovering) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const next = (finish: boolean) => {
        const id = requestAnimationFrame(() => {
          paintBoundaries.current.delete(id);
          if (!owner.current) { reject(new XlsxCommandAdmissionError('document-replaced')); return; }
          if (owner.failure) { reject(owner.failure); return; }
          if (finish) resolve();
          else next(true);
        });
        paintBoundaries.current.set(id, reject);
      };
      next(false);
    });
  }, []);

  const preview = useCallback((draft?: InputDraft): Promise<void> => {
    if (draft && !runRef.current?.recovering) flushSync(() => setPreviewDraft(draft));
    return boundary();
  }, [boundary]);

  const refreshProposals = useCallback(() => {
    const owner = runRef.current;
    if (!owner?.peer) { setProposals([]); return; }
    setProposals(owner.peer.listProposals());
    if (!owner.recovering) sourceRef.current?.schedule();
    commands.refresh();
  }, [commands]);

  const peerPaint = useCallback(() => {
    const owner = runRef.current;
    const request = capture();
    if (!owner?.peer || owner.recovering || !request || !sourceRef.current) return;
    try {
      if (owner.peer.sheetInfo().activeSheet !== request.sheet) return;
      const displayList = owner.peer.displayList(request.viewport);
      const grid = displayList.grid;
      const rows = (grid?.rowOffsets.length ?? 0) - 1;
      const cols = (grid?.colOffsets.length ?? 0) - 1;
      const mergedRanges = grid && rows > 0 && cols > 0 ? owner.peer.mergedRanges(request.sheet,
        `${address({ row: grid.startRow, col: grid.startCol })}:${address({ row: grid.startRow + rows - 1, col: grid.startCol + cols - 1 })}`
      ).slice(0, 1024) : [];
      sourceRef.current.commit(request, { displayList, mergedRanges, version: owner.peer.version() });
      sourceRef.current.schedule();
    } catch (error) { owner.fail(error); }
  }, [capture]);

  const closeWrittenDraft = (draft: InputDraft) => {
    const live = coordinator.draft;
    if (live && (live.source !== draft.source || live.sheet !== draft.sheet || live.row !== draft.row ||
      live.col !== draft.col || live.value !== draft.value && !(live as WorkerCellDraft).prefill)) return;
    if (draft.source === 'cell') {
      const current = editingRef.current;
      if (current && current.sheet === draft.sheet && current.row === draft.row && current.col === draft.col &&
        (current.value === draft.value || (current as WorkerCellDraft).prefill)) {
        suppressBlur.current = true;
        setEditing(null);
      }
    } else setFormulaDraft(null);
  };

  const apply = useCallback((result: EditResult | XlsxEditResult) => {
    const owner = runRef.current;
    if (!owner?.peer) return;
    try {
      mutationRef.current += 1;
      const info = owner.peer.sheetInfo();
      if (viewRef.current && info.activeSheet === activeRef.current) {
        setView({ ...viewRef.current, ...info, sheet: info.activeSheet, version: owner.peer.version() });
      }
      peerPaint();
      refreshProposals();
      if ((!('ok' in result) || result.ok) && result.applied) propsRef.current.onChange?.();
      commands.refresh();
    } catch (error) {
      owner.fail(error);
    }
  }, [commands, peerPaint, refreshProposals]);

  const afterPaint = useCallback((): Promise<boolean> => {
    const owner = runRef.current;
    const request = capture();
    const source = sourceRef.current;
    if (!owner?.current || owner.failure || !request || !source) return Promise.resolve(false);
    const sequence = owner.editPeer?.sentSequence ?? 0;
    if (samePaint(source.painted, request, sequence)) return Promise.resolve(true);
    source.schedule();
    return new Promise((resolve) => {
      const waiter = { request, sequence, resolve(value: boolean) { clearTimeout(timer); resolve(value); } };
      const timer = setTimeout(() => {
        paintWaiters.current = paintWaiters.current.filter((entry) => entry !== waiter);
        resolve(false);
      }, 1000);
      paintWaiters.current.push(waiter);
    });
  }, [capture]);

  const place = useCallback((sheet: number, next: Selection | null): boolean => {
    const owner = runRef.current;
    if (!owner?.peer || !owner.current) return false;
    const info = owner.peer.sheetInfo();
    if (info.activeSheet !== sheet) return false;
    navigationRef.current += 1;
    setActive(sheet);
    setView({ ...viewRef.current!, ...info, sheet, version: owner.peer.version(), frozenWidth: 0, frozenHeight: 0 });
    setSelection(next ? structuredClone(next) : null);
    setSelectedChart(null);
    setCapturedFormat(null);
    paintFormatSource.current = null;
    if (!coordinator.draft) {
      setEditing(null);
      setFormulaDraft(null);
    }
    const scroll = scrollRef.current;
    if (scroll && next) {
      const at = owner.peer.cellPosition(sheet, next.focus.row, next.focus.col);
      scroll.scrollLeft = at.x * zoomRef.current;
      scroll.scrollTop = at.y * zoomRef.current;
    }
    if (!owner.recovering) { peerPaint(); sourceRef.current?.schedule(); }
    commands.refresh();
    return true;
  }, [commands, coordinator, peerPaint]);

  const recoverInput = useCallback(async () => {
    const owner = runRef.current;
    if (!owner?.current || !owner.peer || !owner.editPeer) throw new XlsxCommandAdmissionError('editor-unavailable');
    const edits = owner.editPeer;
    let recoveryMutators: ReturnType<typeof createWorkbookRecoveryMutators>;
    try { recoveryMutators = createWorkbookRecoveryMutators(edits); }
    catch (error) {
      if (error instanceof WorkbookEditPeerFailedError) throw error;
      throw new WorkbookEditPeerFailedError(error instanceof Error ? error : new Error(String(error)));
    }
    owner.recovering = true;
    owner.editPeer = new Proxy(edits, {
      get(target, key) {
        if (key in recoveryMutators) return Reflect.get(recoveryMutators, key);
        if (key === 'flush' || key === 'save') return () => Promise.reject(owner.failure ?? new Error('Worker is unavailable'));
        return Reflect.get(target, key);
      },
    });
    try {
      if (coordinator.draft) coordinator.submit(coordinator.draft);
      await coordinator.recover();
    }
    catch (error) {
      if (error instanceof XlsxCommandAdmissionError || error instanceof WorkbookEditPeerFailedError) throw error;
      throw new WorkbookEditPeerFailedError(error instanceof Error ? error : new Error(String(error)));
    }
    finally {
      owner.editPeer = edits;
      owner.recovering = false;
      coordinator.fail(owner.failure ?? edits.error ?? new Error('Worker is unavailable'));
    }
  }, [coordinator]);

  apiBridgeRef.current = {
    coordinator: () => coordinator, readOnly: () => propsRef.current.readOnly ?? false, focus,
    clearSelection: () => {
      coordinator.settle(); setEditing(null); setFormulaDraft(null); setSelection(null); setSelectedChart(null);
      setCapturedFormat(null); commands.refresh();
    },
    refreshProposals, recoverInput, selectCells: place,
    selectCellsAsync: async (sheet, next) => place(sheet, next), apply,
  };

  const pluginAccessRef = useRef<XlsxPluginEditorAccess | null>(null);
  const [pluginAccess] = useState<XlsxPluginEditorAccess>(() => ({
    handle: () => runRef.current?.ready && !runRef.current.failure ? runRef.current.peer : null,
    admit: (handle, operation) => pluginAccessRef.current!.admit(handle, operation),
    readOnlyRefusal: (handle) => pluginAccessRef.current!.readOnlyRefusal(handle),
    applyEdits: (...args) => pluginAccessRef.current!.applyEdits(...args),
    commands: () => commands,
    navigator: {
      selectCells: (...args) => pluginAccessRef.current!.navigator.selectCells(...args),
      scrollToCell: (...args) => pluginAccessRef.current!.navigator.scrollToCell(...args),
    },
  }));
  const pluginSelection = useMemo(() => {
    const sheetId = run?.session.state.sheets[active]?.id;
    return sheetId === undefined ? null : { sheetId, sheetIndex: active,
      cells: selection, chartId: selectedChart?.id ?? null };
  }, [run, active, selection, selectedChart]);
  const pluginHost = useXlsxPluginHost({ ...props, access: pluginAccess, commands,
    readOnly: props.readOnly ?? false, handle: run?.ready && !run.failure ? run.peer : null,
    selection: pluginSelection, canvasRef, t });
  const presentGridRef = useRef(pluginHost.presentGrid);
  presentGridRef.current = pluginHost.presentGrid;
  const [dockAreaRef, dockArea] = useDockArea(pluginHost.managed);

  const cancelPendingFrames = useCallback(() => {
    for (const [id, reject] of paintBoundaries.current) {
      cancelAnimationFrame(id);
      reject(new XlsxCommandAdmissionError('document-replaced'));
    }
    paintBoundaries.current.clear();
    for (const waiter of paintWaiters.current.splice(0)) waiter.resolve(false);
  }, []);

  useLayoutEffect(() => {
    setPainted(null); setPreviewDraft(null); setSelection(null); setView(null); setEditing(null);
    setFormulaDraft(null); setSelectedChart(null); setChartOffset(null); setFocusedCell(null);
    setCapturedFormat(null); setProposals([]); setStaleFor({}); setProposalsPanelOpen(false);
    mutationRef.current = 0;
    navigationRef.current += 1;
    dragging.current = false;
    chartDrag.current = null;
    nudge.current = null;
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    composition.current?.settle(false);
    composition.current = null;
    pluginHost.beginLoad();
    if (!run) return;
    const source = new WorkerPaintSource({
      generation: run.generation, capture, isCurrent: () => run.current && !run.failure,
      requestFrame: (request) => run.session.call.frame(request.viewport, { sheet: request.sheet }),
      sentSequence: () => run.editPeer?.sentSequence ?? 0,
      publish(painted) {
        const canvas = canvasRef.current;
        const context = canvas?.getContext('2d');
        if (!canvas || !context) throw new Error('Canvas context is unavailable');
        const request = painted.request;
        canvas.width = Math.round(request.width * request.dpr);
        canvas.height = Math.round(request.height * request.dpr);
        canvas.style.width = `${request.width}px`;
        canvas.style.height = `${request.height}px`;
        paintDisplayList(context, painted.displayList, request.dpr * request.zoom);
        setPainted(painted);
        if (painted.source === 'peer') setPreviewDraft(null);
        const sheetId = run.session.state.sheets[request.sheet]?.id;
        presentGridRef.current(sheetId === undefined ? null : {
          frame: painted.displayList, zoom: request.zoom, viewport: request.viewport, version: painted.version, sheetId,
        });
        for (const waiter of paintWaiters.current.slice()) {
          const covered = samePaint(painted, waiter.request, waiter.sequence);
          if (covered || !samePaint(painted, waiter.request, 0)) {
            paintWaiters.current = paintWaiters.current.filter((entry) => entry !== waiter);
            waiter.resolve(covered);
          }
        }
        run.firstPaint();
      },
      onError: (error) => run.fail(error),
    });
    sourceRef.current = source;
    const sheet = run.session.state.activeSheet;
    setActive(sheet);
    void run.session.call.sheetView(sheet).then((next) => {
      if (!run.current || run.failure || sourceRef.current !== source) return;
      setView(next);
      setSelection(selectionAt({ row: 0, col: 0 }));
      if (scrollRef.current) {
        scrollRef.current.scrollLeft = next.initialScrollX * zoomRef.current;
        scrollRef.current.scrollTop = next.initialScrollY * zoomRef.current;
      }
      source.schedule();
    }).catch((error) => run.fail(error));
    return () => {
      source.dispose();
      if (sourceRef.current === source) sourceRef.current = null;
      cancelPendingFrames();
      presentGridRef.current(null);
    };
  }, [run, capture, cancelPendingFrames]);

  useEffect(() => {
    const scroll = scrollRef.current;
    if (!scroll || !run) return;
    const schedule = () => sourceRef.current?.schedule();
    scroll.addEventListener('scroll', schedule, { passive: true });
    const observer = new ResizeObserver(schedule);
    observer.observe(scroll);
    schedule();
    return () => { scroll.removeEventListener('scroll', schedule); observer.disconnect(); };
  }, [run, zoom]);

  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList | undefined;
    const update = () => {
      sourceRef.current?.schedule();
      media?.removeEventListener('change', update);
      media = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      media.addEventListener('change', update);
    };
    update();
    return () => media?.removeEventListener('change', update);
  }, []);

  useEffect(() => {
    if (!run?.failure) return;
    sourceRef.current?.dispose();
    cancelPendingFrames();
    chartDrag.current = null;
    nudge.current = null;
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    composition.current?.settle(false);
    composition.current = null;
  }, [run, run?.failure, cancelPendingFrames]);

  useEffect(() => () => {
    cancelPendingFrames();
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    composition.current?.settle(false);
    composition.current = null;
  }, [cancelPendingFrames]);

  useEffect(() => {
    setFocusedCell(null);
    if (!run?.current || run.failure || !selection || selectedChart) return;
    if (run.ready && run.peer) { setFocusedCell(run.peer.cell(active, selection.focus.row, selection.focus.col)); return; }
    let cancelled = false;
    const navigation = navigationRef.current;
    void run.session.call.cellInputs(active, address(selection.focus)).then((result) => {
      if (!cancelled && run.current && navigationRef.current === navigation) setFocusedCell(result.cells[0]?.[0] ?? null);
    }).catch((error) => { if (!cancelled && run.current) run.fail(error); });
    return () => { cancelled = true; };
  }, [run, run?.peer, selection, active, selectedChart, painted]);

  const limits = () => deriveLimits(sourceRef.current?.painted?.displayList ?? null, viewRef.current!,
    (scrollRef.current?.clientHeight ?? 0) / zoomRef.current);
  const draftFor = (source: InputDraft['source'], cell: CellAddr, value: string): InputDraft => ({
    generation: runRef.current!.generation, sheet: activeRef.current, ...cell, source, value,
  });
  const syncDraft = () => {
    const draft = coordinator.draft;
    const input = draft?.source === 'cell' ? editorInputRef.current : formulaInputRef.current;
    if (draft && input && draft.value !== input.value) {
      coordinator.setDraft({ ...draft, value: input.value, prefill: false } as WorkerCellDraft);
    }
  };
  const seal = (): InputSeal => {
    if (chartDrag.current) return { refused: 'gesture-active' };
    flushNudgeRef.current();
    const pending = composition.current;
    if (!pending) return {};
    if (pending.source === 'cell') {
      suppressBlur.current = true; editorInputRef.current?.blur(); suppressBlur.current = false;
    } else {
      suppressFormulaBlur.current = true; formulaInputRef.current?.blur(); suppressFormulaBlur.current = false;
    }
    return { composition: pending.done };
  };
  inputHooks.current = {
    generation: () => run?.generation ?? 0,
    capture: () => ({ sheet: active, target: JSON.stringify(selection) }),
    isReady: () => run?.ready ?? false, whenReady: () => run!.whenHydrated(),
    requestHydration: (reason) => run!.requestHydration(reason), flushEdits: () => run!.editPeer!.flush(),
    preview, seal, sync: syncDraft,
    async resolveDraft(draft) {
      if (!(draft as WorkerCellDraft).prefill) return draft;
      const peer = runRef.current?.peer;
      if (!peer) throw new XlsxCommandAdmissionError('editor-unavailable');
      return { ...draft, value: peer.cell(draft.sheet, draft.row, draft.col).input };
    },
    write(draft) {
      const owner = runRef.current;
      if (!owner?.ready || !owner.peer || !owner.editPeer) return false;
      if (propsRef.current.readOnly) {
        if (owner.recovering) throw new XlsxCommandAdmissionError('input-failed');
        return true;
      }
      if (owner.peer.cell(draft.sheet, draft.row, draft.col).input !== draft.value) {
        const result = owner.editPeer.editCell(draft.sheet, draft.row, draft.col, draft.value);
        apply(result);
      } else peerPaint();
      closeWrittenDraft(draft);
      return true;
    },
  };

  const openEditor = (seed?: string) => {
    const owner = runRef.current;
    const selected = selectionRef.current;
    if (!acceptInput() || !owner || !selected) return;
    const cell = selected.focus;
    const draft: WorkerCellDraft = {
      ...draftFor('cell', cell, seed ?? owner.peer?.cell(activeRef.current, cell.row, cell.col).input ?? ''),
      prefill: seed === undefined && !owner.peer,
    };
    coordinator.setDraft(draft);
    setEditing(draft);
    if (seed !== undefined || owner.peer) return;
    void owner.session.call.cellInputs(draft.sheet, address(cell)).then((result) => {
      if (!owner.current || owner.failure || coordinator.draft !== draft || editingRef.current !== draft) return;
      const next = { ...draft, value: result.cells[0]?.[0]?.input ?? '', prefill: false };
      const input = editorInputRef.current;
      const start = input?.selectionStart;
      const end = input?.selectionEnd;
      flushSync(() => { setEditing(next); coordinator.setDraft(next); });
      if (input && start != null && end != null) input.setSelectionRange(start, end);
    }).catch((error) => owner.fail(error));
  };

  const commitDraft = (source: InputDraft['source'], move?: Direction) => {
    if (!acceptInput() || composition.current) return;
    syncDraft();
    const draft = coordinator.draft;
    if (draft?.source !== source) return;
    setPreviewDraft(draft);
    coordinator.submit(draft);
    if (source === 'cell') { suppressBlur.current = true; setEditing(null); }
    else setFormulaDraft(null);
    const base = selectionAt({ row: draft.row, col: draft.col });
    setSelection(move ? moveFocus(base, move, { limits: limits() }) : base);
    focus();
  };
  const cancelDraft = () => {
    coordinator.setDraft(null); suppressBlur.current = true; setEditing(null); setFormulaDraft(null); focus();
  };

  const clipboard = (kind: 'copy' | 'cut' | 'paste') => {
    const owner = runRef.current;
    const selected = selectionRef.current;
    if (!owner?.current || owner.failure || !selected || kind !== 'copy' && !acceptInput()) return;
    const sheet = activeRef.current;
    const range = normalizeRange(selected);
    let resolveCut: ((payload: Blob) => void) | undefined;
    let cutWrite: Promise<void> | undefined;
    const read = () => {
      if (owner.peer) return Promise.resolve(owner.peer.rangeCells(sheet, selectedRange(selected)));
      return owner.session.call.cellInputs(sheet, selectedRange(selected)).then((result) => result.cells);
    };
    const captureClipboard = () => {
      if (kind === 'paste') return navigator.clipboard.readText();
      if (kind === 'cut') {
        if (typeof ClipboardItem === 'function' && typeof navigator.clipboard.write === 'function') {
          const payload = new Promise<Blob>((resolve) => { resolveCut = resolve; });
          cutWrite = navigator.clipboard.write([new ClipboardItem({ 'text/plain': payload })]);
          void cutWrite.catch(() => {});
        }
        return '';
      }
      if (owner.peer) return navigator.clipboard.writeText(toTsv(
        owner.peer.rangeCells(sheet, selectedRange(selected))
      )).then(() => '');
      if (typeof ClipboardItem === 'function' && typeof navigator.clipboard.write === 'function') {
        const text = read().then((cells) => new Blob([toTsv(cells)], { type: 'text/plain' }));
        return navigator.clipboard.write([new ClipboardItem({ 'text/plain': text })]).then(() => '');
      }
      return read().then((cells) => navigator.clipboard.writeText(toTsv(cells)).then(() => ''));
    };
    if (kind === 'copy') { void captureClipboard().catch(() => {}); return; }
    void coordinator.clipboard(captureClipboard, async (text, _, markApplied) => {
      if (kind === 'cut') {
        const copied = toTsv(owner.peer!.rangeCells(sheet, selectedRange(selected)));
        resolveCut?.(new Blob([copied], { type: 'text/plain' }));
        if (cutWrite) await cutWrite;
        else await navigator.clipboard.writeText(copied);
      }
      const edits: CellInputEdit[] = [];
      if (kind === 'paste') {
        fromTsv(text).forEach((row, dr) => row.forEach((input, dc) =>
          edits.push({ row: range.top + dr, col: range.left + dc, input })));
      } else {
        for (let row = range.top; row <= range.bottom; row++) {
          for (let col = range.left; col <= range.right; col++) edits.push({ row, col, input: '' });
        }
      }
      if (edits.length === 0) return;
      await preview({ ...draftFor('cell', { row: edits[0].row, col: edits[0].col }, edits[0].input), sheet });
      const result = owner.editPeer!.editCells(sheet, edits);
      markApplied();
      apply(result);
      if (kind === 'paste' && selectionRef.current === selected && activeRef.current === sheet) {
        setSelection({ anchor: { row: range.top, col: range.left }, focus: {
          row: Math.max(...edits.map((edit) => edit.row)), col: Math.max(...edits.map((edit) => edit.col)),
        } });
      }
    }, { sheet, target: JSON.stringify(range) }).catch(reportInputError);
  };

  const clearCells = () => {
    if (!acceptInput() || !selectionRef.current) return;
    const sheet = activeRef.current;
    const range = normalizeRange(selectionRef.current);
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      const edits: CellInputEdit[] = [];
      for (let row = range.top; row <= range.bottom; row++) {
        for (let col = range.left; col <= range.right; col++) edits.push({ row, col, input: '' });
      }
      await preview({ ...draftFor('cell', { row: range.top, col: range.left }, ''), sheet });
      const result = runRef.current!.editPeer!.editCells(sheet, edits);
      markApplied(); apply(result);
    }, { kind: 'input', target: { sheet, target: JSON.stringify(range) } }).catch(reportInputError);
  };

  const navigate = (sheet: number, next: Selection) => {
    const owner = runRef.current;
    if (!owner?.current || owner.failure) return;
    const target = structuredClone(next);
    void coordinator.runAfterPendingInput((_, markApplied) => {
      owner.editPeer!.setActiveSheet(sheet);
      markApplied(); place(sheet, target);
    }, { kind: 'input', target: { sheet, target: JSON.stringify(target) } }).catch(reportInputError);
  };

  const moveChart = (sheet: number, id: string, dx: number, dy: number) => {
    if (!acceptInput() || dx === 0 && dy === 0) return;
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      await boundary();
      const result = runRef.current!.editPeer!.moveChart(sheet, id, dx, dy);
      markApplied(); apply(result); setChartOffset(null);
    }, { kind: 'chart', target: { sheet, target: id }, input: { dx, dy } }).catch(reportInputError);
  };
  flushNudgeRef.current = () => {
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    const pending = nudge.current;
    nudge.current = null;
    if (pending) moveChart(pending.sheet, pending.id, pending.dx, pending.dy);
  };
  const moveChartRef = useRef(moveChart);
  moveChartRef.current = moveChart;
  useEffect(() => {
    const stop = (event: globalThis.MouseEvent) => {
      dragging.current = false;
      const drag = chartDrag.current;
      chartDrag.current = null;
      const paintedZoom = sourceRef.current?.painted?.request.zoom ?? zoomRef.current;
      if (drag && event.button === 0) moveChartRef.current(activeRef.current, drag.id,
        (event.clientX - drag.clientX) / paintedZoom, (event.clientY - drag.clientY) / paintedZoom);
      else setChartOffset(null);
    };
    const cancel = () => { dragging.current = false; chartDrag.current = null; setChartOffset(null); flushNudgeRef.current(); };
    window.addEventListener('mouseup', stop);
    window.addEventListener('blur', cancel);
    return () => { window.removeEventListener('mouseup', stop); window.removeEventListener('blur', cancel); };
  }, []);

  const point = (event: MouseEvent) => {
    const painted = sourceRef.current?.painted;
    const canvas = canvasRef.current;
    if (!painted || painted.request.sheet !== activeRef.current || !canvas) return { cell: null, chart: null };
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / painted.request.zoom;
    const y = (event.clientY - rect.top) / painted.request.zoom;
    return { cell: painted.geometry ? cellAtPoint(painted.geometry, x, y) : null,
      chart: chartRegionAtPoint(painted.displayList.charts, x, y) };
  };
  const onMouseDown = (event: MouseEvent) => {
    if (fromPlugin(event) || runRef.current?.failure || editorInputRef.current?.contains(event.target as Node)) return;
    const { cell, chart } = point(event);
    const dismissing = editingRef.current !== null;
    if (dismissing) commitDraft('cell');
    else if (coordinator.draft?.source === 'formula') commitDraft('formula');
    flushNudgeRef.current();
    chartDrag.current = null;
    clickStart.current = null;
    if (chart) {
      setSelectedChart({ id: chart.id, movable: chart.movable });
      if (acceptInput() && chart.movable && event.button === 0) {
        chartDrag.current = { id: chart.id, clientX: event.clientX, clientY: event.clientY };
      }
      focus(); event.preventDefault(); return;
    }
    setSelectedChart(null);
    if (!cell || !selectionRef.current) return;
    setSelection(event.shiftKey ? extendTo(selectionRef.current, cell, limits()) : selectionAt(cell));
    clickStart.current = dismissing ? null : cell;
    dragging.current = true;
    focus();
  };
  const onMouseMove = (event: MouseEvent<HTMLDivElement>) => {
    if (fromPlugin(event) || runRef.current?.failure) return;
    const drag = chartDrag.current;
    if (drag) {
      if (!event.buttons) { chartDrag.current = null; setChartOffset(null); return; }
      setChartOffset({ x: event.clientX - drag.clientX, y: event.clientY - drag.clientY });
      event.currentTarget.style.cursor = 'move'; event.currentTarget.title = ''; return;
    }
    const { cell, chart } = point(event);
    if (dragging.current && cell && selectionRef.current) {
      setSelection(extendTo(selectionRef.current, cell, limits())); return;
    }
    const hyperlink = cell && sourceRef.current?.painted ?
      hyperlinkAtCell(sourceRef.current.painted.displayList, cell.row, cell.col) : null;
    event.currentTarget.style.cursor = chart?.movable && acceptInput() ? 'move' : chart || hyperlink ? 'pointer' : 'default';
    event.currentTarget.title = chart ? '' : hyperlink?.tooltip ?? '';
  };
  const onClick = (event: MouseEvent) => {
    if (fromPlugin(event) || runRef.current?.failure) return;
    const { cell } = point(event);
    const start = clickStart.current;
    clickStart.current = null;
    const frame = sourceRef.current?.painted?.displayList;
    if (!cell || !start || !frame || cell.row !== start.row || cell.col !== start.col) return;
    const hyperlink = hyperlinkAtCell(frame, cell.row, cell.col);
    if (!hyperlink) return;
    const href = safeExternalHyperlink(hyperlink);
    if (href) { window.open(href, '_blank', 'noopener,noreferrer'); return; }
    if (!hyperlink.location) return;
    const sheets = runRef.current!.session.state.sheets;
    const destination = parseHyperlinkLocation(hyperlink.location, sheets[activeRef.current]?.name ?? '');
    if (!destination) return;
    const sheet = sheets.findIndex((entry) => entry.name.toLowerCase() === destination.sheetName.toLowerCase());
    if (sheet >= 0) navigate(sheet, selectionAt(destination));
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (fromPlugin(event) || runRef.current?.failure || !viewRef.current || !selectionRef.current ||
      editingRef.current || commandForEvent(event)) return;
    const mod = event.metaKey || event.ctrlKey;
    const chart = chartRef.current;
    if (chart) {
      if (event.key === 'Escape') {
        chartDrag.current = null; nudge.current = null; setChartOffset(null); setSelectedChart(null);
        if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
        nudgeTimer.current = null; event.preventDefault(); return;
      }
      const delta = !mod && chartKeys[event.key];
      if (delta) {
        if (acceptInput() && chart.movable) {
          const step = event.shiftKey ? 10 : 1;
          const paintedZoom = sourceRef.current?.painted?.request.zoom ?? zoomRef.current;
          const base = nudge.current ?? { id: chart.id, sheet: activeRef.current, dx: 0, dy: 0 };
          nudge.current = { ...base, dx: base.dx + delta[0] * step, dy: base.dy + delta[1] * step };
          setChartOffset({ x: nudge.current.dx * paintedZoom, y: nudge.current.dy * paintedZoom });
          if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
          nudgeTimer.current = setTimeout(() => flushNudgeRef.current(), 250);
        }
        event.preventDefault();
      }
      return;
    }
    if (mod) {
      const key = event.key.toLowerCase();
      if (key === 'c' || key === 'x' || key === 'v') {
        clipboard(key === 'c' ? 'copy' : key === 'x' ? 'cut' : 'paste'); event.preventDefault(); return;
      }
    }
    const action = selectionKeyReducer(selectionRef.current, event, limits());
    if (action.type === 'move') setSelection(action.selection);
    else if (action.type === 'startEdit') openEditor(action.initialInput);
    else if (action.type === 'clear') clearCells();
    if (action.type !== 'none') event.preventDefault();
  };

  const reveal = (peer: WorkbookHandle, sheet: number, row: number, col: number, align: XlsxRevealAlignment,
    live: () => boolean) => {
    const owner = runRef.current;
    const scroll = scrollRef.current;
    if (!owner?.current || !live() || !scroll || activeRef.current !== sheet) return;
    const info = peer.sheetInfo();
    const at = peer.cellPosition(sheet, row, col);
    const next = peer.cellPosition(sheet, Math.min(row + 1, 1_048_575), Math.min(col + 1, 16_383));
    const zoom = zoomRef.current;
    const frozen = info.frozenRows || info.frozenCols ? peer.displayList({
      x: 0, y: 0, width: scroll.clientWidth / zoom, height: scroll.clientHeight / zoom,
    }).grid : undefined;
    const x = revealAxis(at.x, next.x - at.x, col < info.frozenCols, scroll.scrollLeft / zoom,
      scroll.clientWidth / zoom - (frozen?.colOffsets[info.frozenCols] ?? 0), align);
    const y = revealAxis(at.y, next.y - at.y, row < info.frozenRows, scroll.scrollTop / zoom,
      scroll.clientHeight / zoom - (frozen?.rowOffsets[info.frozenRows] ?? 0), align);
    if (x !== null) scroll.scrollLeft = x * zoom;
    if (y !== null) scroll.scrollTop = y * zoom;
    if (!owner.recovering) sourceRef.current?.schedule();
  };
  const validCell = (peer: WorkbookHandle, sheet: number, cell: CellAddr | undefined) => {
    if (!cell || !Number.isInteger(cell.row) || !Number.isInteger(cell.col) || cell.row < 0 || cell.col < 0 ||
      cell.row > 1_048_575 || cell.col > 16_383) return false;
    try { peer.cellPosition(sheet, cell.row, cell.col); return true; } catch { return false; }
  };
  pluginAccessRef.current = {
    handle: pluginAccess.handle, commands: () => commands,
    async admit(handle, operation) {
      try {
        const value = await coordinator.runAfterPendingInput(async (_, markApplied) => {
          if (runRef.current?.peer !== handle) throw new XlsxCommandAdmissionError('document-replaced');
          await boundary();
          const value = operation();
          if (runRef.current?.recovering && value !== null && typeof value === 'object' &&
            ('refused' in value || 'ok' in value && value.ok === false)) {
            throw new XlsxCommandAdmissionError('input-failed');
          }
          markApplied(); return value;
        }, { kind: 'plugin' });
        return { ok: true, value };
      } catch (error) {
        return { ok: false, code: error instanceof XlsxCommandAdmissionError && error.code === 'document-replaced'
          ? 'document-replaced' : 'input-failed', error };
      }
    },
    readOnlyRefusal: (handle) => propsRef.current.readOnly ? {
      ok: false, version: handle.version(), failure: { code: 'read-only', message: 'The editor is read-only' },
    } : null,
    applyEdits(handle, request, authorize, commit) {
      const denied = authorize();
      if (denied) return denied;
      const readOnly = pluginAccessRef.current!.readOnlyRefusal(handle);
      if (readOnly) return readOnly;
      const result = commit(() => runRef.current!.editPeer!.applyEdits(request));
      apply(result); return result;
    },
    navigator: {
      selectCells(peer, target, wantsFocus, live) {
        const sheet = peer.sheetInfo().sheetIds.indexOf(target?.sheetId);
        if (sheet < 0 || !validCell(peer, sheet, target?.selection?.anchor) ||
          !validCell(peer, sheet, target?.selection?.focus)) return 'missing-target';
        runRef.current!.editPeer!.setActiveSheet(sheet);
        place(sheet, target.selection);
        reveal(peer, sheet, target.selection.focus.row, target.selection.focus.col, 'nearest', live);
        if (wantsFocus) focus();
        return null;
      },
      scrollToCell(peer, target, align, live) {
        const sheet = peer.sheetInfo().sheetIds.indexOf(target?.sheetId);
        if (sheet < 0 || !validCell(peer, sheet, target)) return 'missing-target';
        if (sheet !== activeRef.current) {
          runRef.current!.editPeer!.setActiveSheet(sheet); place(sheet, null);
        }
        reveal(peer, sheet, target.row, target.col, align, live); return null;
      },
    },
  };

  useWorkerXlsxCommands(commands, {
    peer: () => runRef.current?.peer ?? null, editPeer: () => runRef.current?.editPeer ?? null,
    status: () => runRef.current?.ready ? 'ready' : runRef.current && !runRef.current.failure || loading ? 'loading' : 'empty',
    readOnly: () => propsRef.current.readOnly ?? false, collaborative: () => false,
    mutation: () => mutationRef.current, generation: () => runRef.current?.generation ?? 0,
    view: () => ({ sheet: activeRef.current, selection: selectionRef.current, chartSelected: chartRef.current !== null,
      zoom: zoomRef.current, capturedFormat: capturedFormatRef.current, borderStyle: borderStyleRef.current,
      borderColor: borderColorRef.current, proposals: proposalsRef.current, proposalsAvailable: isProposalsAvailable(),
      proposalsPanelOpen: panelRef.current, pngExport: false }),
    coordinator, preview: () => preview(), recovering: () => runRef.current?.recovering ?? false,
    i18n: () => propsRef.current.i18n, translate: t, apply, fail,
    setZoom: (scale) => { setZoom(scale); sourceRef.current?.schedule(); },
    setProposalsPanelOpen, setCapturedFormat: (format, source) => {
      setCapturedFormat(format); paintFormatSource.current = source;
    },
    setBorderStyle: (style) => { borderStyleRef.current = style; },
    setBorderColor: (color) => { borderColorRef.current = color; },
    markStale: (id, cells) => setStaleFor(({ [id]: _removed, ...rest }) => cells ? { ...rest, [id]: cells } : rest),
    refreshProposals,
    deliver(bytes) {
      if (propsRef.current.onSave) propsRef.current.onSave(bytes);
      else {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: XLSX_MIME }));
        const anchor = document.createElement('a');
        anchor.href = url; anchor.download = propsRef.current.fileName ?? 'workbook.xlsx';
        document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
      }
    },
    afterPaint, focusGrid: focus,
  });
  useCommandShortcuts(commands, rootRef);

  useEffect(() => {
    if (!capturedFormat || !selection || !acceptInput() || dragging.current) return;
    const range = normalizeRange(selection);
    const key = `${active}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
    if (key === paintFormatSource.current) return;
    const format = capturedFormat;
    const a1 = selectedRange(selection);
    setCapturedFormat(null); paintFormatSource.current = null;
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      await boundary();
      const result = runRef.current!.editPeer!.applyFormat(active, a1, format);
      markApplied(); apply(result); focus();
    }, { kind: 'input', target: { sheet: active, target: a1 } }).catch(reportInputError);
  }, [capturedFormat, selection, active, coordinator, boundary, apply, reportInputError, focus]);

  const startComposition = (source: InputDraft['source']) => {
    let settle!: (ended: boolean) => void;
    const done = new Promise<boolean>((resolve) => { settle = resolve; });
    composition.current = { source, done, settle };
  };
  const endComposition = () => {
    syncDraft(); composition.current?.settle(true); composition.current = null;
  };
  const attachCellInput = useCallback((input: HTMLInputElement | null) => {
    editorInputRef.current = input;
    if (input) suppressBlur.current = false;
    if (!input && composition.current?.source === 'cell') {
      composition.current.settle(false); composition.current = null;
    }
  }, []);
  const attachFormulaInput = useCallback((input: HTMLInputElement | null) => {
    formulaInputRef.current = input;
    if (!input && composition.current?.source === 'formula') {
      composition.current.settle(false); composition.current = null;
    }
  }, []);
  useLayoutEffect(() => {
    if (!editing || run?.failure) return;
    const input = editorInputRef.current;
    if (input && document.activeElement !== input) {
      input.focus({ preventScroll: true }); input.setSelectionRange(input.value.length, input.value.length);
    }
  }, [editing !== null]);

  const formulaBar: FormulaBarBinding = {
    a1: selection ? address(selection.focus) : '', value: formulaDraft ?? focusedCell?.input ?? '',
    disabled: !view || !selection || Boolean(error), readOnly: !acceptInput() || selectedChart !== null,
    inputRef: attachFormulaInput,
    onChange(value) {
      const selected = selectionRef.current;
      if (!acceptInput() || !selected || chartRef.current) return;
      setFormulaDraft(value); coordinator.setDraft(draftFor('formula', selected.focus, value));
    },
    onCommit: (move) => commitDraft('formula', move), onCancel: cancelDraft,
    onBlur: () => { if (!suppressFormulaBlur.current) commitDraft('formula'); },
    onCompositionStart: () => startComposition('formula'), onCompositionEnd: endComposition,
  };
  const frame = painted?.displayList;
  useEffect(() => {
    if (!selectedChart || !frame || frame.charts?.some((chart) => chart.id === selectedChart.id)) return;
    flushNudgeRef.current(); chartDrag.current = null; setChartOffset(null); setSelectedChart(null);
  }, [selectedChart, frame]);
  const paintedZoom = painted?.request.zoom ?? zoom;
  const visibleSelection = painted?.request.sheet === active ? selection : null;
  const merged = painted?.mergedRanges ?? [];
  const grid = frame?.grid;
  const selected = visibleSelection ? expandRangeToMergedCells(normalizeRange(visibleSelection), merged) : null;
  const focused = visibleSelection ? expandRangeToMergedCells({ top: visibleSelection.focus.row,
    left: visibleSelection.focus.col, bottom: visibleSelection.focus.row, right: visibleSelection.focus.col }, merged) : null;
  const selectionRect = grid && selected ? rangeRect(grid, selected) : null;
  const focusRect = grid && focused ? rangeRect(grid, focused) : null;
  const editRect = grid && editing && editing.sheet === painted?.request.sheet ? cellRect(grid, editing.row, editing.col) : null;
  const previewRect = grid && previewDraft && previewDraft.sheet === painted?.request.sheet ?
    cellRect(grid, previewDraft.row, previewDraft.col) : null;
  const chartRect = selectedChart ? frame?.charts?.find((chart) => chart.id === selectedChart.id)?.rect : null;
  const sheets = run?.session.state.sheets ?? [];
  const a11yGrid = frame && painted ? buildA11yGrid(frame, visibleSelection, sheets[painted.request.sheet]?.name ?? '', {
    gridLabel: t('a11y.gridLabel'), rowHeaderLabel: t('a11y.rowHeaderLabel'), columnHeaderLabel: t('a11y.columnHeaderLabel'),
    cellLabel: t('a11y.cellLabel'), cellLabelSelected: t('a11y.cellLabelSelected'), emptyCellLabel: t('a11y.emptyCellLabel'),
    emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
  }) : null;
  const outline = (rect: NonNullable<typeof selectionRect>, border: string, background?: string): CSSProperties => {
    const scaled = scaledRect(rect, paintedZoom);
    return { position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      boxSizing: 'border-box', border, background };
  };
  const defaultToolbar = <EditorToolbar mode="commands">
    <EditorToolbar.Toolbar />
    <div style={{ display: 'flex', padding: '4px 8px' }}>
      <ToolbarCommandButton id="save" />
      <ToolbarCommandButton id="exportPng" />
      <EditorToolbar.FormulaBar />
      {isProposalsAvailable() && <ToolbarCommandButton id="proposalsPanel" />}
    </div>
  </EditorToolbar>;
  const toolbar = props.showToolbar === false ? null : props.toolbar === undefined ? defaultToolbar : props.toolbar;
  const renderDock = (placement: DockPlacement) => pluginHost.managed ? <PluginDock host={pluginHost.host}
    placement={placement} activations={pluginHost.activations.filter((activation) => activation.plugin.panel?.placement === placement)}
    available={dockArea} /> : null;

  return <XlsxCommandContext.Provider value={commands.store}>
    <EditorChromeContext.Provider value={true}>
      <FormulaBarContext.Provider value={formulaBar}>
        <div ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
          style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%', height: '100%',
            minWidth: 0, color: '#202124', background: '#ffffff', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }}>
          {toolbar != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbar}</div>}
          <div ref={dockAreaRef} data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minWidth: 0, minHeight: 0 }}>
            {renderDock('left')}
            <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
              {isProposalsAvailable() && proposalsPanelOpen && <ProposalsPanel proposals={proposals} staleFor={staleFor}
                style={{ top: 4, right: 8, width: 'min(320px, calc(100% - 16px))', maxHeight: 'min(420px, calc(100% - 12px))' }} />}
              <div ref={scrollRef} data-testid="xlsx-scroll" tabIndex={0} onKeyDown={onKeyDown}
                onMouseDown={onMouseDown} onMouseMove={onMouseMove} onClick={onClick}
                onDoubleClick={(event) => {
                  if (fromPlugin(event) || editingRef.current || point(event).chart) return;
                  const cell = point(event).cell;
                  if (cell && frame && hyperlinkAtCell(frame, cell.row, cell.col)) return;
                  openEditor();
                }}
                onMouseLeave={(event) => { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }}
                style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}>
                <div style={{ position: 'absolute', top: 0, left: 0, width: view ? view.contentWidth * zoom : '100%',
                  height: view ? view.contentHeight * zoom : '100%' }} />
                <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
                  <canvas ref={canvasRef} style={{ display: 'block', position: 'absolute', top: 0, left: 0 }} />
                  {pluginHost.managed && <PluginOverlays host={pluginHost.host} activations={pluginHost.activations}
                    layerRef={pluginHost.overlayLayerRef} width={(frame?.width ?? 0) * paintedZoom} height={(frame?.height ?? 0) * paintedZoom} />}
                  <div data-testid="xlsx-overlay-host" style={{ position: 'absolute', top: 0, left: 0, width: 0, height: 0, pointerEvents: 'none' }}>
                    {selectionRect && !selectedChart && <div data-testid="xlsx-selection"
                      style={outline(selectionRect, `1px solid ${BRAND}`, 'rgba(33, 115, 70, 0.12)')} />}
                    {focusRect && !selectedChart && !editing && <div style={outline(focusRect, `2px solid ${BRAND}`)} />}
                    {chartRect && <div data-testid="xlsx-chart-selection" data-chart-id={selectedChart?.id} aria-hidden
                      style={{ ...outline(chartRect, `2px solid ${BRAND}`),
                        transform: `translate(${chartOffset?.x ?? 0}px, ${chartOffset?.y ?? 0}px)`, boxShadow: '0 1px 6px rgba(0, 0, 0, 0.25)' }} />}
                    {previewDraft && previewRect && <div data-testid="xlsx-commit-preview" aria-hidden
                      style={{ ...outline(previewRect, `2px solid ${BRAND}`, '#ffffff'), padding: '0 3px',
                        overflow: 'hidden', whiteSpace: 'pre', font: `${13 * paintedZoom}px system-ui, sans-serif` }}>
                      {previewDraft.value}
                    </div>}
                    {editing && editRect && <input ref={attachCellInput} data-testid="xlsx-cell-editor" value={editing.value}
                      disabled={Boolean(error)}
                      onChange={(event) => {
                        if (!acceptInput()) return;
                        const draft: WorkerCellDraft = { ...editing, value: event.target.value, prefill: false };
                        setEditing(draft); coordinator.setDraft(draft);
                      }}
                      onCompositionStart={() => startComposition('cell')} onCompositionEnd={endComposition}
                      onKeyDown={(event) => {
                        if (!commandForEvent(event)) event.stopPropagation();
                        if (event.nativeEvent.isComposing || event.keyCode === 229) return;
                        if (event.key === 'Enter') { commitDraft('cell', event.shiftKey ? 'up' : 'down'); event.preventDefault(); }
                        else if (event.key === 'Tab') { commitDraft('cell', event.shiftKey ? 'left' : 'right'); event.preventDefault(); }
                        else if (event.key === 'Escape') { cancelDraft(); event.preventDefault(); }
                      }}
                      onBlur={() => {
                        if (suppressBlur.current) { suppressBlur.current = false; return; }
                        commitDraft('cell');
                      }}
                      style={{ ...outline(editRect, `2px solid ${BRAND}`, '#ffffff'), padding: '0 3px',
                        font: `${13 * paintedZoom}px system-ui, sans-serif`, pointerEvents: 'auto', outline: 'none' }} />}
                  </div>
                </div>
              </div>
              {renderDock('bottom')}
            </div>
            {renderDock('right')}
          </div>
          {a11yGrid && <>
            <div style={visuallyHidden} role="grid" aria-label={a11yGrid.label}>
              <div role="row"><span role="columnheader" />{a11yGrid.columnHeaders.map((header) =>
                <span key={header.col} role="columnheader">{header.label}</span>)}</div>
              {a11yGrid.rows.map((row) => <div key={row.row} role="row"><span role="rowheader">{row.header}</span>
                {row.cells.map((cell) => <span key={cell.col} role="gridcell" aria-selected={cell.selected}>{cell.label}</span>)}
              </div>)}
            </div>
            {a11yGrid.charts.map((chart, index) => <div key={`${index}:${chart.label}`} style={visuallyHidden} role="img" aria-label={chart.label} />)}
          </>}
          {error && <div data-testid="xlsx-error" role="alert" style={{ position: 'absolute', inset: 0, display: 'grid',
            placeItems: 'center', padding: 16, textAlign: 'center', color: '#b00020' }}>{t('editor.openError')}: {error.message}</div>}
          {sheets.length > 0 && <div data-testid="xlsx-sheet-tabs" role="tablist" aria-label={t('editor.sheetTabsLabel')}
            style={{ display: 'flex', gap: 2, padding: '4px 6px', borderTop: '1px solid #e0e0e0', background: '#fafafa', overflowX: 'auto' }}>
            {sheets.map((sheet, index) => <button key={sheet.id} role="tab" aria-selected={index === active} disabled={Boolean(error)}
              onClick={() => navigate(index, selectionAt({ row: 0, col: 0 }))}
              style={{ border: 'none', padding: '4px 12px', cursor: 'pointer', background: index === active ? '#ffffff' : 'transparent',
                borderBottom: index === active ? `2px solid ${BRAND}` : '2px solid transparent', fontWeight: index === active ? 600 : 400 }}>
              {sheet.name}
            </button>)}
          </div>}
        </div>
      </FormulaBarContext.Provider>
    </EditorChromeContext.Provider>
  </XlsxCommandContext.Provider>;
}
