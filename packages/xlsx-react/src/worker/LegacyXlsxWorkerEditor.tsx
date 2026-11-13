import { usePluginGestures } from '../../../../shared/plugin-host/usePluginGestures';
import {
  anchorSelection, clearedCells, copiedRange, createSelectionDataNavigation, mergeAnchorCell, reachesSheetEdge,
  selectionDataRequest, selectionFrameContains, selectionHidden, selectionMerges, workbookSelectionMerges,
} from '../selectionData';
import {
  buildA11yGrid, cellAtPoint, cellRect, chartRegionAtPoint, extendTo, fromTsv, hyperlinkAtCell, isProposalsAvailable,
  moveFocus, normalizeRange, paintDisplayList, parseHyperlinkLocation, rangeRect, safeExternalHyperlink, selectionAt,
  selectionKeyReducer, toTsv, WorkbookEditPeerFailedError,
} from '@betteroffice/xlsx';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '../sheetLimits';
import {
  createWorkbookRecoveryMutators, localNowSerial, WorkbookRecoveryRefusal, workbookEditPeerOperations,
  workbookSessionInternals,
} from '@betteroffice/xlsx/worker';
import type {
  CapturedFormat, CellAddr, CellEdit, CellInputEdit, CellRange, ChartResizeEdges, Direction, DisplayList, EditResult,
  Proposal, Rect, Selection, WorkbookFrame, WorkbookHandle, WorkbookSheetView, XlsxEditResult,
} from '@betteroffice/xlsx';
import type { WorkbookReplayOp } from '@betteroffice/xlsx/worker';
import { useDefaultFonts } from '../useDefaultFonts';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, MouseEvent, SetStateAction, SyntheticEvent } from 'react';
import { flushSync } from 'react-dom';
import { useContextMenu } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { engineStructuralRefusal, REFUSAL_CAUSE_KEYS } from '../commands/refusalCause';

import type { InputDraft, InputSeal } from '../commands/inputCoordinator';
import { inPluginChrome, pluginEventStore } from '../commands/pluginEvents';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { useSaveRequest } from '../commands/useSaveRequest';
import { useWorkerXlsxCommands } from '../commands/useWorkerXlsxCommands';
import {
  createWorkerInputCoordinator, inputRefusal, WorkerInputNotReadyError, WorkerInputRefusal,
  type WorkerInputCoordinatorHooks, type WorkerInputLease,
} from '../commands/workerInputCoordinator';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import { EditorToolbar } from '../components/EditorToolbar';
import { FillHandle, fillHandleAnchor } from '../components/FillHandle';
import { EditorChromeContext } from '../components/EditorToolbarContext';
import type { BorderStyle } from '../components/Toolbar';
import { FormulaBarContext, type FormulaBarBinding } from '../components/toolbar/FormulaBar';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { xlsxContextMenuSections, type XlsxContextFreezeTarget } from '../contextMenu/menus';
import { useTranslation } from '../i18n';
import type { XlsxPluginEditorAccess, XlsxRevealAlignment } from '../plugins/createPluginClients';

import { PluginOverlays } from '../plugins/PluginOverlays';
import { PluginRowGutter } from '../plugins/PluginRowGutter';
import { ROW_GUTTER_WIDTH } from '../plugins/rowDecorations';
import { frozenGridExtent } from '../plugins/geometry';
import { PluginDock, useDockArea, type DockPlacement } from '../plugins/PluginPanels';
import { useXlsxPluginHost } from '../plugins/useXlsxPluginHost';
import { ProposalsPanel } from '../proposals/ProposalsPanel';
import { cellInViewport, coversWholeTrack, deriveLimits, expandRangeToMergedCells, extendScrollExtent, laidOutCellRect, nearestCellScroll, positionAtPoint, scaledRect } from '../viewer/sessionGeometry';
import { cellEditorColors } from '../viewer/cellEditorColors';
import { ACTIVE_CELL_WIDTH, marqueeBox, marqueeFills, MARQUEE_COLOR, MARQUEE_WIDTH, SELECTION_FILL } from '../selectionOutline';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';

import { useLegacySessionWorkbook, EditableWorkbookSession, type EditableSessionWorkbookProps } from './useEditableSessionWorkbook';
import type { WorkerPaintRequest, WorkerPaintResult } from './WorkerPaintSource';
import { WorkerPaintSource } from './WorkerPaintSource';
import { A11yGridTable } from '../components/A11yGridCell';
import {
  COLUMN_HEADER_HEIGHT, SheetGrid, columnWidthFromPx, rowHeightFromPx,
} from '../components/SheetGrid';
import type { SheetGridHeaderHandlers } from '../components/SheetGrid';
import { chartResizeRect, chartResizeSpace, type ChartDragPreview, type SelectedChart } from './chartResize';
import { ChartResizeHandles } from './ChartResizeHandles';
import { useChartResizeGesture } from './useChartResizeGesture';

const BRAND = '#217346';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, border: 0,
  overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap',
};
const chartKeys: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1],
};

type WorkerCellDraft = InputDraft & { prefill?: boolean; modified?: boolean; draftId?: number; revision?: number; unchanged?: boolean };
type PreviewPredecessor = { operation?: WorkbookReplayOp; input?: string; index: number };

function address(cell: CellAddr): string {
  let column = '';
  for (let col = cell.col + 1; col > 0; col = Math.floor((col - 1) / 26)) {
    column = String.fromCharCode(65 + (col - 1) % 26) + column;
  }
  return `${column}${cell.row + 1}`;
}

function rangeA1(range: CellRange): string {
  return `${address({ row: range.top, col: range.left })}:${address({ row: range.bottom, col: range.right })}`;
}

function selectedRange(selection: Selection): string {
  return rangeA1(normalizeRange(selection));
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

export function LegacyXlsxWorkerEditor(props: EditableSessionWorkbookProps) {
  const { t, lang } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  usePluginGestures(rootRef);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const editorInputRef = useRef<HTMLInputElement | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const editorRectRef = useRef<ReturnType<typeof cellRect>>(null);
  const runRef = useRef<EditableWorkbookSession | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [commands] = useState(createXlsxCommandController);
  const [active, setActive, activeRef] = useSyncedState(0);
  const [view, setView, viewRef] = useSyncedState<WorkbookSheetView | null>(null);
  const [selection, setRawSelection, selectionRef] = useSyncedState<Selection | null>(null);
  const resolveSelectionRef = useRef<(selected: Selection) => void>(() => {});
  const setSelection = useCallback((next: SetStateAction<Selection | null>) => {
    const selected = typeof next === 'function' ? next(selectionRef.current) : next;
    if (selected && !runRef.current?.peer && (selected.anchor.row || selected.anchor.col || selected.focus.row || selected.focus.col)) {
      resolveSelectionRef.current(selected);
    }
    else setRawSelection(selected ? anchorSelection(selected, workbookSelectionMerges(runRef.current?.peer, activeRef.current)) : null);
  }, []);
  const [zoom, setZoom, zoomRef] = useSyncedState(1);
  const [editing, setEditing, editingRef] = useSyncedState<InputDraft | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const previewCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const [refusal, setRefusal] = useState<string | null>(null);
  const [previewDraft, setPreviewDraft] = useState<InputDraft | null>(null);
  const [painted, setPainted] = useState<WorkerPaintResult | null>(null);
  const [selectedChart, setSelectedChart, chartRef] = useSyncedState<SelectedChart | null>(null);
  const [chartPreview, setChartPreview] = useState<ChartDragPreview | null>(null);
  const [capturedFormat, setCapturedFormat, capturedFormatRef] = useSyncedState<CapturedFormat | null>(null);
  const paintFormatSource = useRef<string | null>(null);
  const borderStyleRef = useRef<BorderStyle | undefined>(undefined);
  const borderColorRef = useRef<string | undefined>(undefined);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen, panelRef] = useSyncedState(false);
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});
  const mutationRef = useRef(0);
  const navigationRef = useRef(0);
  const [selectionNavigation] = useState(createSelectionDataNavigation);
  const sourceRef = useRef<WorkerPaintSource | null>(null);
  const dragging = useRef(false);
  const clickStart = useRef<CellAddr | null>(null);
  const chartDrag = useRef<{ id: string; sheet: number; zoom: number; clientX: number; clientY: number } | null>(null);
  const nudge = useRef<{ id: string; sheet: number; dx: number; dy: number } | null>(null);
  const nudgeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flushNudgeRef = useRef<(retained?: boolean) => void>(() => {});
  const suppressBlur = useRef(false);
  const suppressFormulaBlur = useRef(false);
  const composition = useRef<{ source: InputDraft['source']; done: Promise<boolean>; settle(ended: boolean): void } | null>(null);
  const paintWaiters = useRef<{ request: WorkerPaintRequest; sequence: number; resolve(value: boolean): void }[]>([]);
  const coordinatorRef = useRef<ReturnType<typeof createWorkerInputCoordinator> | null>(null);
  const draftRevision = useRef(0);
  const settledDraftRevision = useRef(0);
  const focus = useCallback(() => scrollRef.current?.focus({ preventScroll: true }), []);
  const reportInputError = useCallback((error: unknown) => {
    const local = inputRefusal(error);
    if (local || error instanceof XlsxCommandAdmissionError) {
      setRefusal((local ?? error as Error).message);
      return;
    }
    setRefusal(error instanceof Error ? error.message : String(error));
  }, []);
  const acceptInput = () => Boolean(runRef.current?.current && !runRef.current.retiring && !runRef.current.failure && !propsRef.current.readOnly);

  const apiBridgeRef = useRef<WorkerEditorApiBridge | null>(null);
  const [apiBridge] = useState<WorkerEditorApiBridge>(() => ({
    getPositionAtPoint: (x, y) => apiBridgeRef.current?.getPositionAtPoint?.(x, y) ?? null,
    coordinator: () => coordinatorRef.current,
    readOnly: () => propsRef.current.readOnly ?? false,
    allowHostProposals: () => propsRef.current.allowHostProposals ?? false,
    clearSelection: () => apiBridgeRef.current?.clearSelection(),
    focus,
    refreshProposals: () => apiBridgeRef.current?.refreshProposals(),
    recoverInput: () => apiBridgeRef.current!.recoverInput(),
    selectCells: (...args) => apiBridgeRef.current?.selectCells(...args) ?? false,
    selectCellsAsync: (...args) => apiBridgeRef.current?.selectCellsAsync(...args) ?? Promise.resolve(false),
    apply: (result, op) => apiBridgeRef.current?.apply(result, op),
    previewEdits: (sheet, edits, op) => apiBridgeRef.current?.previewEdits?.(sheet, edits, op) ?? Promise.resolve(),
    canNavigateSync: () => apiBridgeRef.current?.canNavigateSync?.() ?? false,
  }));
  const { run, error, loading } = useLegacySessionWorkbook(props, commands.store, apiBridge);
  const heldPreview = useRef<{ canvas: HTMLCanvasElement; displayList: DisplayList; request: WorkerPaintRequest } | null>(null);
  const paintWithDefaultFonts = useDefaultFonts(run, run?.session.call.decodeFontFace, props.defaultFonts,
    () => {
      if (run?.current && !run.failure) sourceRef.current?.schedule();
      repaintHeldPreview();
    },
    () => {
      if (canvasRef.current) canvasRef.current.dataset.defaultFonts = 'settled';
      repaintHeldPreview();
    });
  const paintPreview = (canvas: HTMLCanvasElement, displayList: DisplayList, request: WorkerPaintRequest) => {
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Preview canvas context is unavailable');
    canvas.width = Math.round(request.width * request.dpr);
    canvas.height = Math.round(request.height * request.dpr);
    canvas.style.width = `${request.width}px`;
    canvas.style.height = `${request.height}px`;
    const settled = paintWithDefaultFonts(displayList, () => paintDisplayList(context, displayList, request.dpr * request.zoom));
    heldPreview.current = settled ? null : { canvas, displayList, request };
    canvas.dataset.previewReady = settled ? 'true' : 'pending';
  };
  function repaintHeldPreview() {
    const held = heldPreview.current;
    if (!held) return;
    if (previewCanvasRef.current !== held.canvas || held.canvas.dataset.previewReady !== 'pending') {
      heldPreview.current = null;
      return;
    }
    paintPreview(held.canvas, held.displayList, held.request);
  }
  useEffect(() => () => { heldPreview.current = null; }, [run]);
  useEffect(() => { repaintHeldPreview(); }, [props.defaultFonts]);
  runRef.current = run;
  const inputHooks = useMemo(() => ({ current: null as WorkerInputCoordinatorHooks | null }), [run]);
  const acceptedCells = useMemo(() => new Map<string, string>(), [run]);
  const draftOperations = useMemo(() => new WeakMap<InputDraft, WorkbookReplayOp>(), [run]);
  const previewPredecessors = useMemo(() => new Map<WorkbookReplayOp, Map<string, PreviewPredecessor>>(), [run]);
  const draftDiscards = useMemo(() => new WeakMap<InputDraft, () => void>(), [run]);
  const operationPeer = useMemo(() => ({ current: run?.editPeer ?? null }), [run]);
  if (run?.editPeer && !run.recovering) operationPeer.current = run.editPeer;
  const previewOps = useMemo(() => new Map<string, WorkbookReplayOp>(), [run]);
  const unchangedPreviews = useMemo(() => new WeakSet<WorkbookReplayOp>(), [run]);
  const previewJournal = useMemo(() => [] as WorkbookReplayOp[], [run]);
  const completedPreviews = useMemo(() => new WeakMap<WorkbookReplayOp, number>(), [run]);
  const unpreviewedOps = useMemo(() => new Set<WorkbookReplayOp>(), [run]);
  const previewState = useMemo(() => ({ revision: 0, sequence: Infinity, request: null as WorkerPaintRequest | null, committed: null as InputDraft | null }), [run]);
  const retirement = useMemo(() => {
    const callbacks = new Set<() => void>();
    return {
      resolve() { for (const callback of callbacks) callback(); callbacks.clear(); },
      wait(promise: Promise<void>) {
        return new Promise<void>((resolve, reject) => {
          const complete = () => { callbacks.delete(complete); resolve(); };
          callbacks.add(complete);
          void promise.then(complete, (error) => { callbacks.delete(complete); reject(error); });
        });
      },
    };
  }, [run]);
  const coordinator = useMemo(() => {
    const owner = run;
    const input = createWorkerInputCoordinator({
      generation: () => owner?.generation ?? 0,
      capture: () => ({ sheet: activeRef.current, target: JSON.stringify(selectionRef.current) }),
      isReady: () => owner?.ready ?? false,
      isAcknowledged: () => !owner?.editPeer || owner.editPeer.acknowledgedSequence >= owner.editPeer.sentSequence,
      whenReady: () => owner?.whenHydrated() ?? Promise.reject(new XlsxCommandAdmissionError('editor-unavailable')),
      requestHydration: (reason) => owner?.requestHydration(reason),
      flushEdits: () => owner?.editPeer?.flush() ?? Promise.reject(new XlsxCommandAdmissionError('editor-unavailable')),
      async acknowledgeEdits() {
        if (!owner?.editPeer || owner.recovering || owner.retiring) return;
        await retirement.wait(workbookEditPeerOperations(owner.editPeer).whenAcknowledged());
        if (!owner.retiring) await inputHooks.current?.acknowledgeEdits?.();
      },
      preview: (draft) => inputHooks.current!.preview(draft),
      resolveDraft: (draft) => inputHooks.current!.resolveDraft!(draft),
      sameDraft: (a, b) => (a as WorkerCellDraft).draftId === (b as WorkerCellDraft).draftId,
      seal: () => inputHooks.current?.seal() ?? {},
      sync: () => inputHooks.current?.sync(),
      write: (draft, applied) => inputHooks.current?.write(draft, applied) ?? false,
      onError: (error) => owner?.fail(error),
      onRefusal(error, draft) {
        if (owner?.retiring) { owner.reportRefusal(error); return; }
        const message = error instanceof Error ? error.message : String(error);
        if (owner?.recovering) flushSync(() => setRefusal(message));
        else setRefusal(message);
        inputHooks.current?.sync?.();
        const revision = (draft as WorkerCellDraft | undefined)?.revision ?? 0;
        if (draft && !propsRef.current.readOnly && revision > settledDraftRevision.current && revision === draftRevision.current) {
          input.setDraft(draft);
          if (draft.source === 'cell') setEditing(draft);
          else setFormulaDraft(draft.value);
        }
      },
    });
    return input;
  }, [run, inputHooks, reportInputError, retirement]);
  coordinatorRef.current = coordinator;

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
    if (!run?.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (run.failure && !run.recovering) return Promise.reject(run.failure);
    return Promise.resolve();
  }, [run]);

  const relinkPredecessors = useCallback((key: string, op: WorkbookReplayOp, previous?: PreviewPredecessor) => {
    for (const predecessors of previewPredecessors.values()) {
      const predecessor = predecessors.get(key);
      if (predecessor?.operation === op) {
        predecessor.operation = previous?.operation;
        predecessor.input = previous?.input;
        predecessor.index = previous?.index ?? -1;
      }
    }
    const predecessors = previewPredecessors.get(op);
    predecessors?.delete(key);
    if (!predecessors?.size) previewPredecessors.delete(op);
  }, [previewPredecessors]);

  const preview = useCallback(async (draft?: InputDraft, discardOnFailure = true): Promise<void> => {
    if (!run || run.failure || run.recovering || run.retiring) return;
    if (draft) {
      const key = `${draft.sheet}:${draft.row}:${draft.col}`;
      const preceding = previewOps.get(key);
      const committed = previewState.committed;
      const precedingInput = acceptedCells.get(key) ?? (committed &&
        `${committed.sheet}:${committed.row}:${committed.col}` === key ? committed.value : undefined);
      acceptedCells.set(key, draft.value);
      const op: WorkbookReplayOp = draftOperations.get(draft) ?? {
        method: 'editCell', args: [draft.sheet, draft.row, draft.col, draft.value], calculation: {
          nowSerial: localNowSerial(Date.now()),
          randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0,
        },
      };
      draftOperations.set(draft, op);
      if (preceding !== op) {
        let predecessors = previewPredecessors.get(op);
        if (!predecessors) previewPredecessors.set(op, predecessors = new Map());
        predecessors.set(key, {
          operation: preceding, input: precedingInput, index: preceding ? previewJournal.indexOf(preceding) : -1,
        });
      }
      if ((draft as WorkerCellDraft).unchanged) unchangedPreviews.add(op);
      if (preceding && preceding !== op && unchangedPreviews.has(preceding)) {
        const index = previewJournal.indexOf(preceding);
        if (index >= 0) previewJournal.splice(index, 1);
      }
      previewOps.set(key, op);
      if (!previewJournal.includes(op)) previewJournal.push(op);
      previewState.sequence = Infinity;
      const revision = ++previewState.revision;
      const discard = () => {
        const previous = previewPredecessors.get(op)?.get(key);
        const index = previewJournal.indexOf(op);
        if (index >= 0) previewJournal.splice(index, 1);
        relinkPredecessors(key, op, previous);
        const current = previewOps.get(key) === op;
        if (current) {
          if (previous?.operation) {
            previewOps.set(key, previous.operation);
            if (!previewJournal.includes(previous.operation)) {
              previewJournal.splice(Math.max(0, previous.index), 0, previous.operation);
            }
          } else previewOps.delete(key);
          if (previous?.operation && previous.input !== undefined) acceptedCells.set(key, previous.input);
          else acceptedCells.delete(key);
        }
        if (!previewOps.size && ((current && !unpreviewedOps.size) || revision === previewState.revision) && !run.retiring) {
          setPreviewDraft(previewState.committed);
          previewState.request = null;
        } else if (current && !run.retiring) {
          const pending = [...previewOps.keys()].reverse()[0];
          if (pending) {
            const [sheet, row, col] = pending.split(':').map(Number);
            setPreviewDraft({ ...draft, sheet, row, col, value: acceptedCells.get(pending) ?? '' });
          }
        }
        const complete = !unpreviewedOps.size && previewJournal.every((pending) => completedPreviews.has(pending));
        previewState.sequence = complete ? Math.max(0, ...previewJournal.map((pending) => completedPreviews.get(pending)!)) : Infinity;
        if (complete) sourceRef.current?.schedule();
      };
      if (discardOnFailure) draftDiscards.set(draft, discard);
      if (run.ready) {
        if (!unpreviewedOps.size) {
          flushSync(() => setPreviewDraft(draft));
        }
        return;
      }
      const request = capture();
      if (unpreviewedOps.size) { await boundary(); return; }
      if (request) {
        const internal = workbookSessionInternals.get(run.session);
        if (!internal?.preview) throw new Error('Worker cell preview is unavailable');
        let frame: WorkbookFrame;
        try {
          frame = await internal.preview(request.viewport, request.sheet,
            previewJournal.filter((pending) => !completedPreviews.has(pending)));
        }
        catch (error) {
          if (discardOnFailure) discard();
          throw error;
        }
        if (run.retiring || run.recovering || run.failure || !run.current || revision !== previewState.revision) return;
        previewState.request = request;
        previewState.sequence = previewJournal.every((pending) => completedPreviews.has(pending)) ?
          run.editPeer?.sentSequence ?? Infinity : Infinity;
        if (unpreviewedOps.size) return;
        flushSync(() => setPreviewDraft(draft));
        if (run.ready) return;
        const canvas = previewCanvasRef.current;
        if (!canvas) throw new Error('Preview canvas context is unavailable');
        paintPreview(canvas, frame.displayList, request);
      }
    }
    await boundary();
  }, [run, capture, boundary, acceptedCells, previewOps, previewState, draftOperations, draftDiscards, previewPredecessors, relinkPredecessors, completedPreviews, unpreviewedOps, previewJournal, unchangedPreviews]);

  const previewEdits = useCallback(async (sheet: number, edits: readonly CellInputEdit[], operation?: WorkbookReplayOp) => {
    if (operation?.method === 'applyEdits' && operation.args[0].steps.some((step) => step.op !== 'setCellInputs')) {
      unpreviewedOps.add(operation);
      previewState.revision += 1;
      previewState.request = null;
      setPreviewDraft(null);
      return async () => { unpreviewedOps.delete(operation); };
    }
    if (operation) operation.calculation ??= {
      nowSerial: localNowSerial(Date.now()),
      randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0,
    };
    const preceding = new Map<string, PreviewPredecessor>();
    const recorded = new Map<string, WorkbookReplayOp>();
    for (const edit of edits) {
      if (!Number.isInteger(sheet) || sheet < 0 || sheet >= (run?.session.state.sheets.length ?? 0) ||
        !Number.isInteger(edit.row) || edit.row < 0 || edit.row > 1048575 ||
        !Number.isInteger(edit.col) || edit.col < 0 || edit.col > 16383 ||
        !edit.input.startsWith('=') && [...edit.input.replace(/^'/, '')].length > 32767) {
        throw new WorkerInputRefusal('Invalid cell input or target');
      }
    }
    for (const edit of edits) {
      const key = `${sheet}:${edit.row}:${edit.col}`;
      const op = operation ?? { method: 'editCell' as const, args: [sheet, edit.row, edit.col, edit.input] as [number, number, number, string] };
      const previous = previewOps.get(key);
      let predecessors = previewPredecessors.get(op);
      if (!predecessors) previewPredecessors.set(op, predecessors = new Map());
      const predecessor = preceding.get(key) ?? (previous === op ? predecessors.get(key) : undefined) ?? {
        operation: previous, input: acceptedCells.get(key) ?? (previewState.committed &&
          `${previewState.committed.sheet}:${previewState.committed.row}:${previewState.committed.col}` === key ?
          previewState.committed.value : undefined), index: previous ? previewJournal.indexOf(previous) : -1,
      };
      preceding.set(key, predecessor);
      predecessors.set(key, predecessor);
      acceptedCells.set(key, edit.input);
      if (previous && previous !== op && unchangedPreviews.has(previous)) {
        const index = previewJournal.indexOf(previous);
        if (index >= 0) previewJournal.splice(index, 1);
      }
      previewOps.set(key, op);
      recorded.set(key, op);
      if (!previewJournal.includes(op)) previewJournal.push(op);
    }
    const restore = async () => {
      let changed = false;
      for (const edit of edits) {
        const key = `${sheet}:${edit.row}:${edit.col}`;
        const op = recorded.get(key);
        const predecessor = preceding.get(key);
        const previous = predecessor?.operation;
        const value = predecessor?.input;
        if (op) relinkPredecessors(key, op, predecessor);
        if (previewOps.get(key) !== op) continue;
        if (previous) {
          previewOps.set(key, previous);
          if (!previewJournal.includes(previous)) previewJournal.splice(Math.max(0, predecessor!.index), 0, previous);
        }
        else previewOps.delete(key);
        if (previous && value !== undefined) acceptedCells.set(key, value);
        else acceptedCells.delete(key);
        changed = true;
      }
      for (const op of new Set(recorded.values())) {
        const index = previewJournal.indexOf(op);
        if (index >= 0) previewJournal.splice(index, 1);
      }
      if (!changed || !run || run.retiring || run.failure) return;
      const pending = [...previewOps.entries()].reverse()[0];
      if (pending) {
        const [sheet, row, col] = pending[0].split(':').map(Number);
        const draft: InputDraft = { generation: run.generation, sheet, row, col,
          value: acceptedCells.get(pending[0]) ?? '', source: 'cell' };
        draftOperations.set(draft, pending[1]);
        await preview(draft);
        if (previewJournal.every((op) => completedPreviews.has(op))) {
          previewState.sequence = run.editPeer?.sentSequence ?? Infinity;
          sourceRef.current?.schedule();
        }
      } else {
        previewState.revision += 1;
        previewState.request = null;
        previewState.sequence = run.editPeer?.sentSequence ?? Infinity;
        setPreviewDraft(previewState.committed);
      }
    };
    try {
      const draft: InputDraft | undefined = edits[0] ? { generation: run!.generation, sheet, row: edits[0].row,
        col: edits[0].col, value: edits[0].input, source: 'cell' } : undefined;
      if (draft) draftOperations.set(draft, recorded.get(`${sheet}:${draft.row}:${draft.col}`)!);
      const prepared = draft ? preview(draft, false) : Promise.resolve();
      await prepared;
    } catch (error) {
      await restore();
      throw error;
    }
    return restore;
  }, [run, preview, relinkPredecessors, acceptedCells, previewOps, previewPredecessors, previewState, draftOperations, completedPreviews, unpreviewedOps, previewJournal, unchangedPreviews]);

  const refreshProposals = useCallback(() => {
    const owner = runRef.current;
    if (!owner?.peer) { setProposals([]); return; }
    setProposals(owner.peer.listProposals());
    if (!owner.recovering) sourceRef.current?.schedule();
    commands.refresh();
  }, [commands]);

  const peerPaint = useCallback(() => {
    if (run?.retiring || run?.recovering) return;
    sourceRef.current?.schedule();
  }, [run]);

  const trackCommit = useCallback((op?: WorkbookReplayOp) => {
    if (op) {
      completedPreviews.set(op, run?.editPeer?.sentSequence ?? 0);
      unpreviewedOps.delete(op);
    }
    const complete = !unpreviewedOps.size && previewJournal.every((pending) => completedPreviews.has(pending));
    previewState.sequence = complete ? Math.max(0, ...previewJournal.map((pending) => completedPreviews.get(pending)!)) : Infinity;
    if (!run?.recovering && !run?.retiring) sourceRef.current?.schedule();
  }, [run, previewState, completedPreviews, unpreviewedOps, previewJournal]);

  const closeWrittenDraft = (draft: InputDraft) => {
    const live = coordinator.draft;
    if (live && (live.source !== draft.source || live.sheet !== draft.sheet || live.row !== draft.row ||
      live.col !== draft.col || (live as WorkerCellDraft).revision !== (draft as WorkerCellDraft).revision ||
      live.value !== draft.value && (live as WorkerCellDraft).modified !== false)) return;
    if (live) coordinator.setDraft(null);
    if (draft.source === 'cell') {
      const current = editingRef.current;
      if (current && current.sheet === draft.sheet && current.row === draft.row && current.col === draft.col &&
        (current as WorkerCellDraft).revision === (draft as WorkerCellDraft).revision &&
        (current.value === draft.value || (current as WorkerCellDraft).modified === false)) {
        suppressBlur.current = true;
        setEditing(null);
      }
    } else setFormulaDraft(null);
  };

  const apply = useCallback((result: EditResult | XlsxEditResult, op?: WorkbookReplayOp) => {
    const owner = run;
    if (!owner?.peer || owner.retiring) return;
    try {
      if (!op && previewState.request && (!('ok' in result) || result.ok) && result.applied) {
        previewState.revision += 1;
        previewState.request = null;
        setPreviewDraft(null);
      }
      trackCommit(op);
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
  }, [run, commands, peerPaint, refreshProposals, trackCommit, previewState]);

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
    const owner = run;
    if (!owner?.peer || !owner.current || owner.retiring) return false;
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
    operationPeer.current = edits;
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
      flushNudgeRef.current(true);
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
    getPositionAtPoint: (x, y) => {
      const painted = sourceRef.current?.painted;
      if (!painted || painted.request.generation !== runRef.current?.generation) return null;
      return positionAtPoint(painted.displayList, painted.request.sheet, canvasRef.current, x, y);
    },
    coordinator: () => coordinator, readOnly: () => propsRef.current.readOnly ?? false,
    allowHostProposals: () => propsRef.current.allowHostProposals ?? false, focus,
    clearSelection: () => {
      coordinator.settle(); setEditing(null); setFormulaDraft(null); setSelection(null); setSelectedChart(null);
      setCapturedFormat(null); commands.refresh();
    },
    refreshProposals, recoverInput, selectCells: place, previewEdits,
    canNavigateSync: () => previewState.request === null,
    selectCellsAsync: async (sheet, next) => place(sheet, next), apply,
  };

  const pluginAccessRef = useRef<XlsxPluginEditorAccess | null>(null);
  const [pluginAccess] = useState<XlsxPluginEditorAccess>(() => ({
    handle: () => runRef.current?.ready && !runRef.current.failure ? runRef.current.peer : null,
    admit: (handle, operation) => pluginAccessRef.current!.admit(handle, operation),
    readOnlyRefusal: (handle) => pluginAccessRef.current!.readOnlyRefusal(handle),
    applyEdits: (...args) => pluginAccessRef.current!.applyEdits(...args),
    admitEdits: (...args) => pluginAccessRef.current!.admitEdits!(...args),
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
  const pluginHost = useXlsxPluginHost({ locale: lang, focusRoot: scrollRef, ...props, access: pluginAccess, commands, editGate: false,
    readOnly: props.readOnly ?? false, handle: run?.ready && !run.failure ? run.peer : null,
    selection: pluginSelection, proposals, staleFor, proposalsPending: () => coordinator.pending, canvasRef, t,
    selectionPaintReady: (sheet) => sourceRef.current?.selectionPaintReady(sheet, activeRef.current) ?? (() => true),
    positionAtPoint: (clientX, clientY) => {
      const owner = runRef.current;
      return owner?.current && !owner.retiring && !owner.failure ? apiBridge.getPositionAtPoint?.(clientX, clientY) ?? null : null;
    } });
  const presentGridRef = useRef(pluginHost.presentGrid);
  presentGridRef.current = pluginHost.presentGrid;
  const [dockAreaRef, dockArea] = useDockArea(pluginHost.managed);

  const cancelPendingFrames = useCallback(() => {
    for (const waiter of paintWaiters.current.splice(0)) waiter.resolve(false);
  }, []);

  useLayoutEffect(() => {
    setRefusal(null);
    setPainted(null); setPreviewDraft(null); setSelection(null); setView(null); setEditing(null);
    setFormulaDraft(null); setSelectedChart(null); setChartPreview(null); setFocusedCell(null);
    editorRectRef.current = null;
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
        canvas.dataset.defaultFonts = paintWithDefaultFonts(painted.displayList,
          () => paintDisplayList(context, painted.displayList, request.dpr * request.zoom)) ? 'settled' : 'pending';
        canvas.dataset.workerSequence = String(painted.sequence);
        canvas.dataset.paintSource = painted.source;
        canvas.dataset.workerZoom = String(request.zoom);
        setPainted(painted);
        if (painted.sequence >= previewState.sequence) {
          previewOps.clear();
          previewJournal.length = 0;
          previewPredecessors.clear();
          acceptedCells.clear();
          previewState.request = null;
          previewState.committed = null;
          flushSync(() => setPreviewDraft(null));
        }
      },
      onInstalled(painted) {
        const request = painted.request;
        const sheetId = run.session.state.sheets[request.sheet]?.id;
        presentGridRef.current(sheetId === undefined ? null : {
          frame: painted.displayList, zoom: request.zoom, viewport: request.viewport, version: painted.version, sheetId,
          sheet: request.sheet,
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
    flushNudgeRef.current(true);
    syncDraft();
    composition.current?.settle(true);
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

  const navigationContext = () => {
    const owner = runRef.current;
    if (!owner?.current || owner.retiring || owner.failure || !selectionRef.current) return null;
    const basis = owner.session.confirmedBasis;
    return { owner, sheet: activeRef.current, selection: selectionRef.current, navigation: navigationRef.current,
      version: `${basis.epoch}:${basis.version}:${owner.peer?.version() ?? basis.version}:${owner.session.state.version}` };
  };
  const deferNavigation = (task: () => void) => selectionNavigation.defer(navigationContext(), task);
  const hidden = () => selectionHidden(sourceRef.current?.painted?.displayList.grid);
  const merges = () => {
    const painted = sourceRef.current?.painted;
    const owner = runRef.current;
    const context = navigationContext();
    const visible = selectionMerges(painted?.request.sheet === activeRef.current &&
      painted.version === owner?.session.confirmedBasis.version ? painted.mergedRanges : undefined);
    return workbookSelectionMerges(owner?.peer, activeRef.current,
      context ? selectionNavigation.merges(context, visible) : visible);
  };
  const navigationOptions = () => ({
    context: navigationContext, limits, hidden, merges,
    known: (context: import('../selectionData').NavigationContext, cell: CellAddr) => {
      const owner = runRef.current;
      const painted = sourceRef.current?.painted;
      return Boolean(owner?.peer || selectionNavigation.known(context, cell) ||
        painted?.request.sheet === context.sheet && painted.version === owner?.session.confirmedBasis.version &&
        selectionFrameContains(painted.displayList.grid, cell));
    },
    query: (context: import('../selectionData').NavigationContext,
      direction: Direction | 'usedRange' | 'rowEnd' | 'merge' | undefined) =>
      (runRef.current!.peer ?? runRef.current!.session.call).selectionData(context.sheet,
        context.selection.focus.row, context.selection.focus.col, direction),
    apply: (next: Selection) => {
      setRawSelection(next);
      revealSelection(activeRef.current, next);
      return true;
    },
  });
  resolveSelectionRef.current = (selected) => {
    void selectionNavigation.select(selected, { ...navigationOptions(), apply: (next) => {
      setRawSelection(next); return true;
    } });
  };
  const beforeNavigationInput = (task: () => void) => runRef.current?.peer ? deferNavigation(task) :
    selectionNavigation.beforeInput({ ...navigationOptions(), apply: (next) => { setRawSelection(next); return true; } },
      task, () => reportInputError(new XlsxCommandAdmissionError('editor-unavailable')));
  const limits = () => deriveLimits(sourceRef.current?.painted?.displayList ?? null,
    (scrollRef.current?.clientHeight ?? 0) / zoomRef.current);
  const draftFor = (source: InputDraft['source'], cell: CellAddr, value: string): WorkerCellDraft => ({
    generation: runRef.current!.generation, sheet: activeRef.current, ...cell, source, value,
    revision: ++draftRevision.current, draftId: draftRevision.current,
  });
  const syncDraft = () => {
    const draft = coordinator.draft;
    const input = draft?.source === 'cell' ? editorInputRef.current : formulaInputRef.current;
    if (draft && input && draft.value !== input.value) {
      const next: WorkerCellDraft = { ...draft, value: input.value, prefill: false, modified: true,
        revision: ++draftRevision.current };
      coordinator.setDraft(next);
      if (next.source === 'cell') setEditing(next);
      else setFormulaDraft(next.value);
    }
  };
  const seal = (): InputSeal => {
    if (chartDrag.current || chartResizeActive()) return { refused: 'gesture-active' };
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
    async acknowledgeEdits() {
      if (!run?.peer || !run.editPeer) return;
      for (const op of previewJournal.slice()) {
        const sequence = completedPreviews.get(op);
        if (sequence === undefined || sequence > run.editPeer.acknowledgedSequence) continue;
        const keys = [...(previewPredecessors.get(op)?.keys() ?? [])];
        if (keys[0]) {
          const [sheet, row, col] = keys[0].split(':').map(Number);
          previewState.committed = { generation: run.generation, sheet, row, col,
            value: run.peer.cell(sheet, row, col).input, source: 'cell' };
        }
        for (const key of keys) {
          const [sheet, row, col] = key.split(':').map(Number);
          relinkPredecessors(key, op, { input: run.peer.cell(sheet, row, col).input, index: -1 });
          if (previewOps.get(key) === op) {
            previewOps.delete(key);
            acceptedCells.delete(key);
          }
        }
        previewJournal.splice(previewJournal.indexOf(op), 1);
        completedPreviews.delete(op);
        if (previewDraft && draftOperations.get(previewDraft) === op) {
          draftOperations.delete(previewDraft);
          draftDiscards.delete(previewDraft);
        }
      }
      if (!unpreviewedOps.size && previewJournal.every((op) => completedPreviews.has(op))) {
        previewState.sequence = run.editPeer.sentSequence;
      }
    },
    async resolveDraft(draft) {
      if (!draft.value.startsWith('=') && [...draft.value.replace(/^'/, '')].length > 32767) {
        throw new WorkerInputRefusal('Cell text exceeds 32,767 characters');
      }
      if ((draft as WorkerCellDraft).modified !== false) return draft;
      const key = `${draft.sheet}:${draft.row}:${draft.col}`;
      const accepted = acceptedCells.get(key);
      if (accepted !== undefined) return { ...draft, value: accepted, modified: true } as WorkerCellDraft;
      const internal = run && workbookSessionInternals.get(run.session);
      const value = run?.peer ? run.peer.cell(draft.sheet, draft.row, draft.col).input :
        internal?.cellInput ? await internal.cellInput(draft.sheet, draft.row, draft.col) :
        (await run!.session.call.cellInputs(draft.sheet, address(draft))).cells[0]?.[0]?.input ?? '';
      return { ...draft, value, modified: true, unchanged: true } as WorkerCellDraft;
    },
    write(draft, markApplied) {
      const owner = run;
      if (!owner?.ready || !owner.peer || !owner.editPeer) return false;
      if (owner.peer.cell(draft.sheet, draft.row, draft.col).input !== draft.value) {
        markApplied.check();
        const op = draftOperations.get(draft);
        const operations = workbookEditPeerOperations(owner.recovering ? operationPeer.current! : owner.editPeer);
        let result: EditResult;
        try {
          result = op ? (run!.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult :
            owner.editPeer.editCell(draft.sheet, draft.row, draft.col, draft.value);
        } catch (error) {
          if (inputRefusal(error)) draftDiscards.get(draft)?.();
          throw error;
        }
        markApplied();
        apply(result, op);
      } else trackCommit(draftOperations.get(draft));
      if (!owner.retiring) closeWrittenDraft(draft);
      return true;
    },
  };

  const openEditor = (seed?: string) => {
    if (beforeNavigationInput(() => openEditor(seed))) return;
    const owner = runRef.current;
    const selected = selectionRef.current;
    if (!acceptInput() || !owner || !selected) return;
    if (coordinator.draft) {
      syncDraft(); coordinator.settle(); setFormulaDraft(null);
    }
    const cell = mergeAnchorCell(selected.focus, merges());
    const grid = sourceRef.current?.painted?.displayList.grid;
    editorRectRef.current = grid ? cellRect(grid, cell.row, cell.col) : null;
    const draft: WorkerCellDraft = {
      ...draftFor('cell', cell, seed ?? owner.peer?.cell(activeRef.current, cell.row, cell.col).input ?? ''),
      prefill: seed === undefined && !owner.peer,
      modified: seed !== undefined,
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

  const commitDraft = (source: InputDraft['source'], move?: Direction, queued = false) => {
    if (!queued && deferNavigation(() => commitDraft(source, move, true))) return;
    if (!acceptInput() || composition.current) return;
    if (!queued) syncDraft();
    const draft = coordinator.draft;
    if (draft?.source !== source) return;
    if (!coordinator.submit(draft)) return;
    if ((draft as WorkerCellDraft).modified !== false && !unpreviewedOps.size) {
      previewState.sequence = Infinity;
      heldPreview.current = null;
      if (previewCanvasRef.current) previewCanvasRef.current.dataset.previewReady = 'false';
      setPreviewDraft(draft);
    }
    if (source === 'cell') { suppressBlur.current = true; setEditing(null); }
    else setFormulaDraft(null);
    const base = selectionAt({ row: draft.row, col: draft.col });
    if (move && !runRef.current?.peer) {
      void selectionNavigation.select(base, navigationOptions());
      selectionNavigation.enqueue({ key: move === 'down' || move === 'up' ? 'Enter' : 'Tab',
        shiftKey: move === 'up' || move === 'left' }, navigationOptions());
    } else setSelection(move ? moveFocus(base, move, { limits: limits(), hidden: hidden(), merges: merges() }) : base);
    focus();
  };
  const cancelDraft = () => {
    if (deferNavigation(cancelDraft)) return;
    coordinator.setDraft(null); suppressBlur.current = true; setEditing(null); setFormulaDraft(null); focus();
  };

  const clipboard = (kind: 'copy' | 'cut' | 'paste') => {
    if (beforeNavigationInput(() => clipboard(kind))) return;
    const owner = runRef.current;
    const raw = selectionRef.current;
    const selected = raw && anchorSelection(raw, merges());
    if (!owner?.current || owner.failure || !selected || kind !== 'copy' && !acceptInput()) return;
    if (selected !== raw) setSelection(selected);
    const sheet = activeRef.current;
    const range = normalizeRange(selected);
    const usedRange = () => run!.peer?.selectionData(sheet, 0, 0, 'usedRange').usedRange;
    let clipboardWrite: { resolve(payload: Blob): void; reject(error: unknown): void; done: Promise<boolean> } | undefined;
    if (kind !== 'paste') {
      if (typeof ClipboardItem !== 'function' || typeof navigator.clipboard.write !== 'function') {
        if (kind === 'cut') { reportInputError(new WorkerInputRefusal('Cut requires clipboard write access')); return; }
      } else {
        let resolve!: (payload: Blob) => void;
        let reject!: (error: unknown) => void;
        const payload = new Promise<Blob>((done, failed) => { resolve = done; reject = failed; });
        void payload.catch(() => {});
        try {
          const done = navigator.clipboard.write([new ClipboardItem({ 'text/plain': payload })]).then(
            () => true, (error) => { reject(error); throw new WorkerInputRefusal(error instanceof Error ? error.message : 'Clipboard write failed', kind === 'cut'); }
          );
          void done.catch(() => {});
          clipboardWrite = { resolve, reject, done };
        } catch (error) { reject(error); reportInputError(new WorkerInputRefusal(error instanceof Error ? error.message : 'Clipboard write failed')); return; }
      }
    }
    const captureClipboard = () => kind === 'paste' ? navigator.clipboard.readText() : '';
    let mutation: Extract<WorkbookReplayOp, { method: 'editCells' }> | undefined;
    let discardPreview: (() => Promise<void>) | undefined;
    let clipboardRefusal: unknown;
    void clipboardWrite?.done.catch((error) => {
      clipboardRefusal = error;
      if (kind === 'cut' && discardPreview) void discardPreview().catch(reportInputError);
      reportInputError(error);
    });
    const prepareMutation = async (text: string, queued = false) => {
      if (kind === 'cut' && reachesSheetEdge(range) && !queued) return;
      if (kind !== 'copy' && !mutation) {
        const data = kind === 'cut' && reachesSheetEdge(range) ?
          await (run!.peer ?? run!.session.call).selectionData?.(sheet, 0, 0, 'usedRange') : undefined;
        const edits = kind === 'paste' ? [] : clearedCells(range, () => data?.usedRange);
        if (!edits) throw new WorkerInputRefusal(t('editor.clearExceedsLimit'));
        if (kind === 'cut' && edits.length === 0) return;
        if (kind === 'paste') {
          fromTsv(text).forEach((row, dr) => row.forEach((input, dc) =>
            edits.push({ row: range.top + dr, col: range.left + dc, input })));
        }
        mutation = { method: 'editCells', args: [sheet, edits] };
      }
      if (mutation && !queued) discardPreview = await previewEdits(sheet, mutation.args[1], mutation);
      if (clipboardRefusal) { await discardPreview?.(); throw clipboardRefusal; }
    };
    void coordinator.clipboard(captureClipboard, async (text, _, markApplied) => {
      if (kind !== 'copy' && !mutation) await prepareMutation(text, true);
      if (kind !== 'paste') {
        const copied = toTsv(owner.peer!.rangeCells(sheet, rangeA1(copiedRange(range, usedRange))));
        let written: boolean;
        if (clipboardWrite) {
          clipboardWrite.resolve(new Blob([copied], { type: 'text/plain' }));
          written = await clipboardWrite.done;
        } else {
          try { await navigator.clipboard.writeText(copied); written = true; }
          catch (error) { throw new WorkerInputRefusal(error instanceof Error ? error.message : 'Clipboard write failed'); }
        }
        if (!owner.current) throw new XlsxCommandAdmissionError('document-replaced');
        markApplied.check();
        if (!written) throw new WorkerInputRefusal('Clipboard write was refused', kind === 'cut');
        if (kind === 'copy') return;
      }
      if (!mutation) return;
      const op = mutation;
      const edits = op.args[1];
      if (edits.length === 0) return;
      markApplied.check();
      const operations = workbookEditPeerOperations(run!.recovering ? operationPeer.current! : run!.editPeer!);
      const result = (run!.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult;
      markApplied();
      apply(result, op);
      if (kind === 'paste' && selectionRef.current === selected && activeRef.current === sheet) {
        setSelection({ anchor: { row: range.top, col: range.left }, focus: {
          row: Math.max(...edits.map((edit) => edit.row)), col: Math.max(...edits.map((edit) => edit.col)),
        } });
      }
    }, { sheet, target: JSON.stringify(range) }, kind !== 'copy' ? prepareMutation : undefined, kind !== 'copy').catch(async (error) => {
      clipboardWrite?.reject(error);
      if (inputRefusal(error)) await discardPreview?.();
      reportInputError(error);
    });
  };

  const clearCells = () => {
    if (!acceptInput() || !selectionRef.current) return;
    const sheet = activeRef.current;
    const range = normalizeRange(selectionRef.current);
    let op: WorkbookReplayOp | undefined;
    let discardPreview: (() => Promise<void>) | undefined;
    const prepareClear = async () => {
      const data = reachesSheetEdge(range) ?
        await (run!.peer ?? run!.session.call).selectionData?.(sheet, 0, 0, 'usedRange') : undefined;
      const edits = clearedCells(range, () => data?.usedRange);
      if (!edits) throw new WorkerInputRefusal(t('editor.clearExceedsLimit'));
      if (edits.length === 0) return;
      op = { method: 'editCells', args: [sheet, edits] };
      if (!reachesSheetEdge(range)) discardPreview = await previewEdits(sheet, edits, op);
    };
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      if (!op) await prepareClear();
      if (!op) return;
      markApplied.check();
      const operations = workbookEditPeerOperations(run!.recovering ? operationPeer.current! : run!.editPeer!);
      const result = (run!.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult;
      markApplied(); apply(result, op);
    }, { kind: 'input', target: { sheet, target: JSON.stringify(range) },
      prepare: async () => { if (!reachesSheetEdge(range)) await prepareClear(); } }).catch(async (error) => {
      if (inputRefusal(error)) await discardPreview?.();
      reportInputError(error);
    });
  };

  const navigate = (sheet: number, next: Selection) => {
    const owner = run;
    if (!owner?.current || owner.failure) return;
    const enqueue = (target: Selection) => {
      const predicted = !owner.peer;
      if (predicted) {
        navigationRef.current += 1;
        setActive(sheet);
        setRawSelection(target);
      }
      void coordinator.runAfterPendingInput(async (_, markApplied) => {
        markApplied.check();
        owner.editPeer!.setActiveSheet(sheet);
        markApplied();
        if (!predicted) place(sheet, target);
        else if (activeRef.current === sheet) place(sheet, selectionRef.current);
      }, { kind: 'input', target: { sheet, target: JSON.stringify(target) } }).catch(reportInputError);
    };
    const target = structuredClone(next);
    if (owner.peer) enqueue(target);
    else void selectionNavigation.select(target, { ...navigationOptions(), apply: (resolved) => {
      enqueue(resolved); return true;
    } }, sheet, true);
  };

  const moveChart = (sheet: number, id: string, dx: number, dy: number, retained = false) => {
    const owner = run;
    if (!owner?.current || !retained && !acceptInput() || dx === 0 && dy === 0) return;
    const operation = async (_: unknown, markApplied: WorkerInputLease) => {
      await boundary();
      markApplied.check();
      const result = owner.editPeer!.moveChart(sheet, id, dx, dy);
      markApplied(); apply(result);
      setChartPreview((preview) => preview?.kind === 'move' ? null : preview);
    };
    const options = { kind: 'chart' as const, target: { sheet, target: id }, input: { dx, dy } };
    void (retained ? coordinator.input(operation, options) :
      coordinator.runAfterPendingInput(operation, options)).catch(reportInputError);
  };
  const resizeChart = (sheet: number, id: string, edges: ChartResizeEdges) => {
    const owner = run;
    if (!owner?.current || !acceptInput() ||
      edges.left === 0 && edges.top === 0 && edges.right === 0 && edges.bottom === 0) return;
    const operation = async (_: unknown, markApplied: WorkerInputLease) => {
      await boundary();
      markApplied.check();
      const result = owner.editPeer!.resizeChart(sheet, id, edges);
      markApplied(); apply(result); setChartPreview(null);
    };
    const options = { kind: 'chart' as const, target: { sheet, target: id }, input: edges };
    void coordinator.runAfterPendingInput(operation, options).catch(reportInputError);
  };
  const { start: startChartResize, cancel: cancelChartResize, nudge: nudgeChartResize,
    active: chartResizeActive } = useChartResizeGesture({
    zoom: () => sourceRef.current?.painted?.request.zoom ?? zoomRef.current,
    preview: setChartPreview,
    commit: (_handle, edges) => {
      const chart = chartRef.current;
      if (chart) resizeChart(activeRef.current, chart.id, edges);
    },
  });
  const flushNudge = (retained = false) => {
    const pending = nudge.current;
    nudge.current = null;
    if (pending) moveChart(pending.sheet, pending.id, pending.dx, pending.dy, retained);
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
  };
  flushNudgeRef.current = flushNudge;
  if (run) run.beforeRetire = () => {
    coordinator.retire();
    retirement.resolve();
    cancelChartResize();
    sourceRef.current?.dispose();
    syncDraft();
    composition.current?.settle(true);
    composition.current = null;
    flushNudge(true);
  };
  const moveChartRef = useRef(moveChart);
  moveChartRef.current = moveChart;
  useEffect(() => {
    const stop = (event: globalThis.MouseEvent) => {
      dragging.current = false;
      const drag = chartDrag.current;
      chartDrag.current = null;
      if (drag && event.button === 0) moveChartRef.current(drag.sheet, drag.id,
        (event.clientX - drag.clientX) / drag.zoom, (event.clientY - drag.clientY) / drag.zoom);
      else setChartPreview(null);
    };
    const cancel = () => { dragging.current = false; chartDrag.current = null;
      setChartPreview(null); flushNudgeRef.current(); };
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
  const cellAtPointForFill = useCallback((clientX: number, clientY: number): CellAddr | null => {
    const painted = sourceRef.current?.painted;
    const canvas = canvasRef.current;
    if (!painted || painted.request.sheet !== activeRef.current || !canvas || !painted.geometry) return null;
    const rect = canvas.getBoundingClientRect();
    return cellAtPoint(painted.geometry, (clientX - rect.left) / painted.request.zoom,
      (clientY - rect.top) / painted.request.zoom);
  }, []);
  const fillPreviewRect = useCallback((range: CellRange) => {
    const painted = sourceRef.current?.painted;
    if (!painted?.displayList.grid) return null;
    const rect = rangeRect(painted.displayList.grid, range);
    return rect ? scaledRect(rect, painted.request.zoom) : null;
  }, []);
  const fillSelection = useCallback((source: CellRange, destination: CellRange) => {
    const owner = runRef.current;
    if (!acceptInput() || !owner?.ready || !owner.peer || !owner.editPeer) return;
    const sheet = activeRef.current;
    const from = address({ row: source.top, col: source.left });
    const sourceEnd = address({ row: source.bottom, col: source.right });
    const end = address({ row: destination.bottom, col: destination.right });
    const sourceA1 = from === sourceEnd ? from : `${from}:${sourceEnd}`;
    const destinationA1 = `${from}:${end}`;
    const op: WorkbookReplayOp = { method: 'fillRange', args: [sheet, sourceA1, destinationA1] };
    void coordinator.runAfterPendingInput((_, markApplied) => {
      const current = runRef.current;
      const peer = current?.recovering ? operationPeer.current : current?.editPeer;
      if (!current?.current || current.retiring || current.failure || !peer || !acceptInput()) return;
      markApplied.check();
      try {
        const operations = workbookEditPeerOperations(peer);
        const result = (current.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult;
        markApplied();
        apply(result, op);
      } catch (error) {
        setRefusal(error instanceof Error ? error.message : String(error));
      }
    }, { kind: 'input', target: { sheet, target: destinationA1 } }).catch(reportInputError);
  }, [acceptInput, apply, coordinator, operationPeer, reportInputError]);
  const onMouseDown = (event: MouseEvent) => {
    if (fromPlugin(event) || runRef.current?.failure || editorInputRef.current?.contains(event.target as Node)) return;
    const { cell, chart } = point(event);
    const dismissing = editingRef.current !== null;
    if (dismissing) commitDraft('cell');
    else if (coordinator.draft?.source === 'formula') commitDraft('formula');
    flushNudgeRef.current();
    chartDrag.current = null;
    setChartPreview(null);
    clickStart.current = null;
    if (chart) {
      setSelectedChart({ id: chart.id, movable: chart.movable, resizable: chart.resizable,
        placeholder: chart.placeholder ?? false });
      if (acceptInput() && chart.movable && event.button === 0) {
        chartDrag.current = { id: chart.id, sheet: activeRef.current,
          zoom: sourceRef.current?.painted?.request.zoom ?? zoomRef.current,
          clientX: event.clientX, clientY: event.clientY };
      }
      focus(); event.preventDefault(); return;
    }
    setSelectedChart(null);
    if (!cell || !selectionRef.current) return;
    setSelection(event.shiftKey ? extendTo(selectionRef.current, cell, limits()) : selectionAt(mergeAnchorCell(cell, merges())));
    clickStart.current = dismissing ? null : cell;
    dragging.current = true;
    focus();
  };
  const onMouseMove = (event: MouseEvent<HTMLDivElement>) => {
    if (fromPlugin(event) || runRef.current?.failure) return;
    const drag = chartDrag.current;
    if (drag) {
      if (!event.buttons) { chartDrag.current = null; setChartPreview(null); return; }
      setChartPreview({ kind: 'move', x: event.clientX - drag.clientX, y: event.clientY - drag.clientY });
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
  const followLink = (row: number, col: number) => {
    const frame = sourceRef.current?.painted?.displayList;
    if (!frame || !runRef.current || runRef.current.failure) return;
    const hyperlink = hyperlinkAtCell(frame, row, col);
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
  const onClick = (event: MouseEvent) => {
    if (fromPlugin(event) || runRef.current?.failure) return;
    const { cell } = point(event);
    const start = clickStart.current;
    clickStart.current = null;
    const frame = sourceRef.current?.painted?.displayList;
    if (!cell || !start || !frame || cell.row !== start.row || cell.col !== start.col) return;
    followLink(cell.row, cell.col);
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (fromPlugin(event) || runRef.current?.failure || !viewRef.current || !selectionRef.current ||
      editingRef.current || commands.claimsKey(event)) return;
    const mod = event.metaKey || event.ctrlKey;
    const chart = chartRef.current;
    if (chart) {
      if (event.key === 'Escape') {
        chartDrag.current = null; cancelChartResize(); nudge.current = null;
        setChartPreview(null); setSelectedChart(null);
        clearTimeout(nudgeTimer.current ?? undefined);
        nudgeTimer.current = null; event.preventDefault(); return;
      }
      const delta = !mod && chartKeys[event.key];
      if (delta) {
        if (acceptInput() && chart.movable) {
          const step = event.shiftKey ? 10 : 1;
          const paintedZoom = sourceRef.current?.painted?.request.zoom ?? zoomRef.current;
          const base = nudge.current ?? { id: chart.id, sheet: activeRef.current, dx: 0, dy: 0 };
          nudge.current = { ...base, dx: base.dx + delta[0] * step, dy: base.dy + delta[1] * step };
          setChartPreview({ kind: 'move', x: nudge.current.dx * paintedZoom, y: nudge.current.dy * paintedZoom });
          if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
          nudgeTimer.current = setTimeout(() => flushNudgeRef.current(), 250);
        }
        event.preventDefault();
      }
      return;
    }
    const pressed = commands.clipboardKey(event);
    if (pressed) {
      const kind = pressed === 'pastePlainText' ? 'paste' : pressed;
      if (!deferNavigation(() => clipboard(kind))) clipboard(kind);
      event.preventDefault(); return;
    }
    const request = selectionDataRequest(event);
    const key = { key: event.key, ctrlKey: event.ctrlKey, metaKey: event.metaKey, altKey: event.altKey, shiftKey: event.shiftKey };
    const movement = request || selectionKeyReducer(selectionRef.current, key, limits()).type === 'move';
    const act = () => {
      if (!selectionRef.current) return;
      if (editingRef.current) {
        if (key.key === 'Enter' || key.key === 'Tab') commitDraft('cell', key.key === 'Enter'
          ? key.shiftKey ? 'up' : 'down' : key.shiftKey ? 'left' : 'right', true);
        else if (key.key === 'Escape') cancelDraft();
        else if (key.key.length === 1 && !key.ctrlKey && !key.metaKey && !key.altKey) {
          const draft = { ...editingRef.current, value: editingRef.current.value + key.key };
          if (editorInputRef.current) editorInputRef.current.value = draft.value;
          coordinator.setDraft(draft); setEditing(draft);
        }
        return;
      }
      if (!runRef.current?.peer && movement) {
        selectionNavigation.enqueue(key, navigationOptions());
        return;
      }
      const selected = selectionRef.current;
      const data = request ? runRef.current?.peer?.selectionData(activeRef.current, selected.focus.row,
        selected.focus.col, request === 'region' ? undefined : request) : undefined;
      const action = selectionKeyReducer(selected, key, limits(), data, hidden(), merges());
      if (action.type === 'move') { setSelection(action.selection); revealSelection(activeRef.current, selectionRef.current!); }
      else if (action.type === 'startEdit') openEditor(action.initialInput);
      else if (action.type === 'clear') clearCells();
    };
    if (!runRef.current?.peer && movement) {
      event.preventDefault();
      if (!selectionNavigation.defer(navigationContext(), act, true)) selectionNavigation.enqueue(key, navigationOptions());
      return;
    }
    const action = selectionKeyReducer(selectionRef.current, key, limits(), undefined, hidden(), merges());
    if (action.type !== 'none' || request) {
      event.preventDefault();
      if (!deferNavigation(act)) act();
    }
  };

  const reveal = (peer: WorkbookHandle, sheet: number, row: number, col: number, align: XlsxRevealAlignment,
    live: () => boolean) => {
    const owner = runRef.current;
    const scroll = scrollRef.current;
    if (!owner?.current || !live() || !scroll || activeRef.current !== sheet) return;
    const info = peer.sheetInfo();
    const at = peer.cellPosition(sheet, row, col);
    const rect = peer.cellRect(sheet, row, col);
    const zoom = zoomRef.current;
    if (viewRef.current) setView(extendScrollExtent(scroll, viewRef.current, rect, zoom));
    const frozen = frozenGridExtent(sourceRef.current?.painted?.displayList.grid,
      info.frozenRows, info.frozenCols);
    const x = revealAxis(at.x, rect.w, col < info.frozenCols, scroll.scrollLeft / zoom,
      scroll.clientWidth / zoom - frozen.width, align);
    const y = revealAxis(at.y, rect.h, row < info.frozenRows, scroll.scrollTop / zoom,
      scroll.clientHeight / zoom - frozen.height, align);
    if (x !== null) scroll.scrollLeft = x * zoom;
    if (y !== null) scroll.scrollTop = y * zoom;
    if (!owner.recovering) sourceRef.current?.schedule();
  };
  const revealSelection = (sheet: number, selected: Selection, previousZoom?: number) => {
    const owner = runRef.current;
    if (!owner?.current) return;
    const live = () => owner.current && runRef.current === owner && activeRef.current === sheet &&
      selectionRef.current === selected;
    if (owner.peer) {
      const scroll = scrollRef.current;
      if (previousZoom !== undefined && (!scroll ||
        !cellInViewport(scroll, owner.peer.cellRect(sheet, selected.focus.row, selected.focus.col), previousZoom))) return;
      reveal(owner.peer, sheet, selected.focus.row, selected.focus.col, 'nearest', live);
      return;
    }
    const apply = (rect: Rect) => {
      const scroll = scrollRef.current;
      const view = viewRef.current;
      if (!live() || !scroll || !view || view.sheet !== sheet) return;
      if (previousZoom !== undefined && !cellInViewport(scroll, rect, previousZoom)) return;
      const scale = zoomRef.current;
      if (rect.x + rect.w > view.contentWidth || rect.y + rect.h > view.contentHeight) {
        setView(extendScrollExtent(scroll, view, rect, scale));
      }
      const target = nearestCellScroll(rect, selected.focus, view, {
        x: scroll.scrollLeft / scale, y: scroll.scrollTop / scale,
        width: scroll.clientWidth / scale, height: scroll.clientHeight / scale,
      });
      const moved = scroll.scrollLeft !== target.x * scale || scroll.scrollTop !== target.y * scale;
      if (moved) {
        scroll.scrollLeft = target.x * scale;
        scroll.scrollTop = target.y * scale;
        sourceRef.current?.schedule();
      }
    };
    const painted = sourceRef.current?.painted;
    const view = viewRef.current;
    const rect = painted?.request.sheet === sheet && view
      ? laidOutCellRect(painted.displayList, painted.request.viewport, view, selected.focus) : null;
    if (rect) apply(rect);
    else void owner.session.call.cellGeometry(sheet, selected.focus.row, selected.focus.col)
      .then((geometry) => apply(geometry.rect)).catch(reportInputError);
  };
  const revealedZoom = useRef(zoom);
  useLayoutEffect(() => {
    if (revealedZoom.current === zoom) return;
    const previousZoom = revealedZoom.current;
    revealedZoom.current = zoom;
    const selected = selectionRef.current;
    if (selected && !chartRef.current) revealSelection(activeRef.current, selected, previousZoom);
  }, [zoom]);
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
          markApplied.check();
          const value = operation();
          if (runRef.current?.recovering && value !== null && typeof value === 'object' &&
            ('refused' in value || 'ok' in value && value.ok === false)) {
            throw new XlsxCommandAdmissionError('input-failed');
          }
          if (value !== null && typeof value === 'object' && ('refused' in value || 'ok' in value && value.ok === false)) {
            markApplied.refuse(new WorkerInputRefusal('Plugin request was refused'));
          } else markApplied();
          return value;
        }, { kind: 'plugin', recover: false });
        return { ok: true, value };
      } catch (error) {
        return { ok: false, code: error instanceof XlsxCommandAdmissionError && error.code === 'document-replaced'
          ? 'document-replaced' : 'input-failed', error };
      }
    },
    async admitEdits(handle, request, authorize, commit) {
      const owner = run;
      const facade = owner?.editPeer;
      if (!owner || !facade || owner.peer !== handle) return {
        ok: false, code: 'document-replaced', error: new XlsxCommandAdmissionError('document-replaced'),
      };
      const op: WorkbookReplayOp = { method: 'applyEdits', args: [structuredClone(request)] };
      const operations = workbookEditPeerOperations(facade);
      const readOnly = pluginAccessRef.current!.readOnlyRefusal(handle);
      try {
        const value = await coordinator.runAfterPendingInput(async (_, applied) => {
          await boundary();
          applied.check();
          const denied = authorize(owner.recovering);
          if (denied) { applied.refuse(new WorkerInputRefusal('Plugin write permission was refused')); return denied; }
          if (readOnly) { applied.refuse(new WorkerInputRefusal(readOnly.failure.message)); return readOnly; }
          const write = () => owner.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op);
          let result: XlsxEditResult;
          try { result = (owner.recovering ? write() : commit(write)) as XlsxEditResult; }
          catch (error) {
            if (!(error instanceof WorkbookRecoveryRefusal)) throw error;
            result = error.result as XlsxEditResult;
          }
          if (!result.ok) {
            applied.refuse(new WorkerInputRefusal(result.failure.message), true);
            return result;
          }
          applied();
          apply(result);
          return result;
        }, { kind: 'plugin', input: op });
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

  const requestSave = useSaveRequest({
    document: () => runRef.current?.current && !runRef.current.retiring && !runRef.current.failure ? runRef.current : null,
    onSaveRequest: props.onSaveRequest,
    fail: reportInputError,
  });
  useWorkerXlsxCommands(commands, {
    requestSave,
    peer: () => run?.peer ?? null, editPeer: () => run?.editPeer ?? null,
    status: () => run?.ready ? 'ready' : run && !run.failure || loading ? 'loading' : 'empty',
    readOnly: () => propsRef.current.readOnly ?? false,
    allowHostProposals: () => propsRef.current.allowHostProposals ?? false, collaborative: () => false,
    mutation: () => mutationRef.current, generation: () => run?.generation ?? 0,
    view: () => ({ sheet: activeRef.current, selection: selectionRef.current, chartSelected: chartRef.current !== null,
      zoom: zoomRef.current, capturedFormat: capturedFormatRef.current, borderStyle: borderStyleRef.current,
      borderColor: borderColorRef.current, proposals: proposalsRef.current, proposalsAvailable: isProposalsAvailable(),
      proposalsPanelOpen: panelRef.current, pngExport: false }),
    coordinator, preview: () => preview(), recovering: () => runRef.current?.recovering ?? false,
    i18n: () => propsRef.current.i18n, translate: t, apply, fail: (error) => run?.fail(error),
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
    afterPaint, focusGrid: focus, refuse: (reason) => setRefusal(reason.message),
  });
  useCommandShortcuts(commands, rootRef, props);

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
      markApplied.check();
      const result = run!.editPeer!.applyFormat(active, a1, format);
      markApplied(); apply(result); focus();
    }, { kind: 'input', target: { sheet: active, target: a1 } }).catch(reportInputError);
  }, [capturedFormat, selection, active, coordinator, boundary, apply, reportInputError, focus]);

  useLayoutEffect(() => {
    if (!props.readOnly) return;
    syncDraft();
    settledDraftRevision.current = draftRevision.current;
    const source = coordinator.draft?.source;
    suppressBlur.current = true;
    suppressFormulaBlur.current = true;
    composition.current?.settle(true);
    composition.current = null;
    setEditing(null);
    setFormulaDraft(null);
    setCapturedFormat(null);
    paintFormatSource.current = null;
    dragging.current = false;
    clickStart.current = null;
    chartDrag.current = null;
    cancelChartResize();
    nudge.current = null;
    setChartPreview(null);
    clearTimeout(nudgeTimer.current ?? undefined);
    nudgeTimer.current = null;
    if (coordinator.draft || coordinator.pending) void coordinator.drain().catch(reportInputError);
    if (source === 'cell') focus();
    suppressFormulaBlur.current = false;
  }, [props.readOnly, coordinator]);

  const startComposition = (source: InputDraft['source']) => {
    if (beforeNavigationInput(() => startComposition(source))) return;
    if (!acceptInput()) return;
    if (source === 'formula' && coordinator.draft?.source !== 'formula' && selectionRef.current) {
      syncDraft(); coordinator.settle(); suppressBlur.current = true; setEditing(null);
      coordinator.setDraft(draftFor(source, mergeAnchorCell(selectionRef.current.focus, merges()),
        formulaInputRef.current?.value ?? ''));
    }
    let settle!: (ended: boolean) => void;
    const done = new Promise<boolean>((resolve) => { settle = resolve; });
    composition.current = { source, done, settle };
  };
  const endComposition = (value?: string) => {
    if (deferNavigation(() => endComposition(value))) return;
    if (!composition.current) return;
    const draft = coordinator.draft;
    if (value !== undefined && draft) {
      const completed = { ...draft, value };
      coordinator.setDraft(completed);
      if (draft.source === 'cell') setEditing(completed); else setFormulaDraft(value);
    } else syncDraft();
    composition.current?.settle(true); composition.current = null;
  };
  const attachCellInput = useCallback((input: HTMLInputElement | null) => {
    if (!input) inputHooks.current?.sync?.();
    editorInputRef.current = input;
    if (input) suppressBlur.current = false;
    if (!input && composition.current?.source === 'cell') {
      composition.current.settle(true); composition.current = null;
    }
  }, [inputHooks]);
  const attachFormulaInput = useCallback((input: HTMLInputElement | null) => {
    if (!input) inputHooks.current?.sync?.();
    formulaInputRef.current = input;
    if (!input && composition.current?.source === 'formula') {
      composition.current.settle(true); composition.current = null;
    }
  }, [inputHooks]);
  useLayoutEffect(() => {
    if (!editing || run?.failure) return;
    const input = editorInputRef.current;
    if (input && document.activeElement !== input) {
      input.focus({ preventScroll: true }); input.setSelectionRange(input.value.length, input.value.length);
    }
  }, [editing !== null]);

  const changeFormulaDraft = (value: string) => {
    const selected = selectionRef.current;
    if (!acceptInput() || !selected || chartRef.current) return;
    let draft = coordinator.draft;
    if (draft && draft.source !== 'formula') {
      syncDraft(); coordinator.settle(); suppressBlur.current = true; setEditing(null);
      draft = null;
    }
    coordinator.setDraft({ ...(draft ?? draftFor('formula', mergeAnchorCell(selected.focus, merges()), value)), value, modified: true,
      revision: ++draftRevision.current } as WorkerCellDraft);
  };
  const formulaBar: FormulaBarBinding = {
    onNavigate: (cell) => { void selectionNavigation.select(selectionAt(cell), navigationOptions()); focus(); },
    a1: selection ? address(selection.focus) : '',
    value: (composition.current?.source === 'formula' ? formulaInputRef.current?.value : undefined) ?? formulaDraft ?? focusedCell?.input ?? '',
    disabled: !view || !selection || Boolean(error), readOnly: !acceptInput() || selectedChart !== null,
    inputRef: attachFormulaInput,
    onChange(value) {
      if (!acceptInput() || !selectionRef.current || chartRef.current) return;
      setFormulaDraft(value);
      if (!beforeNavigationInput(() => changeFormulaDraft(value))) changeFormulaDraft(value);
    },
    onCommit: (move) => commitDraft('formula', move), onCancel: cancelDraft,
    onBlur: () => { if (!suppressFormulaBlur.current) commitDraft('formula'); },
    onCompositionStart: () => startComposition('formula'),
    onCompositionEnd: () => endComposition(formulaInputRef.current?.value),
  };
  const frame = painted?.displayList;
  useEffect(() => {
    if (!selectedChart || !frame || frame.charts?.some((chart) => chart.id === selectedChart.id)) return;
    flushNudgeRef.current(); chartDrag.current = null; cancelChartResize();
    setChartPreview(null); setSelectedChart(null);
  }, [selectedChart, frame]);
  const paintedZoom = painted?.request.zoom ?? zoom;
  const visibleSelection = painted?.request.sheet === active ? selection : null;
  const merged = painted?.mergedRanges ?? [];
  const grid = frame?.grid;
  const selected = visibleSelection ? expandRangeToMergedCells(normalizeRange(visibleSelection), merged, Boolean(props.showSheetHeaders)) : null;
  const focused = visibleSelection ? expandRangeToMergedCells({ top: visibleSelection.focus.row,
    left: visibleSelection.focus.col, bottom: visibleSelection.focus.row, right: visibleSelection.focus.col }, merged) : null;
  const selectionRect = grid && selected ? rangeRect(grid, selected) : null;
  const selectionBottomRight = grid && selected ? cellRect(grid, selected.bottom, selected.right) : null;
  const fillHandleRect = fillHandleAnchor(
    selectionRect ? scaledRect(selectionRect, paintedZoom) : null,
    selectionBottomRight ? scaledRect(selectionBottomRight, paintedZoom) : null
  );
  const focusRect = grid && focused ? rangeRect(grid, focused) : null;
  const editRect = editing ? (grid && editing.sheet === painted?.request.sheet ?
    cellRect(grid, editing.row, editing.col) : null) ?? editorRectRef.current ?? { x: 0, y: 0, w: 96, h: 24 } : null;
  const editColors = editRect ? cellEditorColors(editing?.sheet === painted?.request.sheet ? painted?.displayList : null, editRect) : null;
  const previewRect = previewDraft && grid && previewDraft.sheet === painted?.request.sheet ?
    rangeRect(grid, expandRangeToMergedCells({ top: previewDraft.row, left: previewDraft.col,
      bottom: previewDraft.row, right: previewDraft.col }, merged)) : null;
  const chartRegion = selectedChart
    ? frame?.charts?.find((chart) => chart.id === selectedChart.id) ?? null : null;
  const chartFrozen = painted && view
    ? frozenGridExtent(painted.displayList.grid, view.frozenRows, view.frozenCols) : null;
  const chartSpace = chartRegion && painted && view ? chartResizeSpace(chartRegion, {
    x: painted.request.viewport.x, y: painted.request.viewport.y,
    frozenRows: view.frozenRows, frozenCols: view.frozenCols,
    frozenWidth: chartFrozen?.width, frozenHeight: chartFrozen?.height,
  }) : undefined;
  const shownChart = chartRegion && chartPreview?.kind === 'resize'
    ? chartResizeRect(chartRegion.rect, chartPreview.handle, chartPreview.dx, chartPreview.dy, chartSpace)
    : chartRegion?.rect ?? null;
  const showChartHandles = Boolean(shownChart && selectedChart?.movable && selectedChart.resizable &&
    !selectedChart.placeholder && acceptInput() && chartPreview?.kind !== 'move');
  const scaledChart = shownChart ? scaledRect(shownChart, paintedZoom) : null;
  const sheets = run?.session.state.sheets ?? [];
  const a11yGrid = frame && painted ? buildA11yGrid(frame, visibleSelection, sheets[painted.request.sheet]?.name ?? '', {
    gridLabel: t('a11y.gridLabel'), rowHeaderLabel: t('a11y.rowHeaderLabel'), columnHeaderLabel: t('a11y.columnHeaderLabel'),
    cellLabel: t('a11y.cellLabel'), cellLabelSelected: t('a11y.cellLabelSelected'), emptyCellLabel: t('a11y.emptyCellLabel'),
    emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
  }) : null;
  /**
   * Applies structural ops as one edit through the edit peer, reusing the
   * command bridge's funnel. A queued cell draft is written first: the ops name
   * coordinates the user typed against (a header resize keeps its track), so
   * they must never overtake that input the way a bare `applyOps` call would.
   */
  const applyStructuralOps = (ops: Record<string, unknown>[]): boolean => {
    const owner = runRef.current;
    if (!owner?.ready || owner.retiring || !acceptInput()) return false;
    const edits = owner.editPeer;
    if (!edits) return false;
    const operation = (markApplied: WorkerInputLease) => {
      markApplied.check();
      let result: EditResult;
      try {
        result = edits.applyOps(ops) as EditResult;
      } catch (error) {
        // The workbook declined the edit and left itself unchanged. Report it as a local
        // refusal before the coordinator treats the engine's message as a fatal failure.
        const structural = engineStructuralRefusal(error, 'applyOps');
        if (structural) throw new WorkerInputRefusal(t(REFUSAL_CAUSE_KEYS[structural.cause]));
        throw error;
      }
      if (result.applied) { markApplied(); apply(result); }
      return result.applied;
    };
    const report = (error: unknown) => {
      const structural = engineStructuralRefusal(error, 'applyOps');
      if (structural) { setRefusal(t(REFUSAL_CAUSE_KEYS[structural.cause])); return false; }
      const local = inputRefusal(error);
      if (local || error instanceof XlsxCommandAdmissionError) {
        setRefusal((local ?? error as Error).message);
        return false;
      }
      owner.fail(error);
      return false;
    };
    try {
      // nothing queued: keep the gesture's synchronous single-edit apply.
      return coordinator.runSync((_, markApplied) => operation(markApplied), { kind: 'host' });
    } catch (error) {
      if (!(error instanceof WorkerInputNotReadyError)) return report(error);
    }
    // a draft or an unacknowledged edit is queued; join the same FIFO the cell
    // edits use, so pending input is written before the op is applied.
    void coordinator.runAfterPendingInput((_, markApplied) => operation(markApplied),
      { kind: 'host' }).catch(report);
    return true;
  };
  /** Applies one structural op as its own edit. */
  const applyStructuralOp = (op: Record<string, unknown>): boolean => applyStructuralOps([op]);
  /**
   * Selects a whole column or row track; `extend` grows the range from the
   * anchor's track to the pressed one, the way a shifted press inside the grid
   * extends. A track spans the sheet's own extent, not the painted window's, so
   * the explicit range is what makes the selection a whole column or row.
   * The anchor is written as the pressed track, never through the cell path's
   * merge snap: with `B1:C1` merged, pressing the C header must anchor at C, or
   * a shifted press on D would grow B:D behind the user's back.
   */
  const selectHeader = (axis: 'col' | 'row', index: number, extend: boolean) => {
    const current = selectionRef.current;
    if (extend && !current) return;
    const anchor = extend && current
      ? current.anchor
      : axis === 'col' ? { row: 0, col: index } : { row: index, col: 0 };
    const track = axis === 'col'
      ? { top: 0, left: Math.min(anchor.col, index), bottom: XLSX_MAX_ROWS - 1, right: Math.max(anchor.col, index) }
      : { top: Math.min(anchor.row, index), left: 0, bottom: Math.max(anchor.row, index), right: XLSX_MAX_COLUMNS - 1 };
    const focus = axis === 'col' ? { row: 0, col: index } : { row: index, col: 0 };
    setRawSelection({ anchor, focus, range: track });
  };
  const endChartGesture = () => { chartDrag.current = null; cancelChartResize(); setChartPreview(null); };
  const onHeaderMouseDown = (axis: 'col' | 'row', index: number, event: MouseEvent) => {
    if (event.button !== 0 || fromPlugin(event) || runRef.current?.failure) return;
    if (editingRef.current !== null) commitDraft('cell');
    else if (coordinator.draft?.source === 'formula') commitDraft('formula');
    flushNudgeRef.current();
    endChartGesture();
    setSelectedChart(null);
    selectHeader(axis, index, event.shiftKey);
    clickStart.current = null;
    focus();
  };
  /**
   * The corner above the row numbers selects the whole sheet. It settles the
   * draft and anchors a raw range that spans the sheet's own extent: the
   * painted window is not the selection.
   */
  const onSelectAll = () => {
    if (!acceptInput()) return;
    if (editingRef.current !== null) commitDraft('cell');
    else if (coordinator.draft?.source === 'formula') commitDraft('formula');
    flushNudgeRef.current();
    endChartGesture();
    setSelectedChart(null);
    setRawSelection({
      anchor: { row: 0, col: 0 },
      focus: { row: 0, col: 0 },
      range: { top: 0, left: 0, bottom: XLSX_MAX_ROWS - 1, right: XLSX_MAX_COLUMNS - 1 },
    });
    clickStart.current = null;
    focus();
  };
  const onHeaderResize = (axis: 'col' | 'row', index: number, size: number,
    run?: { at: number; count: number }) => {
    if (run && ((axis === 'row' && run.count >= XLSX_MAX_ROWS) || (axis === 'col' && run.count >= XLSX_MAX_COLUMNS))) return;
    if (!acceptInput()) return;
    const sheet = activeRef.current;
    // the band reports pixels; the engine stores characters and points.
    if (run) {
      // one batch keeps a whole-run drag a single undo step.
      applyStructuralOps(Array.from({ length: run.count }, (_, offset) => axis === 'col'
        ? { type: 'setColWidth', sheet, col: run.at + offset, width: columnWidthFromPx(size) }
        : { type: 'setRowHeight', sheet, row: run.at + offset, height: rowHeightFromPx(size) }));
      return;
    }
    applyStructuralOp(axis === 'col'
      ? { type: 'setColWidth', sheet, col: index, width: columnWidthFromPx(size) }
      : { type: 'setRowHeight', sheet, row: index, height: rowHeightFromPx(size) });
  };
  // keyed on the session for the same reason the peerless editor is: the first
  // render has no session, and a `[]` memo would pin handlers bound to it.
  const headerHandlers = useMemo<SheetGridHeaderHandlers>(() => ({
    onHeaderMouseDown, onHeaderResize, onSelectAll,
    // a drag from an unselected header extends from the anchor the press set;
    // the press already committed the draft and took focus.
    onHeaderDragSelect: (axis, index) => {
      if (runRef.current?.failure) return;
      selectHeader(axis, index, true);
    },
  }), [run]);
  const headerMenu = useContextMenu({
    // the bands are siblings of the scroller, so their contextmenu never
    // reaches a listener bound to scrollRef; the editor root sees both.
    surface: rootRef,
    resolve: (request) => {
      const owner = runRef.current;
      if (!owner?.ready || owner.failure || owner.retiring || inPluginChrome(request.target)) return null;
      const band = request.target instanceof Element
        ? request.target.closest<HTMLElement>('[data-testid="xlsx-column-headers"], [data-testid="xlsx-row-headers"]')
        : null;
      if (!band) return null;
      const axis = band.dataset.testid === 'xlsx-column-headers' ? 'col' : 'row';
      const track = request.target instanceof Element
        ? request.target.closest<HTMLElement>('[data-col], [data-row]') : null;
      const index = Number(axis === 'col' ? track?.dataset.col : track?.dataset.row);
      if (!track || !Number.isInteger(index)) return null;
      // select the pressed track so the visible selection matches the menu's target.
      selectHeader(axis, index, false);
      const sheet = activeRef.current;
      const disabled = !acceptInput();
      // the legacy command store has no insert/delete row or column commands,
      // so the entries go straight to the structural op helper.
      const structural = (op: Record<string, unknown>) => () => { applyStructuralOp(op); };
      const entries: OverflowMenuEntry[] = axis === 'col' ? [
        { kind: 'item', id: 'insertColumnsLeft', label: t('commands.insertColumnsLeft'), disabled,
          onSelect: structural({ type: 'insertCols', sheet, at: index, count: 1 }) },
        { kind: 'item', id: 'insertColumnsRight', label: t('commands.insertColumnsRight'), disabled,
          onSelect: structural({ type: 'insertCols', sheet, at: index + 1, count: 1 }) },
        { kind: 'item', id: 'deleteColumns', label: t('commands.deleteColumns'), disabled,
          onSelect: structural({ type: 'deleteCols', sheet, at: index, count: 1 }) },
      ] : [
        { kind: 'item', id: 'insertRowsAbove', label: t('commands.insertRowsAbove'), disabled,
          onSelect: structural({ type: 'insertRows', sheet, at: index, count: 1 }) },
        { kind: 'item', id: 'insertRowsBelow', label: t('commands.insertRowsBelow'), disabled,
          onSelect: structural({ type: 'insertRows', sheet, at: index + 1, count: 1 }) },
        { kind: 'item', id: 'deleteRows', label: t('commands.deleteRows'), disabled,
          onSelect: structural({ type: 'deleteRows', sheet, at: index, count: 1 }) },
      ];
      // freezing and unhiding read the frame the bands painted; both go
      // through the same structural helper as insert/delete, so one gesture is
      // still one undo step and one applyOps call.
      const paintedFrame = sourceRef.current?.painted;
      const view = paintedFrame?.request.sheet === sheet ? paintedFrame.displayList.grid : null;
      if (view) {
        const freeze: XlsxContextFreezeTarget = {
          rows: view.frozenRows ?? 0,
          cols: view.frozenCols ?? 0,
          index,
          apply: (rows: number, cols: number) => {
            // the menu can outlive a navigation; only the painted sheet is safe.
            if (sourceRef.current?.painted?.request.sheet !== activeRef.current) return;
            applyStructuralOp({ type: 'setFreezePane', sheet: activeRef.current,
              pane: rows === 0 && cols === 0 ? null : { rows, cols, top_left: { row: rows, col: cols } } });
          },
        };
        // menus.ts owns the view section: freezing to the pressed header and
        // clearing the split.
        const section = xlsxContextMenuSections(axis === 'col' ? 'columnHeader' : 'rowHeader',
          commands.store, t, freeze).find((entry) => entry.id === 'view');
        if (section) {
          entries.push(...section.entries.map((entry) => entry.kind === 'item' ? { ...entry, disabled } : entry));
        }
      }
      return {
        label: t(axis === 'col' ? 'editor.columnHeaderMenuLabel' : 'editor.rowHeaderMenuLabel'),
        testId: axis === 'col' ? 'xlsx-column-header-menu' : 'xlsx-row-header-menu',
        entries: () => entries,
      };
    },
  });

  const headers = Boolean(props.showSheetHeaders);
  const outline = (rect: NonNullable<typeof selectionRect>, border: string, background?: string): CSSProperties => {
    const scaled = scaledRect(rect, paintedZoom);
    return { position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      boxSizing: 'border-box', border, background };
  };
  // With headers the chart box keeps an outline: it is a movable frame rather
  // than a cell edge and has no snapped interior to straddle.
  const edgeOutline = (rect: NonNullable<typeof selectionRect>, color: string, width: 1 | 2 = 2): CSSProperties => {
    const scaled = scaledRect(rect, paintedZoom);
    return { position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      outline: `${width}px solid ${color}`, outlineOffset: -width / 2 };
  };
  // Cell selections ride the same snapped pixel grid as the canvas gridlines:
  // the range ring is one pixel on the snapped rect and the active-cell ring is
  // two pixels straddling its boundaries, never an outline that rasterizes
  // asymmetrically at a fractional zoom.
  const marquee = (rect: NonNullable<typeof selectionRect>, width = MARQUEE_WIDTH): CSSProperties => {
    const box = marqueeBox(scaledRect(rect, paintedZoom), width);
    return { position: 'absolute', left: box.left, top: box.top, width: box.width, height: box.height,
      boxShadow: `inset 0 0 0 ${width}px ${MARQUEE_COLOR}` };
  };
  const fills = headers && selectionRect && !selectedChart
    ? marqueeFills(scaledRect(selectionRect, paintedZoom), focusRect ? scaledRect(focusRect, paintedZoom) : null) : [];
  const defaultToolbar = <EditorToolbar mode="commands">
    <EditorToolbar.Toolbar />
    <div style={{ display: 'flex', padding: '4px 8px' }}>
      <EditorToolbar.FormulaBar />
      {isProposalsAvailable() && <ToolbarCommandButton id="proposalsPanel" />}
    </div>
  </EditorToolbar>;
  const reviewToolbar = props.allowHostProposals && isProposalsAvailable() ? <ToolbarCommandButton id="proposalsPanel" /> : null;
  const toolbar = props.showToolbar === false ? null : props.toolbar === undefined ? props.readOnly ? reviewToolbar : defaultToolbar : props.toolbar;
  const renderDock = (placement: DockPlacement) => pluginHost.managed ? <PluginDock host={pluginHost.host}
    placement={placement} activations={pluginHost.activations.filter((activation) => activation.plugin.panel?.placement === placement)}
    available={dockArea} /> : null;

  return <XlsxCommandContext.Provider value={commands.store}>
    <EditorChromeContext.Provider value={true}>
      <FormulaBarContext.Provider value={formulaBar}>
        <div dir="ltr" ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
          style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%', height: '100%',
            minWidth: 0, color: '#202124', background: '#ffffff', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }}>
          {toolbar != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbar}</div>}
          <div ref={dockAreaRef} data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minWidth: 0, minHeight: 0 }}>
            {renderDock('left')}
            <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
              {isProposalsAvailable() && proposalsPanelOpen && <ProposalsPanel proposals={proposals} staleFor={staleFor}
                style={{ top: 4, right: 8, width: 'min(320px, calc(100% - 16px))', maxHeight: 'min(420px, calc(100% - 12px))' }} />}
              {pluginHost.rowGutter && <PluginRowGutter host={pluginHost.host} index={pluginHost.rowGutter}
                painted={pluginHost.painted} zoom={zoom} top={props.showSheetHeaders ? COLUMN_HEADER_HEIGHT : 0} />}
              <SheetGrid frame={frame} zoom={paintedZoom} leadingInset={pluginHost.rowGutter ? ROW_GUTTER_WIDTH * zoom : 0} showSheetHeaders={headers} headerHandlers={headerHandlers} selection={selection}
                sheet={painted?.request.sheet}
                selectAllLabel={t('commands.selectAll')}>
              <div ref={scrollRef} data-testid="xlsx-scroll" tabIndex={0} onKeyDown={onKeyDown}
                onMouseDown={onMouseDown} onMouseMove={onMouseMove} onClick={onClick}
                onDoubleClick={(event) => {
                  if (fromPlugin(event) || editingRef.current || point(event).chart) return;
                  const cell = point(event).cell;
                  if (cell && frame && hyperlinkAtCell(frame, cell.row, cell.col)) return;
                  openEditor();
                }}
                onMouseLeave={(event) => { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }}
                style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none',
                  ...(pluginHost.rowGutter && !props.showSheetHeaders ? { marginLeft: ROW_GUTTER_WIDTH * zoom } : {}) }}>
                <div style={{ position: 'absolute', top: 0, left: 0, width: view ? view.contentWidth * zoom : '100%',
                  height: view ? view.contentHeight * zoom : '100%' }} />
                <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
                  <canvas ref={canvasRef} style={{ display: 'block', position: 'absolute', top: 0, left: 0 }} />
                  {previewDraft && (run?.ready ? <div data-testid="xlsx-commit-preview" aria-hidden
                    style={{ ...outline(previewRect ?? { x: 0, y: 0, w: 96, h: 24 }, 'none', '#fff'),
                      color: '#000', padding: `0 ${8 * paintedZoom}px`, overflow: 'hidden', whiteSpace: 'pre',
                      display: previewDraft.sheet === active && (!grid || previewRect) ? 'block' : 'none',
                      font: `${13 * paintedZoom}px system-ui, sans-serif`,
                      lineHeight: `${(previewRect?.h ?? 24) * paintedZoom}px`, pointerEvents: 'none' }}>{previewDraft.value}</div> :
                    <canvas ref={previewCanvasRef} data-testid="xlsx-commit-preview" aria-hidden
                      style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }}>{previewDraft.value}</canvas>)}
                  {pluginHost.managed && <PluginOverlays host={pluginHost.host} activations={pluginHost.activations}
                    layerRef={pluginHost.overlayLayerRef} width={(frame?.width ?? 0) * paintedZoom} height={(frame?.height ?? 0) * paintedZoom} />}
                  <div data-testid="xlsx-overlay-host" style={{ position: 'absolute', top: 0, left: 0, width: 0, height: 0, pointerEvents: 'none' }}>
                    {fills.map((fill, index) => <div key={index} data-testid="xlsx-selection-fill"
                      style={{ position: 'absolute', left: fill.x, top: fill.y, width: fill.w, height: fill.h,
                        background: SELECTION_FILL }} />)}
                    {selectionRect && !selectedChart && <div data-testid="xlsx-selection"
                      style={headers ? marquee(selectionRect) : outline(selectionRect, `1px solid ${BRAND}`, 'rgba(33, 115, 70, 0.12)')} />}
                    {props.fillHandle && !props.readOnly && !editing && !coordinator.draft && !selectedChart && acceptInput() &&
                      fillHandleRect && selected && !coversWholeTrack(selected, limits()) && (
                      <FillHandle rect={fillHandleRect} source={selected}
                        rectForRange={fillPreviewRect} cellAt={cellAtPointForFill} onFill={fillSelection}
                        ariaLabel={t('a11y.fillHandleLabel')} />
                    )}
                    {focusRect && !selectedChart && !editing && <div data-testid={headers ? 'xlsx-active-cell' : undefined}
                      style={headers ? marquee(focusRect, ACTIVE_CELL_WIDTH) : outline(focusRect, `2px solid ${BRAND}`)} />}
                    {shownChart && <div data-testid="xlsx-chart-selection" data-chart-id={selectedChart?.id} aria-hidden
                      style={{ ...(headers ? edgeOutline(shownChart, BRAND) : outline(shownChart, `2px solid ${BRAND}`)),
                        transform: `translate(${chartPreview?.kind === 'move' ? chartPreview.x : 0}px, ${chartPreview?.kind === 'move' ? chartPreview.y : 0}px)`,
                        boxShadow: '0 1px 6px rgba(0, 0, 0, 0.25)' }} />}
                    {showChartHandles && scaledChart && <ChartResizeHandles rect={scaledChart}
                      onStart={(handle, event) => { flushNudgeRef.current(); startChartResize(handle, event); }}
                      onNudge={nudgeChartResize} />}
                    {editing && editRect && <input ref={attachCellInput} data-testid="xlsx-cell-editor"
                      value={(composition.current?.source === 'cell' ? editorInputRef.current?.value : undefined) ?? editing.value}
                      disabled={Boolean(error) || props.readOnly}
                      onChange={(event) => {
                        if (!acceptInput()) return;
                        const draft: WorkerCellDraft = { ...editing, value: event.target.value, prefill: false, modified: true,
                          revision: ++draftRevision.current };
                        setEditing(draft); coordinator.setDraft(draft);
                      }}
                      onCompositionStart={() => startComposition('cell')} onCompositionEnd={(event) => endComposition(event.currentTarget.value)}
                      onKeyDown={(event) => {
                        if (!commands.claimsKey(event)) event.stopPropagation();
                        if (event.nativeEvent.isComposing || event.keyCode === 229) return;
                        if (event.key === 'Enter') { commitDraft('cell', event.shiftKey ? 'up' : 'down'); event.preventDefault(); }
                        else if (event.key === 'Tab') { commitDraft('cell', event.shiftKey ? 'left' : 'right'); event.preventDefault(); }
                        else if (event.key === 'Escape') { cancelDraft(); event.preventDefault(); }
                      }}
                      onBlur={() => {
                        if (suppressBlur.current) { suppressBlur.current = false; return; }
                        commitDraft('cell');
                      }}
                      style={{ ...outline(editRect, `2px solid ${BRAND}`, editColors?.background), color: editColors?.color, padding: '0 3px',
                        font: `${13 * paintedZoom}px system-ui, sans-serif`, pointerEvents: 'auto', outline: 'none' }} />}
                  </div>
                </div>
              </div>
              </SheetGrid>
              {renderDock('bottom')}
            </div>
            {renderDock('right')}
          </div>
          {a11yGrid && <>
            <A11yGridTable grid={a11yGrid} style={visuallyHidden} onFollowLink={followLink} />
            {a11yGrid.charts.map((chart, index) => <div key={`${index}:${chart.label}`} style={visuallyHidden} role="img" aria-label={chart.label} />)}
          </>}
          {refusal && <div data-testid="xlsx-input-refusal" role="alert" style={{ position: 'relative', zIndex: 1 }}>{refusal}</div>}
          {headerMenu.element}
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
