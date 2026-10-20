import { createSelectionDataNavigation, sameSelection, selectionDataRequest, selectionHidden } from '../selectionData';
import {
  cellAtPoint, cellRect, chartRegionAtPoint, extendTo, fromTsv, hyperlinkAtCell, isProposalsAvailable,
  moveFocus, normalizeRange, paintDisplayList, parseHyperlinkLocation, safeExternalHyperlink, selectionAt,
  selectionKeyReducer, toTsv,
} from '@betteroffice/xlsx';
import type {
  CellAddr, CellInputEdit, Proposal, Selection, WorkbookSheetView, WorkbookCommandSelection,
  WorkbookSheetSummary,
} from '@betteroffice/xlsx';
import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { KeyboardEvent, MouseEvent, SetStateAction, SyntheticEvent } from 'react';
import { flushSync } from 'react-dom';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { commandForEvent } from '../commands/descriptors';
import { commandReason } from '../commands/evaluate';
import type { XlsxCommandFailureCode } from '../commands/types';
import type { InputSeal } from '../commands/inputCoordinator';
import { inPluginChrome, pluginEventStore } from '../commands/pluginEvents';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { useSaveRequest } from '../commands/useSaveRequest';
import { usePeerlessWorkerXlsxCommands } from '../commands/usePeerlessWorkerXlsxCommands';
import { WorkerInputRefusal } from '../commands/workerInputCoordinator';
import { readWorkerCellInput } from './WorkerCellInputReader';
import {
  createPeerlessWorkerInputCoordinator, PeerlessWorkerCommandRefusal, type PeerlessWorkerCommandInput,
  type PeerlessWorkerCommandTicket, type PeerlessWorkerInputEvent,
} from '../commands/peerlessWorkerInputCoordinator';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import { EditorToolbar } from '../components/EditorToolbar';
import { EditorChromeContext, PeerlessChromeContext } from '../components/EditorToolbarContext';
import type { BorderStyle } from '../components/Toolbar';

import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { useTranslation } from '../i18n';
import type { XlsxRevealAlignment } from '../plugins/createPluginClients';
import { createPeerlessPluginClients } from '../plugins/createPeerlessPluginClients';
import { createPeerlessPluginDocument, createPeerlessPluginGeometry } from '../plugins/createPeerlessPluginDocument';
import type { XlsxPluginDocumentAccess } from '../plugins/createPluginDocumentAccess';
import type { XlsxPluginNavigationFailureCode } from '../plugins/types';
import { PluginOverlays } from '../plugins/PluginOverlays';
import { PluginDock, useDockArea, type DockPlacement } from '../plugins/PluginPanels';
import { useXlsxPluginHost, type XlsxPluginHostBinding } from '../plugins/useXlsxPluginHost';
import { ProposalsPanel } from '../proposals/ProposalsPanel';
import { deriveLimits } from '../viewer/sessionGeometry';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { XlsxAsyncExportBridge } from './createAsyncEditorApi';
import { PeerlessWorkbookSession, type PeerlessSessionWorkbookProps } from './usePeerlessSessionWorkbook';
import { usePeerlessSessionWorkbook } from './usePeerlessSessionWorkbook';
import { PeerlessWorkerPaintSource, type WorkerPaintRequest, type WorkerPaintResult } from './PeerlessWorkerPaintSource';

import { WorkerCellEcho } from './WorkerCellEcho';
import { WorkerDraftController, type WorkerNativeDraft } from './WorkerDraftController';
import { PeerlessWorkerInputs } from './PeerlessWorkerInputs';
import { WorkerPresentationStore } from './WorkerPresentationStore';
import {
  PeerlessSelectionEffect, PeerlessWorkerAccessibility, PeerlessWorkerExtent, PeerlessWorkerOverlay,
  PeerlessWorkerPresence,
} from './PeerlessWorkerSurface';

const BRAND = '#217346';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const chartKeys: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1],
};

function address(cell: CellAddr): string {
  let column = '';
  for (let col = cell.col + 1; col > 0; col = Math.floor((col - 1) / 26)) {
    column = String.fromCharCode(65 + (col - 1) % 26) + column;
  }
  return `${column}${cell.row + 1}`;
}

function commandSelection(target: WorkbookCommandSelection, previous?: Selection | null): Selection | null {
  let start: CellAddr;
  let end: CellAddr;
  if (target.range.kind === 'rowCol') ({ start, end } = target.range);
  else {
    const match = /^\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/i.exec(target.range.a1);
    if (!match) return null;
    const cell = (letters: string, row: string): CellAddr => ({ row: Number(row) - 1,
      col: [...letters.toUpperCase()].reduce((col, letter) => col * 26 + letter.charCodeAt(0) - 64, 0) - 1 });
    start = cell(match[1], match[2]);
    end = cell(match[3] ?? match[1], match[4] ?? match[2]);
  }
  const focus = target.activeCell ?? end;
  if (previous?.range && previous.focus.row === focus.row && previous.focus.col === focus.col &&
    previous.range.top === start.row && previous.range.left === start.col &&
    previous.range.bottom === end.row && previous.range.right === end.col) return previous;
  if ((focus.row !== start.row && focus.row !== end.row) ||
    (focus.col !== start.col && focus.col !== end.col)) {
    return { ...selectionAt(focus), range: { top: start.row, left: start.col, bottom: end.row, right: end.col } };
  }
  return { focus: { ...focus }, anchor: {
    row: focus.row === start.row ? end.row : start.row,
    col: focus.col === start.col ? end.col : start.col,
  } };
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

function samePaintRequest(painted: WorkerPaintResult, request: WorkerPaintRequest): boolean {
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

const PeerlessDefaultToolbar = memo(function PeerlessDefaultToolbar() {
  return <EditorToolbar mode="commands"><EditorToolbar.Toolbar />
    <div style={{ display: 'flex', padding: '4px 8px' }}>
      <ToolbarCommandButton id="save" /><ToolbarCommandButton id="exportPng" /><EditorToolbar.FormulaBar />
      {isProposalsAvailable() && <ToolbarCommandButton id="proposalsPanel" />}
    </div>
  </EditorToolbar>;
});

export function XlsxPeerlessWorkerEditor(props: PeerlessSessionWorkbookProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pluginBindingRef = useRef<XlsxPluginHostBinding | null>(null);
  useLayoutEffect(() => () => pluginBindingRef.current?.host.close('unmounted'), []);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [commands] = useState(createXlsxCommandController);
  const [active, setActive, activeRef] = useSyncedState(0);
  const activeSheetId = useRef<string | undefined>(undefined);
  const [presentation] = useState(() => new WorkerPresentationStore());
  const { setSelection, selectionRef, setView, viewRef, setPainted } = presentation;
  const [catalog, setCatalog, catalogRef] = useSyncedState<WorkbookSheetSummary[]>([]);
  const [zoom, setZoom, zoomRef] = useSyncedState(1);
  const [refusal, setRefusal] = useState<{ code: XlsxCommandFailureCode; detail?: string } | null>(null);
  const [selectedChart, setSelectedChart, chartRef] = useSyncedState<{ id: string; movable: boolean } | null>(null);
  const [chartOffset, setChartOffset] = useState<{ x: number; y: number } | null>(null);
  const [formatToken, setFormatTokenState, formatTokenRef] = useSyncedState<string | null>(null);
  const formatSource = useRef<string | null>(null);
  const formatRevision = useRef(0);
  const applyingFormat = useRef<number | null>(null);
  const setFormatToken = (token: string | null) => {
    formatRevision.current++;
    applyingFormat.current = null;
    setFormatTokenState(token);
    if (!token) formatSource.current = null;
  };
  const borderStyle = useRef<BorderStyle | undefined>(undefined);
  const borderColor = useRef<string | undefined>(undefined);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen, panelRef] = useSyncedState(false);
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});
  const navigation = useRef(0);
  const [selectionNavigation] = useState(createSelectionDataNavigation);
  const causalSelection = useRef<number | null>(null);
  const mutation = useRef(0);
  const sourceRef = useRef<PeerlessWorkerPaintSource | null>(null);
  const echoRef = useRef<WorkerCellEcho | null>(null);
  const dragging = useRef(false);
  const [selectionFinished, setSelectionFinished] = useState(0);
  const clickStart = useRef<CellAddr | null>(null);
  const chartDrag = useRef<{ generation: number; sheet: number; id: string; zoom: number; x: number; y: number } | null>(null);
  const nudge = useRef<{ generation: number; sheet: number; id: string; dx: number; dy: number } | null>(null);
  const nudgeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const revealing = useRef<{ generation: number; navigation: number } | null>(null);
  const flushNudgeRef = useRef<(retained?: boolean) => void>(() => {});
  const focus = useCallback(() => scrollRef.current?.focus({ preventScroll: true }), []);
  const exportInput = useRef<ReturnType<typeof createPeerlessWorkerInputCoordinator> | null>(null);
  const [bridge] = useState<WorkerEditorApiBridge & XlsxAsyncExportBridge>(() => ({
    coordinator: () => null, readOnly: () => propsRef.current.readOnly ?? false, focus,
    exportSheet: () => activeRef.current, exportCoordinator: () => exportInput.current,
    clearSelection: () => {}, refreshProposals: () => {}, recoverInput: async () => {},
    selectCells: () => false, selectCellsAsync: async () => false, apply: () => {},
  }));
  const { run, error, loading } = usePeerlessSessionWorkbook(props, commands.store, bridge, {
    backend: 'peerless', initialViewport: () => {
      const scroll = scrollRef.current;
      if (!scroll?.clientWidth || !scroll.clientHeight) return undefined;
      return { width: scroll.clientWidth, height: scroll.clientHeight, zoom: zoomRef.current,
        dpr: window.devicePixelRatio || 1, requestGeneration: 0 };
    },
  });
  const ownerRef = useRef(run);
  ownerRef.current = run;
  const sheetSelections = useMemo(() => new Map<string, Selection>(), [run]);
  const sheetScrolls = useMemo(() => new Map<string, { x: number; y: number }>(), [run]);
  const sheetId = (sheet: number) => (ownerRef.current?.session.editState.sheets ??
    ownerRef.current?.session.state.sheets)?.find((entry) => entry.index === sheet)?.id;
  const sheetSelection = (sheet: number): Selection => {
    const id = sheetId(sheet);
    if (id && id === activeSheetId.current && selectionRef.current) return selectionRef.current;
    const target = ownerRef.current?.session.selection;
    return (id ? sheetSelections.get(id) : undefined) ??
      (target && target.sheetId === id ? commandSelection(target) : null) ?? selectionAt({ row: 0, col: 0 });
  };
  const restoreSheetScroll = (view: WorkbookSheetView) => {
    const scroll = scrollRef.current;
    if (!scroll) return;
    const content = scroll.firstElementChild as HTMLElement | null;
    if (content) {
      content.style.width = `${view.contentWidth * zoomRef.current}px`;
      content.style.height = `${view.contentHeight * zoomRef.current}px`;
    }
    const id = sheetId(view.sheet);
    const position = id ? sheetScrolls.get(id) : undefined;
    scroll.scrollLeft = (position?.x ?? view.initialScrollX) * zoomRef.current;
    scroll.scrollTop = (position?.y ?? view.initialScrollY) * zoomRef.current;
  };
  const editable = () => Boolean(ownerRef.current?.ready && !ownerRef.current.retiring &&
    !ownerRef.current.failure && !ownerRef.current.readOnly && !propsRef.current.readOnly && !coordinatorRef.current?.readOnly);
  const report = useCallback((error: unknown) => {
    if (error instanceof PeerlessWorkerCommandRefusal && error.status.reason === 'read-only') {
      setRefusal({ code: 'read-only' });
    } else if (error instanceof XlsxCommandAdmissionError) setRefusal({ code: error.code });
    else setRefusal({ code: 'input-failed', detail: error instanceof Error ? error.message : String(error) });
  }, []);
  const rangeTarget = (sheet: number, selected: Selection): WorkbookCommandSelection => {
    const sheetId = ownerRef.current?.session.editState.sheets?.find((entry) => entry.index === sheet)?.id ??
      ownerRef.current?.session.state.sheets.find((entry) => entry.index === sheet)?.id;
    if (!sheetId) throw new XlsxCommandAdmissionError('editor-unavailable');
    const range = normalizeRange(selected);
    return { sheetId, range: { kind: 'rowCol', start: { row: range.top, col: range.left },
      end: { row: range.bottom, col: range.right } }, activeCell: { ...selected.focus } };
  };
  const selectedTarget = () => {
    const owner = ownerRef.current;
    const sheets = owner?.session.editState.sheets ?? owner?.session.state.sheets;
    return selectionRef.current && sheets?.some((sheet) => sheet.index === activeRef.current)
      ? rangeTarget(activeRef.current, selectionRef.current) : undefined;
  };
  const hidden = () => selectionHidden(sourceRef.current?.painted?.displayList.grid);
  const limits = () => deriveLimits(sourceRef.current?.painted?.displayList ?? null, viewRef.current!,
    (scrollRef.current?.clientHeight ?? 0) / zoomRef.current);
  const predict = (sheet: number, next: Selection) => {
    const nextId = sheetId(sheet);
    if (sheet === activeRef.current && nextId === activeSheetId.current && sameSelection(selectionRef.current, next)) return;
    const previousId = activeSheetId.current;
    const scroll = scrollRef.current;
    if ((nextId !== previousId || sheet !== activeRef.current) && scroll) {
      if (previousId) sheetScrolls.set(previousId, { x: scroll.scrollLeft / zoomRef.current, y: scroll.scrollTop / zoomRef.current });
      const position = nextId ? sheetScrolls.get(nextId) : undefined;
      scroll.scrollLeft = (position?.x ?? 0) * zoomRef.current;
      scroll.scrollTop = (position?.y ?? 0) * zoomRef.current;
    }
    if (nextId !== previousId && previousId && selectionRef.current) {
      sheetSelections.set(previousId, structuredClone(selectionRef.current));
    }
    if (nextId) sheetSelections.set(nextId, structuredClone(next));
    navigation.current += 1;
    activeSheetId.current = nextId;
    setActive(sheet);
    setSelection(structuredClone(next));
    setSelectedChart(null);
    draftController.selectionChanged();
    commands.refresh();
  };
  const reveal = (sheet: number, cell: CellAddr) => {
    const owner = ownerRef.current;
    const frame = sourceRef.current?.painted;
    if (!owner?.current || owner.failure || frame?.geometry && frame.request.sheet === sheet &&
      cellRect(frame.geometry, cell.row, cell.col)) return;
    const at = navigation.current;
    if (revealing.current?.generation === owner.generation && revealing.current.navigation === at) return;
    const request = { generation: owner.generation, navigation: at };
    revealing.current = request;
    void owner.session.call.cellGeometry(sheet, cell.row, cell.col).then((geometry) => {
      const scroll = scrollRef.current;
      if (!owner.current || ownerRef.current !== owner || navigation.current !== at || !scroll) return;
      const id = sheetId(sheet);
      if (id) sheetScrolls.set(id, geometry.scrollPosition);
      scroll.scrollLeft = geometry.scrollPosition.x * zoomRef.current;
      scroll.scrollTop = geometry.scrollPosition.y * zoomRef.current;
      sourceRef.current?.schedule();
    }).catch((error) => {
      if (owner.current && ownerRef.current === owner && navigation.current === at) owner.fail(error);
    }).finally(() => { if (revealing.current === request) revealing.current = null; });
  };
  const requests = useMemo(() => new Map<number, WorkerPaintRequest>(), [run]);
  const selectionCommands = useMemo(() => new Set<number>(), [run]);
  const selectionRevisions = useMemo(() => new Map<number, number>(), [run]);
  const issued = useMemo(() => new Map<number, PeerlessWorkerCommandInput>(), [run]);
  const inputEvents = useRef<(event: PeerlessWorkerInputEvent, publish?: () => void) => void>(() => {});
  const coordinatorRef = useRef<ReturnType<typeof createPeerlessWorkerInputCoordinator> | null>(null);
  const draftController = useMemo(() => new WorkerDraftController({
    generation: () => run?.generation ?? 0,
    identity: () => `${run?.generation}:${run?.session.commandIdentity.session}:${run?.session.commandIdentity.incarnation}:${run?.session.confirmedBasis.epoch}`,
    valueRevision: () => run?.session.confirmedBasis.version ?? '',
    navigation: () => navigation.current, selection: () => selectionRef.current, sheet: () => activeRef.current,
    target: (sheet, row, col) => rangeTarget(sheet, selectionAt({ row, col })),
    resolveSheet: (sheetId) => (run?.session.editState.sheets ?? run?.session.state.sheets)?.find((sheet) => sheet.id === sheetId)?.index,
    editable,
    prefill: (sheet, cell) => readWorkerCellInput(run!.session, sheet, address(cell)),
    select: (sheet, cell, move) => {
      const base = selectionAt(cell);
      predict(sheet, move ? moveFocus(base, move, { limits: limits(), hidden: hidden() }) : base);
      reveal(sheet, selectionRef.current!.focus);
    },
    selectionTarget: selectedTarget,
    dependency: () => causalSelection.current,
    navigate: () => {
      const selected = selectedTarget();
      if (selected) void coordinatorRef.current?.command({ kind: 'select', payload: { selection: selected } }).result.catch(report);
    },
    committed: (ticket) => { void ticket.result.catch(report); commands.refresh(); },
    fail: (error) => run?.fail(error), report, focus,
  }), [run]);
  const capture = useCallback((): WorkerPaintRequest | null => {
    const owner = ownerRef.current;
    const scroll = scrollRef.current;
    if (!owner?.current || !scroll?.clientWidth || !scroll.clientHeight || !canvasRef.current || !viewRef.current) return null;
    const scale = zoomRef.current;
    return { generation: owner.generation, sheet: activeRef.current, navigation: navigation.current,
      coversThroughSeq: Math.max(owner.session.throughSeq, coordinatorRef.current?.submittedThroughSeq ?? 0),
      width: scroll.clientWidth, height: scroll.clientHeight, zoom: scale, dpr: window.devicePixelRatio || 1,
      viewport: { x: scroll.scrollLeft / scale, y: scroll.scrollTop / scale,
        width: scroll.clientWidth / scale, height: scroll.clientHeight / scale } };
  }, []);
  const coordinator = useMemo(() => createPeerlessWorkerInputCoordinator({
    generation: () => run?.generation ?? 0, session: () => run?.session ?? null,
    isReady: () => run?.ready ?? false, readOnly: () => propsRef.current.readOnly ?? false,
    seal: (): InputSeal => {
      if (chartDrag.current) return { refused: 'gesture-active' };
      flushNudgeRef.current();
      return draftController.seal();
    },
    sync: draftController.sync,
    compositionId: () => draftController.compositionId,
    dependency: (draft) => (draft as WorkerNativeDraft).after,
    target: (sheet, row, col) => rangeTarget(sheet, selectionAt({ row, col })),
    draftTarget: (draft) => (draft as WorkerNativeDraft).target,
    frame: () => sourceRef.current?.captureCommandFrame() ?? undefined,
    selectionRevision: () => navigation.current,
    onEvent: (event, publish) => inputEvents.current(event, publish), onError: (error) => run?.fail(error),
  }), [run, draftController]);
  coordinatorRef.current = coordinator;
  exportInput.current = coordinator;
  draftController.attach(coordinator);
  const pluginNavigator = useRef<XlsxPluginDocumentAccess['navigator'] | null>(null);
  const pluginNavigation = useRef<{ owner: PeerlessWorkbookSession; sheetId: string;
    clear: boolean; ticket: PeerlessWorkerCommandTicket | null } | null>(null);
  const acknowledgedSelection = useRef<{ owner: PeerlessWorkbookSession; sheetId: string;
    target: string | undefined; cells: Selection | null } | null>(null);
  const [pluginAccess] = useState(() => createPeerlessPluginDocument({
    owner: () => ownerRef.current, coordinator: () => coordinatorRef.current,
    readOnly: () => propsRef.current.readOnly ?? false, commands: () => commands,
    selection(owner) {
      const target = owner.session.selection;
      const sheets = owner.session.editState.sheets ?? owner.session.state.sheets;
      const activeSheet = owner.session.editState.activeSheet ?? owner.session.state.activeSheet;
      const sheet = sheets.find((entry) => entry.index === activeSheet);
      const selected = target && sheets.find((entry) => entry.id === target.sheetId);
      const current = sheet ?? selected;
      const projection = acknowledgedSelection.current;
      const cells = projection?.owner === owner && projection.sheetId === current?.id &&
        projection.target === JSON.stringify(target) ? projection.cells
        : target && selected?.id === current?.id ? commandSelection(target) : null;
      return current ? { sheetId: current.id, sheetIndex: current.index,
        cells: chartRef.current ? null : cells,
        chartId: chartRef.current?.id ?? null } : null;
    },
    navigator: {
      selectCells: (...args) => pluginNavigator.current!.selectCells(...args),
      scrollToCell: (...args) => pluginNavigator.current!.scrollToCell(...args),
    },
  }));
  const pluginIdentity = pluginAccess.identity();
  const pluginSelection = pluginIdentity ? pluginAccess.selection!(pluginIdentity) : null;
  const pluginHost = useXlsxPluginHost({ ...props, access: pluginAccess, commands,
    createClients: createPeerlessPluginClients, createGeometry: createPeerlessPluginGeometry,
    readOnly: props.readOnly ?? false, handle: pluginIdentity, selection: pluginSelection, canvasRef, t });
  pluginBindingRef.current = pluginHost;
  const [dockAreaRef, dockArea] = useDockArea(pluginHost.managed);
  useLayoutEffect(() => run?.session.bindCollaboration?.(async (update) => {
    const status = await coordinator.command({ kind: 'applyUpdate', payload: { update } }, { after: null }).result;
    if (status.status !== 'applied' && status.status !== 'noop') throw new Error(status.reason ?? 'Remote update refused');
    return status;
  }), [run, coordinator]);
  useLayoutEffect(() => { run?.connectRecoveryInput(coordinator, draftController); }, [run, coordinator, draftController]);
  const proposalsRevision = useRef(0);
  const proposalsMutation = useRef(0);
  const proposalsLanded = useRef<{ owner: PeerlessWorkbookSession; revision: number; sequence: number } | null>(null);
  const proposalsPending = useRef<{ owner: PeerlessWorkbookSession; requiredSequence: number; promise: Promise<void> } | null>(null);
  const refreshProposals = useCallback((force = false): Promise<void> => {
    const owner = ownerRef.current;
    if (!owner?.ready || owner.failure || owner.retiring || !isProposalsAvailable()) return Promise.resolve();
    const requiredSequence = force ? coordinator.submittedThroughSeq : proposalsMutation.current;
    if (proposalsPending.current?.owner === owner) {
      proposalsPending.current.requiredSequence = Math.max(proposalsPending.current.requiredSequence, requiredSequence);
      return proposalsPending.current.promise;
    }
    if (force) proposalsRevision.current++;
    if (proposalsLanded.current?.owner === owner && proposalsLanded.current.revision === proposalsRevision.current &&
      proposalsLanded.current.sequence >= proposalsMutation.current) return Promise.resolve();
    const pending = { owner, requiredSequence, promise: Promise.resolve() };
    proposalsPending.current = pending;
    pending.promise = Promise.resolve().then(async () => {
      while (true) {
        if (ownerRef.current !== owner || !owner.current || owner.retiring) throw new XlsxCommandAdmissionError('document-replaced');
        if (owner.failure) throw owner.failure;
        const required = pending.requiredSequence;
        await coordinator.whenApplied(required);
        if (required !== pending.requiredSequence) continue;
        const revision = proposalsRevision.current;
        const ticket = coordinator.command({ kind: 'applyOps', payload: { toolbar: { id: 'refreshProposals', args: null } } },
          { after: null, frame: undefined });
        const status = await ticket.result;
        if (ownerRef.current !== owner || !owner.current || owner.retiring) throw new XlsxCommandAdmissionError('document-replaced');
        if (owner.failure) throw owner.failure;
        const sequence = ticket.seq ?? 0;
        const requiredSequence = Math.max(proposalsMutation.current, pending.requiredSequence);
        if (revision !== proposalsRevision.current || sequence < requiredSequence) {
          const landed = proposalsLanded.current;
          if (landed?.owner === owner && landed.revision === proposalsRevision.current && landed.sequence >= requiredSequence) return;
          continue;
        }
        const result = status.result as { proposals?: Proposal[] } | undefined;
        if (result?.proposals && JSON.stringify(result.proposals) !== JSON.stringify(proposalsRef.current)) {
          setProposals(result.proposals); commands.refresh();
        }
        proposalsLanded.current = { owner, revision, sequence };
        return;
      }
    }).finally(() => { if (proposalsPending.current === pending) proposalsPending.current = null; });
    return pending.promise;
  }, [coordinator, commands]);
  const requestProposals = (force = false) => {
    const owner = ownerRef.current;
    void refreshProposals(force).catch((error) => {
      if (owner?.current && !owner.retiring && !owner.failure && ownerRef.current === owner) report(error);
    });
  };
  const geometryRevision = useRef(0);
  const refreshSheetGeometry = useCallback((paint = false) => {
    const owner = ownerRef.current;
    if (!owner?.ready || owner.failure) return;
    const sheet = activeRef.current;
    const revision = ++geometryRevision.current;
    void owner.session.call.sheetView(sheet).then((next) => {
      if (ownerRef.current !== owner || !owner.current || owner.failure ||
        revision !== geometryRevision.current || activeRef.current !== next.sheet) return;
      const changedSheet = viewRef.current?.sheet !== next.sheet;
      setView(next);
      if (changedSheet) restoreSheetScroll(next);
      if (paint || changedSheet) sourceRef.current?.schedule();
    }).catch((error) => {
      if (ownerRef.current === owner && owner.current && !owner.failure) owner.fail(error);
    });
  }, [setView, viewRef, sheetScrolls]);
  const selectionKey = () => {
    const selected = selectionRef.current;
    if (!selected) return '';
    const range = normalizeRange(selected);
    return `${activeRef.current}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
  };
  inputEvents.current = (event, publish) => {
    if (!run?.current || run.failure) return;
    if (event.type === 'dispatched' && event.seq !== null && event.command) {
      const command = event.command;
      issued.set(event.seq, command);
      const revision = command.kind === 'editCells' &&
        JSON.stringify(command.payload.selection) === JSON.stringify(selectedTarget())
        ? navigation.current : event.selectionRevision ?? navigation.current;
      selectionRevisions.set(event.seq, revision);
      while (selectionRevisions.size > 256) selectionRevisions.delete(selectionRevisions.keys().next().value!);
      if (command.frame) requests.set(event.seq, { ...command.frame as WorkerPaintRequest,
        navigation: revision, coversThroughSeq: event.seq });
      while (requests.size > 256) requests.delete(requests.keys().next().value!);
      if (command.kind === 'select' || command.kind === 'setActiveSheet' || command.kind === 'undo' || command.kind === 'redo' ||
        'selection' in command.payload) {
        selectionCommands.add(event.seq); causalSelection.current = event.seq;
      }
      if (command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'paintFormat') {
        const target = command.payload.toolbar.target;
        const sheet = run.session.editState.sheets?.find((entry) => entry.id === target?.sheetId)?.index;
        if (sheet !== undefined && target?.range.kind === 'rowCol') formatSource.current =
          `${sheet}:${target.range.start.row}:${target.range.start.col}:${target.range.end.row}:${target.range.end.col}`;
      }
      if (event.draft) {
        echoRef.current?.patch(event.seq, [{ ...event.draft, text: event.draft.value }]);
        draftController.submitted(event.draft, event.seq);
      }
      else if (command.kind === 'editCells') {
        const [sheet, edits] = command.payload.args;
        echoRef.current?.paste(event.seq, sheet, edits);
      }
      else if (command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'clear') {
        const target = command.payload.toolbar.target;
        const sheet = run.session.editState.sheets?.find((entry) => entry.id === target?.sheetId)?.index;
        if (sheet !== undefined && target?.range.kind === 'rowCol') echoRef.current?.clear(event.seq, {
          sheet, top: target.range.start.row, left: target.range.start.col,
          bottom: target.range.end.row, right: target.range.end.col,
        });
      }
      return;
    }
    if (event.type === 'accepted' && event.seq !== null) {
      const command = issued.get(event.seq);
      issued.delete(event.seq);
      const refreshed = command?.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'refreshProposals';
      if (command && !refreshed && command.kind !== 'select' && command.kind !== 'setActiveSheet' &&
        command.kind !== 'setReadOnly' && command.kind !== 'applyUpdate' &&
        (event.status?.committed || command.kind === 'rejectProposal' ||
          (command.kind === 'propose' || command.kind === 'acceptProposal' || command.kind === 'undo' || command.kind === 'redo') &&
            event.status?.status === 'applied')) {
        proposalsMutation.current = Math.max(proposalsMutation.current, event.seq);
        requestProposals();
      }
      return;
    }
    if (event.type === 'refused' || event.type === 'cancelled') {
      if (event.status?.reason === 'stale-proposal') { proposalsRevision.current++; requestProposals(); }
      if (event.seq !== null) issued.delete(event.seq);
      if (event.seq !== null) echoRef.current?.refuse(event.seq);
      if (event.type === 'refused' && event.draft) {
        draftController.refused(event.seq, event.draft);
        report(new PeerlessWorkerCommandRefusal(event.status ?? { seq: event.seq ?? 0, status: 'refused' }, event.target, event.draft));
      }
      return;
    }
    if (event.type !== 'coverage') return;
    echoRef.current?.acknowledge(event.ack);
    const ack = event.ack;
    const pendingNavigation = pluginNavigation.current;
    const navigationStatus = pendingNavigation?.owner === run
      ? ack.statuses.find((status) => status.seq === pendingNavigation.ticket?.seq) : undefined;
    if (pendingNavigation?.clear && navigationStatus &&
      (navigationStatus.status === 'applied' || navigationStatus.status === 'noop')) {
      acknowledgedSelection.current = { owner: run, sheetId: pendingNavigation.sheetId,
        target: JSON.stringify(run.session.selection), cells: null };
    } else if (ack.statuses.some((status) => selectionCommands.has(status.seq) &&
      (status.status === 'applied' || status.status === 'noop'))) {
      acknowledgedSelection.current = null;
    }
    for (const status of ack.statuses) {
      const result = status.result as { remoteUpdate?: boolean; proposals?: Proposal[] } | undefined;
      if (!result?.remoteUpdate) continue;
      proposalsMutation.current = Math.max(proposalsMutation.current, status.seq);
      if (result.proposals) {
        setProposals(result.proposals);
        proposalsLanded.current = { owner: run, revision: proposalsRevision.current, sequence: status.seq };
      }
      sourceRef.current?.schedule();
      commands.refresh();
    }
    if (ack.editState?.sheets && JSON.stringify(ack.editState.sheets) !== JSON.stringify(catalogRef.current)) {
      setCatalog(ack.editState.sheets);
    }
    const latestSelection = Math.max(0, ...selectionCommands);
    const reconcile = latestSelection > 0 && ack.throughSeq >= latestSelection &&
      selectionRevisions.get(latestSelection) === navigation.current &&
      !draftController.ownsFocus;
    if (ack.selection && reconcile) {
      const target = ack.selection;
      const sheet = run.session.editState.sheets?.find((entry) => entry.id === target.sheetId)?.index;
      const restored = commandSelection(target, selectionRef.current);
      if (sheet !== undefined && restored) predict(sheet, restored);
    } else if (!ack.selection && ack.editState?.activeSheet !== undefined && reconcile && ack.editState.activeSheet !== activeRef.current) {
      predict(ack.editState.activeSheet, sheetSelection(ack.editState.activeSheet));
    }
    publish?.();
    if (ack.frame) {
      const frameRequest = requests.get(ack.frame.coversThroughSeq);
      if (!frameRequest || !sourceRef.current?.install(ack.frame,
        reconcile ? { ...frameRequest, navigation: navigation.current } : frameRequest)) sourceRef.current?.schedule();
    }
    for (const seq of requests.keys()) if (seq <= (sourceRef.current?.coversThroughSeq ?? 0)) requests.delete(seq);
    for (const seq of selectionCommands) if (seq <= ack.throughSeq) selectionCommands.delete(seq);
    for (const seq of selectionRevisions.keys()) if (seq <= ack.throughSeq) selectionRevisions.delete(seq);
    if (causalSelection.current !== null && causalSelection.current <= ack.throughSeq) causalSelection.current = null;
    if (ack.statuses.some((status) => status.committed && status.status === 'applied')) {
      mutation.current += 1;
      refreshSheetGeometry();
      propsRef.current.onChange?.();
    }
  };

  const pluginReveal = async (owner: PeerlessWorkbookSession, sheet: number, cell: CellAddr,
    align: XlsxRevealAlignment, live: () => boolean, version: string): Promise<XlsxPluginNavigationFailureCode | null> => {
    const revision = navigation.current;
    const painted = sourceRef.current?.painted;
    const cachedView = viewRef.current?.sheet === sheet && viewRef.current.version === version ? viewRef.current : null;
    const rect = painted?.version === version && painted.request.sheet === sheet && painted.geometry
      ? cellRect(painted.geometry, cell.row, cell.col) : null;
    const [geometry, view] = await Promise.all([
      rect && painted ? { version, rect }
        : owner.session.call.cellGeometry(sheet, cell.row, cell.col),
      cachedView ?? owner.session.call.sheetView(sheet),
    ]);
    if (ownerRef.current !== owner || !owner.current || owner.failure || !live() ||
      navigation.current !== revision || activeRef.current !== sheet) return 'unsupported';
    if (geometry.version !== version || view.version !== version || owner.session.confirmedBasis.version !== version) {
      return 'stale-version';
    }
    const scroll = scrollRef.current;
    if (!scroll) return 'layout-unavailable';
    if (viewRef.current !== view) { geometryRevision.current += 1; setView(view); }
    const scale = zoomRef.current;
    const left = scroll.scrollLeft / scale;
    const top = scroll.scrollTop / scale;
    const x = revealAxis(geometry.rect.x + (rect && painted && cell.col >= view.frozenCols ? painted.request.viewport.x : 0),
      geometry.rect.w, cell.col < view.frozenCols, left, scroll.clientWidth / scale - view.frozenWidth, align);
    const y = revealAxis(geometry.rect.y + (rect && painted && cell.row >= view.frozenRows ? painted.request.viewport.y : 0),
      geometry.rect.h, cell.row < view.frozenRows, top, scroll.clientHeight / scale - view.frozenHeight, align);
    const movedX = x !== null && x !== left;
    const movedY = y !== null && y !== top;
    if (movedX) scroll.scrollLeft = x * scale;
    if (movedY) scroll.scrollTop = y * scale;
    if (movedX || movedY || painted?.request.sheet !== sheet || painted.version !== version) sourceRef.current?.schedule();
    return null;
  };
  const pluginNavigate = async (identity: object, target: { sheetId: string; selection?: Selection; row?: number; col?: number },
    align: XlsxRevealAlignment, wantsFocus: boolean, live: () => boolean,
    expectVersion: string): Promise<XlsxPluginNavigationFailureCode | null> => {
    const owner = ownerRef.current;
    if (!owner || owner !== identity || !owner.current || owner.failure || !live()) return 'unsupported';
    const sheet = (owner.session.editState.sheets ?? owner.session.state.sheets).find((entry) => entry.id === target?.sheetId);
    const cell = target?.selection?.focus ?? { row: target?.row!, col: target?.col! };
    const valid = (value: CellAddr | undefined) => value && Number.isInteger(value.row) && Number.isInteger(value.col) &&
      value.row >= 0 && value.row < 1_048_576 && value.col >= 0 && value.col < 16_384;
    if (!sheet || !valid(cell) || target.selection && !valid(target.selection.anchor)) return 'missing-target';
    let pendingNavigation: typeof pluginNavigation.current = null;
    try {
      if (target.selection || sheet.index !== activeRef.current) {
        const selected = target.selection ? rangeTarget(sheet.index, target.selection) : undefined;
        const revision = navigation.current;
        let ticket: PeerlessWorkerCommandTicket | null = null;
        let acknowledgedVersion: string | null = null;
        pendingNavigation = { owner, sheetId: sheet.id, clear: !target.selection,
          ticket: null as PeerlessWorkerCommandTicket | null };
        pluginNavigation.current = pendingNavigation;
        const offAck = owner.session.onAck((ack) => {
          const status = ack.statuses.find((entry) => entry.seq === ticket?.seq);
          if (!status) return;
          acknowledgedVersion = ack.statuses.some((entry) => entry.seq !== ticket?.seq && entry.committed) ||
            !status.committed && ack.version !== expectVersion ? null : ack.version;
        });
        try {
          ticket = coordinator.command({ kind: 'setActiveSheet', payload: { args: [sheet.index],
            target: { sheetId: sheet.id, range: { kind: 'rowCol', start: cell, end: cell } },
            ...(selected ? { selection: selected } : {}), expectVersion } }, {
            frame: undefined,
            dispatch(write) {
              if (ownerRef.current !== owner || !owner.current || owner.failure || !live()) {
                throw new XlsxCommandAdmissionError('document-replaced');
              }
              return write();
            },
          });
          pendingNavigation.ticket = ticket;
          await ticket.result;
        } finally { offAck(); }
        if (!acknowledgedVersion || owner.session.confirmedBasis.version !== acknowledgedVersion) return 'stale-version';
        if (ownerRef.current !== owner || !owner.current || owner.failure || !live() ||
          navigation.current > revision + 1 || activeRef.current !== sheet.index ||
          target.selection && !sameSelection(selectionRef.current, target.selection)) return 'unsupported';
        if (!target.selection) setSelection(null);
        expectVersion = acknowledgedVersion;
      }
      const revision = navigation.current;
      const failure = await pluginReveal(owner, sheet.index, cell, align, live, expectVersion);
      if (failure) return failure;
      if (ownerRef.current !== owner || !owner.current || owner.retiring || owner.failure || !live() ||
        navigation.current !== revision || activeRef.current !== sheet.index) return 'unsupported';
      if (owner.session.confirmedBasis.version !== expectVersion) return 'stale-version';
      if (target.selection) {
        flushSync(() => setSelectedChart(null));
        if (ownerRef.current !== owner || !owner.current || owner.retiring || owner.failure || !live() ||
          navigation.current !== revision || activeRef.current !== sheet.index) return 'unsupported';
        if (owner.session.confirmedBasis.version !== expectVersion) return 'stale-version';
        const source = sourceRef.current;
        const painted = source?.painted;
        const request = source?.captureCommandFrame();
        const seq = pendingNavigation?.ticket?.seq;
        if (source && seq && painted?.version === expectVersion && request &&
          samePaintRequest(painted, { ...request, navigation: painted.request.navigation })) {
          source.install({ ...owner.session.commandIdentity, sheet: request.sheet, viewport: request.viewport,
            requestGeneration: request.requestGeneration!, zoom: request.zoom, dpr: request.dpr,
            version: painted.version, epoch: owner.session.confirmedBasis.epoch, sequence: painted.sequence,
            coversThroughSeq: seq, displayList: painted.displayList, mergedRanges: [...painted.mergedRanges] },
          { ...request, coversThroughSeq: seq });
        }
      }
      if (wantsFocus) focus();
      return null;
    } catch (error) {
      if (error instanceof PeerlessWorkerCommandRefusal) {
        if (error.status.reason === 'stale-version') return 'stale-version';
        if (error.status.reason === 'missing-target' || error.status.reason === 'stale-target') return 'missing-target';
      }
      if (error instanceof RangeError) return 'missing-target';
      throw error;
    } finally {
      if (pluginNavigation.current === pendingNavigation) {
        pluginNavigation.current = null;
        if (ownerRef.current === owner && owner.current && !owner.failure &&
          activeRef.current !== viewRef.current?.sheet) refreshSheetGeometry(true);
      }
    }
  };
  pluginNavigator.current = {
    selectCells: (identity, target, wantsFocus, live, version) => pluginNavigate(identity,
      structuredClone(target), 'nearest', wantsFocus, live, version),
    scrollToCell: (identity, target, align, live, version) => pluginNavigate(identity,
      structuredClone(target), align, false, live, version),
  };

  useLayoutEffect(() => {
    const cancelGestures = () => {
      if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
      nudgeTimer.current = null;
      nudge.current = null; chartDrag.current = null; dragging.current = false; clickStart.current = null;
      revealing.current = null;
    };
    pluginHost.beginLoad();
    cancelGestures();
    setRefusal(null); presentation.reset();
    activeSheetId.current = undefined;
    setSelectedChart(null); setChartOffset(null); setFormatToken(null); setProposals([]); setStaleFor({});
    setProposalsPanelOpen(false); formatSource.current = null;
    setCatalog(run?.session.editState.sheets ?? run?.session.state.sheets ?? []);
    navigation.current += 1;
    causalSelection.current = null;
    proposalsRevision.current++;
    proposalsMutation.current = 0;
    acknowledgedSelection.current = null;
    pluginNavigation.current = null;
    if (!run) return;
    const source = new PeerlessWorkerPaintSource({
      generation: run.generation, commandIdentity: run.session.commandIdentity,
      capture, isCurrent: () => run.current && !run.failure,
      sentSequence: () => coordinator.submittedThroughSeq,
      requestFrame: async (request) => {
        if (run.ready) await coordinator.whenApplied(Math.min(request.coversThroughSeq ?? 0, coordinator.submittedThroughSeq));
        const coverage = Math.max(request.coversThroughSeq ?? 0, run.session.throughSeq);
        const frame = await run.session.call.frame(request.viewport, { sheet: request.sheet });
        return { ...frame, coversThroughSeq: 'coversThroughSeq' in frame && typeof frame.coversThroughSeq === 'number'
          ? frame.coversThroughSeq : coverage };
      },
      publish(frame) {
        const canvas = canvasRef.current;
        const context = canvas?.getContext('2d');
        if (!canvas || !context) throw new Error('Canvas context is unavailable');
        const request = frame.request;
        canvas.width = Math.round(request.width * request.dpr);
        canvas.height = Math.round(request.height * request.dpr);
        canvas.style.width = `${request.width}px`; canvas.style.height = `${request.height}px`;
        paintDisplayList(context, frame.displayList, request.dpr * request.zoom);
        canvas.dataset.workerSequence = String(frame.sequence);
        canvas.dataset.commandCoverage = String(frame.coversThroughSeq ?? 0);
        canvas.dataset.paintSource = frame.source; canvas.dataset.workerZoom = String(request.zoom);
      },
      onInstalled(frame) {
        setPainted(frame);
        const sheetId = (run.session.editState.sheets ?? run.session.state.sheets).find((sheet) => sheet.index === frame.request.sheet)?.id;
        pluginHost.presentGrid(sheetId ? { frame: frame.displayList, version: frame.version,
          sheetId, zoom: frame.request.zoom, viewport: frame.request.viewport } : null);
        if (frame.coversThroughSeq !== undefined) run.session.markPresented({
          ...run.session.commandIdentity, coversThroughSeq: frame.coversThroughSeq,
        });
        run.firstPaint();
      },
      onError: (error) => run.fail(error),
    });
    sourceRef.current = source;
    const echo = new WorkerCellEcho({ source, context: () => canvasRef.current?.getContext('2d') ?? null });
    echoRef.current = echo;
    const sheet = run.session.editState.activeSheet ?? run.session.state.activeSheet;
    activeSheetId.current = sheetId(sheet);
    setActive(sheet);
    const openView = (next: WorkbookSheetView) => {
      if (!run.current || run.failure || sourceRef.current !== source) return;
      setView(next);
      activeSheetId.current = sheetId(sheet);
      setSelection(structuredClone(sheetSelection(sheet)));
      restoreSheetScroll(next);
    };
    const initial = run.session.initialPaint;
    if (initial) { openView(initial.sheetView); if (!source.installInitial(initial.frame)) source.schedule(); }
    else void run.session.call.sheetView(sheet).then((next) => {
      openView(next); source.schedule();
    }).catch((error) => run.fail(error));
    run.beforeRetire = () => { pluginHost.beginLoad(); cancelGestures(); presentation.cancel(); coordinator.retire(); draftController.dispose(); source.dispose(); };
    return () => {
      cancelGestures(); presentation.cancel();
      echo.dispose(); source.dispose(); coordinator.retire(); draftController.dispose();
      if (sourceRef.current === source) sourceRef.current = null;
      if (echoRef.current === echo) echoRef.current = null;
    };
  }, [run, coordinator, draftController, capture]);

  useEffect(() => {
    if (!run?.current || run.failure) return;
    const scroll = scrollRef.current;
    if (!scroll) return;
    const schedule = () => {
      const source = sourceRef.current;
      const request = capture();
      const frame = source?.painted;
      if (frame && request && (frame.coversThroughSeq ?? -1) >= coordinator.submittedThroughSeq &&
        samePaintRequest(frame, request)) return;
      source?.schedule();
    };
    const observer = new ResizeObserver(schedule);
    observer.observe(scroll); scroll.addEventListener('scroll', schedule, { passive: true });
    schedule();
    return () => { observer.disconnect(); scroll.removeEventListener('scroll', schedule); };
  }, [run, run?.failure, zoom, capture, coordinator]);
  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList | undefined;
    const changed = () => {
      sourceRef.current?.schedule(); media?.removeEventListener('change', changed);
      media = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      media.addEventListener('change', changed);
    };
    changed();
    return () => media?.removeEventListener('change', changed);
  }, []);
  useEffect(() => {
    if (!run?.ready || active === viewRef.current?.sheet) return;
    if (pluginNavigation.current?.owner === run && pluginNavigation.current.sheetId ===
      run.session.editState.sheets?.find((sheet) => sheet.index === active)?.id) return;
    refreshSheetGeometry(true);
  }, [run, active, refreshSheetGeometry]);
  const mode = useRef<{ owner: PeerlessWorkbookSession | null; readOnly: boolean }>({ owner: null, readOnly: false });
  useLayoutEffect(() => {
    if (!run?.ready || mode.current.owner === run && mode.current.readOnly === Boolean(props.readOnly)) return;
    const previous = mode.current;
    mode.current = { owner: run, readOnly: Boolean(props.readOnly) };
    if (previous.owner !== run && Boolean(run.session.editState.readOnly) === Boolean(props.readOnly)) return;
    if (props.readOnly) {
      flushNudgeRef.current(true);
      draftController.finish(); setFormatToken(null); chartDrag.current = null; nudge.current = null;
      setChartOffset(null);
    }
    void coordinator.setReadOnly(Boolean(props.readOnly)).result.catch(report);
  }, [run, run?.ready, props.readOnly, coordinator, draftController]);
  useEffect(() => {
    if (!run?.failure) return;
    echoRef.current?.dispose(); sourceRef.current?.dispose(); coordinator.fail(run.failure); draftController.dispose();
    chartDrag.current = null; nudge.current = null;
  }, [run, run?.failure, coordinator, draftController]);

  const command = (input: PeerlessWorkerCommandInput) => {
    const frame = sourceRef.current?.captureCommandFrame() ?? undefined;
    const selectionRevision = navigation.current;
    void coordinator.runAfterPendingInput(() => coordinator.command(input, { frame, selectionRevision }).result).catch(report);
  };
  const clear = () => {
    const target = selectedTarget();
    if (!editable() || !target) return;
    command({ kind: 'applyOps', payload: { toolbar: { id: 'clear', args: null, target }, selection: target } });
  };
  const clipboardAllowed = (kind: 'copy' | 'cut' | 'paste') => {
    const owner = ownerRef.current;
    return Boolean(owner?.ready && !owner.failure && selectionRef.current && (kind === 'copy' || editable()));
  };
  const clipboardAccess = (kind: 'copy' | 'cut' | 'paste'): { read?: Promise<string>; write?: Promise<void>;
    resolve?: (blob: Blob) => void; reject?: (error: unknown) => void; refused?: { error: unknown } } => {
    if (kind === 'paste') {
      const read = (async () => navigator.clipboard.readText())();
      void read.catch(() => {});
      return { read };
    }
    let resolve!: (blob: Blob) => void;
    let reject!: (error: unknown) => void;
    const payload = new Promise<Blob>((done, fail) => { resolve = done; reject = fail; });
    void payload.catch(() => {});
    let write: Promise<void> | undefined;
    let refused: { error: unknown } | undefined;
    try {
      if (typeof ClipboardItem === 'function' && typeof navigator.clipboard?.write === 'function') {
        write = navigator.clipboard.write([new ClipboardItem({ 'text/plain': payload })]);
        void write.catch(() => {});
      } else if (kind === 'cut' && typeof navigator.clipboard?.writeText !== 'function') {
        throw new WorkerInputRefusal('Cut requires clipboard write access');
      }
    } catch (error) { refused = { error }; }
    return { write, resolve, reject, refused };
  };
  const clipboard = (kind: 'copy' | 'cut' | 'paste', access: ReturnType<typeof clipboardAccess>) => {
    if (!clipboardAllowed(kind)) { access.reject?.(new XlsxCommandAdmissionError('editor-unavailable')); return; }
    const owner = ownerRef.current!;
    const selected = selectionRef.current!;
    const sheet = activeRef.current;
    const at = navigation.current;
    const target = rangeTarget(sheet, selected);
    const range = normalizeRange(selected);
    if (kind === 'paste') {
      const ticket = coordinator.clipboard(() => access.read!, (text) => {
        const rows = fromTsv(text);
        const edits: CellInputEdit[] = [];
        let lastCol = 0;
        rows.forEach((row, dr) => {
          lastCol = Math.max(lastCol, row.length - 1);
          row.forEach((input, dc) => edits.push({ row: range.top + dr, col: range.left + dc, input }));
        });
        const next: Selection = { anchor: { row: range.top, col: range.left }, focus: {
          row: range.top + Math.max(0, rows.length - 1),
          col: range.left + lastCol,
        } };
        if (ownerRef.current === owner && navigation.current === at && !draftController.ownsFocus) predict(sheet, next);
        return { kind: 'editCells', payload: { args: [sheet, edits], target, selection: rangeTarget(sheet, next) } };
      });
      void ticket.result.catch(report);
      return;
    }
    const checkOwner = () => {
      if (ownerRef.current !== owner || !owner.current || owner.retiring) throw new XlsxCommandAdmissionError('document-replaced');
      if (owner.failure) throw owner.failure;
    };
    const { write, resolve, reject, refused } = access;
    if (refused) { reject!(refused.error); report(refused.error); return; }
    const copied = coordinator.runAfterPendingInput(async () => {
      checkOwner();
      const result = await owner.session.call.cellInputs(sheet, selectedRange(selected));
      checkOwner();
      return toTsv(result.cells);
    }, kind === 'cut' ? true : 'posted');
    const written = copied.then(async (text) => {
      checkOwner();
      if (write) { resolve!(new Blob([text], { type: 'text/plain' })); await write; }
      else await navigator.clipboard.writeText(text);
      checkOwner();
    });
    if (kind === 'cut') {
      const ticket = coordinator.clipboard(() => written, () => ({ kind: 'applyOps', payload: {
        toolbar: { id: 'clear', args: null, target }, selection: target,
      } }));
      void ticket.result.catch((error) => { reject!(error); report(error); });
    } else void written.catch((error) => { reject!(error); report(error); });
  };
  const navigate = (sheet: number, next?: Selection) => {
    if (!run?.ready || run.failure || draftController.isComposing('cell') || draftController.isComposing('formula')) return;
    draftController.commit('cell'); draftController.commit('formula');
    const selected = next ?? sheetSelection(sheet);
    const target = rangeTarget(sheet, selected);
    predict(sheet, selected);
    if (next) reveal(sheet, selected.focus);
    command({ kind: 'setActiveSheet', payload: { args: [sheet], target, selection: target } });
  };
  const moveChart = (sheet: number, id: string, dx: number, dy: number, retained = false, generation = run?.generation) => {
    if (!run?.ready || !run.current || run.retiring || generation !== run.generation || run.failure ||
      !retained && !editable() || dx === 0 && dy === 0) return;
    const target = rangeTarget(sheet, selectionRef.current ?? selectionAt({ row: 0, col: 0 }));
    command({ kind: 'moveChart', payload: { args: [sheet, id, dx, dy], target } });
    setChartOffset(null);
  };
  const flushNudge = (retained = false) => {
    const pending = nudge.current;
    nudge.current = null;
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    if (pending) moveChart(pending.sheet, pending.id, pending.dx, pending.dy, retained, pending.generation);
  };
  flushNudgeRef.current = flushNudge;
  const gestures = useRef({ moveChart, flushNudge });
  gestures.current = { moveChart, flushNudge };
  useEffect(() => {
    const stop = (event: globalThis.MouseEvent) => {
      const selecting = dragging.current;
      dragging.current = false;
      if (selecting && event.button === 0) setSelectionFinished((revision) => revision + 1);
      const drag = chartDrag.current;
      chartDrag.current = null;
      if (drag && event.button === 0) gestures.current.moveChart(drag.sheet, drag.id,
        (event.clientX - drag.x) / drag.zoom, (event.clientY - drag.y) / drag.zoom, false, drag.generation);
      else setChartOffset(null);
    };
    const cancel = () => { dragging.current = false; chartDrag.current = null; setChartOffset(null); gestures.current.flushNudge(); };
    window.addEventListener('mouseup', stop); window.addEventListener('blur', cancel);
    return () => {
      window.removeEventListener('mouseup', stop); window.removeEventListener('blur', cancel);
      if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    };
  }, []);

  const requestSave = useSaveRequest({
    document: () => ownerRef.current?.current ? ownerRef.current : null,
    onSaveRequest: props.onSaveRequest, fail: report,
  });
  usePeerlessWorkerXlsxCommands(commands, {
    requestSave,
    session: () => run?.session ?? null, coordinator,
    status: () => run?.ready ? 'ready' : run && !run.failure || loading ? 'loading' : 'empty',
    readOnly: () => propsRef.current.readOnly ?? false, collaborative: () => propsRef.current.collaboration !== undefined,
    generation: () => run?.generation ?? 0, mutation: () => mutation.current,
    selection: () => selectedTarget() ?? null,
    frame: () => sourceRef.current?.captureCommandFrame() ?? undefined,
    selectionRevision: () => navigation.current,
    view: () => ({ sheet: activeRef.current, selection: selectionRef.current, chartSelected: chartRef.current !== null,
      zoom: zoomRef.current, capturedFormat: null, borderStyle: borderStyle.current, borderColor: borderColor.current,
      proposals: proposalsRef.current, proposalsAvailable: isProposalsAvailable(), proposalsPanelOpen: panelRef.current,
      pngExport: run?.session.pngExport === true }),
    formatToken: () => formatTokenRef.current,
    formatRevision: () => formatRevision.current,
    setFormatToken,
    setZoom: (scale) => setZoom(scale),
    setBorderStyle: (value) => { borderStyle.current = value; }, setBorderColor: (value) => { borderColor.current = value; },
    setProposalsPanelOpen: (open) => { setProposalsPanelOpen(open); if (open) requestProposals(true); },
    markStale: (id, cells) => setStaleFor(({ [id]: _removed, ...rest }) => cells ? { ...rest, [id]: cells } : rest),
    refreshProposals: () => requestProposals(), i18n: () => propsRef.current.i18n, translate: t,
    fail: (error) => run?.fail(error), afterPaint: () => coordinator.whenPresented().then(() => true, () => false),
    deliver(bytes) {
      if (propsRef.current.onSave) propsRef.current.onSave(bytes);
      else {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: XLSX_MIME }));
        const anchor = document.createElement('a');
        anchor.href = url; anchor.download = propsRef.current.fileName ?? 'workbook.xlsx';
        document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
      }
    },
    deliverPng(bytes) {
      const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' }));
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `${(propsRef.current.fileName ?? 'workbook.xlsx').replace(/\.xlsx$/i, '')}.png`;
      document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
    },
    focusGrid: focus, refuse: (reason) => setRefusal({ code: reason.code as XlsxCommandFailureCode }),
  });
  useCommandShortcuts(commands, rootRef);
  const applySelectionFormat = useCallback(() => {
    const selection = selectionRef.current;
    const active = activeRef.current;
    if (!formatToken || formatTokenRef.current !== formatToken || !selection || !editable() || dragging.current ||
      applyingFormat.current !== null || selectionKey() === formatSource.current) return;
    const target = rangeTarget(active, selection);
    const owner = run;
    const revision = ++formatRevision.current;
    applyingFormat.current = revision;
    const frame = sourceRef.current?.captureCommandFrame() ?? undefined;
    const selectionRevision = navigation.current;
    void coordinator.runAfterPendingInput(() => coordinator.command({ kind: 'applyFormat',
      payload: { target, formatToken, selection: target },
    }, { frame, selectionRevision }).result).then(() => {
      if (ownerRef.current !== owner || !owner?.current || formatRevision.current !== revision) return;
      setFormatToken(null); commands.refresh();
    }).catch(report).finally(() => {
      if (applyingFormat.current === revision) applyingFormat.current = null;
    });
  }, [formatToken, run, coordinator, selectionFinished]);

  const point = (event: MouseEvent) => {
    const frame = sourceRef.current?.painted;
    const canvas = canvasRef.current;
    if (!frame || frame.request.sheet !== activeRef.current || !canvas) return { cell: null, chart: null };
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / frame.request.zoom;
    const y = (event.clientY - rect.top) / frame.request.zoom;
    return { cell: frame.geometry ? cellAtPoint(frame.geometry, x, y) : null,
      chart: chartRegionAtPoint(frame.displayList.charts, x, y) };
  };
  const openCell = (seed?: string) => {
    const selected = selectionRef.current;
    if (!run?.ready || !editable() || !selected) return;
    reveal(activeRef.current, selected.focus);
    draftController.open(seed);
  };
  const onMouseDown = (event: MouseEvent) => {
    if (fromPlugin(event) || !run?.ready || run.failure || draftController.cellInput?.contains(event.target as Node) ||
      draftController.isComposing('cell') || draftController.isComposing('formula')) return;
    draftController.commit('cell'); draftController.commit('formula'); flushNudge();
    const { cell, chart } = point(event);
    clickStart.current = null;
    if (chart) {
      setSelectedChart({ id: chart.id, movable: chart.movable });
      if (editable() && chart.movable && event.button === 0) chartDrag.current = {
        generation: run.generation, sheet: activeRef.current, id: chart.id,
        zoom: sourceRef.current?.painted?.request.zoom ?? zoomRef.current,
        x: event.clientX, y: event.clientY,
      };
      focus(); event.preventDefault(); return;
    }
    setSelectedChart(null);
    const selected = selectionRef.current;
    if (!cell || event.shiftKey && !selected) return;
    const next = event.shiftKey ? extendTo(selected!, cell, limits()) : selectionAt(cell);
    predict(activeRef.current, next);
    command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, next) } });
    clickStart.current = cell; dragging.current = true; focus();
  };
  const onMouseMove = (event: MouseEvent<HTMLDivElement>) => {
    if (fromPlugin(event) || run?.failure) return;
    const drag = chartDrag.current;
    if (drag) {
      setChartOffset({ x: event.clientX - drag.x, y: event.clientY - drag.y });
      event.currentTarget.style.cursor = 'move'; return;
    }
    const { cell, chart } = point(event);
    if (dragging.current && cell && selectionRef.current) {
      const next = extendTo(selectionRef.current, cell, limits());
      predict(activeRef.current, next);
      command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, next) } });
      return;
    }
    const link = cell && sourceRef.current?.painted ? hyperlinkAtCell(sourceRef.current.painted.displayList, cell.row, cell.col) : null;
    event.currentTarget.style.cursor = chart?.movable && editable() ? 'move' : chart || link ? 'pointer' : 'default';
    event.currentTarget.title = chart ? '' : link?.tooltip ?? '';
  };
  const onClick = (event: MouseEvent) => {
    if (fromPlugin(event) || run?.failure) return;
    const { cell } = point(event);
    const start = clickStart.current;
    clickStart.current = null;
    const frame = sourceRef.current?.painted?.displayList;
    if (!cell || !start || !frame || cell.row !== start.row || cell.col !== start.col) return;
    const link = hyperlinkAtCell(frame, cell.row, cell.col);
    if (!link) return;
    const href = safeExternalHyperlink(link);
    if (href) { window.open(href, '_blank', 'noopener,noreferrer'); return; }
    const sheets = run?.session.editState.sheets ?? run?.session.state.sheets ?? [];
    const destination = link.location ? parseHyperlinkLocation(link.location, sheets[activeRef.current]?.name ?? '') : null;
    if (!destination) return;
    const sheet = sheets.findIndex((entry) => entry.name.toLowerCase() === destination.sheetName.toLowerCase());
    if (sheet >= 0) navigate(sheet, selectionAt(destination));
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (fromPlugin(event) || !run?.ready || run.failure || !selectionRef.current || !viewRef.current ||
      draftController.isComposing('formula') || commandForEvent(event)) return;
    const mod = event.metaKey || event.ctrlKey;
    const draft = draftController.getSnapshot().cell;
    if (draft) {
      if (draftController.cellInput || draftController.key('cell', event)) return;
      if (draftController.cellKey(event)) { event.preventDefault(); return; }
      if (event.key === 'Enter') draftController.commit('cell', event.shiftKey ? 'up' : 'down');
      else if (event.key === 'Tab') draftController.commit('cell', event.shiftKey ? 'left' : 'right');
      else if (event.key === 'Escape') draftController.cancel();
      else return;
      event.preventDefault(); return;
    }
    const chart = chartRef.current;
    if (chart) {
      if (event.key === 'Escape') {
        chartDrag.current = null; nudge.current = null; setChartOffset(null); setSelectedChart(null);
        event.preventDefault(); return;
      }
      const delta = !mod && chartKeys[event.key];
      if (delta && editable() && chart.movable) {
        const step = event.shiftKey ? 10 : 1;
        const base = nudge.current ?? { generation: run.generation, sheet: activeRef.current, id: chart.id, dx: 0, dy: 0 };
        nudge.current = { ...base, dx: base.dx + delta[0] * step, dy: base.dy + delta[1] * step };
        setChartOffset({ x: nudge.current.dx * zoomRef.current, y: nudge.current.dy * zoomRef.current });
        if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
        const owner = run;
        nudgeTimer.current = setTimeout(() => {
          if (owner.current && !owner.retiring && ownerRef.current === owner) gestures.current.flushNudge();
        }, 250); event.preventDefault();
      }
      return;
    }
    const owner = run;
    const sheet = activeRef.current;
    const navigationContext = () => {
      if (!owner.current || owner.retiring || owner.failure || ownerRef.current !== owner ||
        activeRef.current !== sheet || chartRef.current || draftController.getSnapshot().cell ||
        !selectionRef.current) return null;
      const basis = owner.session.confirmedBasis;
      return { owner, sheet, selection: selectionRef.current, navigation: navigation.current,
        version: `${basis.epoch}:${basis.version}:${owner.session.state.version}` };
    };
    const afterNavigation = (task: () => void, cancel: () => void) => {
      if (!selectionNavigation.defer(navigationContext(), task, false, cancel)) task();
    };
    if (mod && ['c', 'x', 'v'].includes(event.key.toLowerCase())) {
      const kind = event.key.toLowerCase() === 'c' ? 'copy' : event.key.toLowerCase() === 'x' ? 'cut' : 'paste';
      if (clipboardAllowed(kind)) {
        const access = clipboardAccess(kind);
        const refuse = () => access.reject?.(new XlsxCommandAdmissionError('editor-unavailable'));
        afterNavigation(() => draftController.getSnapshot().cell ? refuse() : clipboard(kind, access), refuse);
      }
      event.preventDefault(); return;
    }
    const request = selectionDataRequest(event);
    if (request) {
      event.preventDefault();
      selectionNavigation.enqueue(event, {
        context: navigationContext,
        limits,
        hidden,
        query: (context, direction) => owner.session.call.selectionData(sheet,
          context.selection.focus.row, context.selection.focus.col, direction),
        apply: (next) => {
          predict(sheet, next);
          reveal(sheet, next.focus);
          command({ kind: 'select', payload: { selection: rangeTarget(sheet, next) } });
          return true;
        },
      });
      return;
    }
    const key = { key: event.key, shiftKey: event.shiftKey, metaKey: event.metaKey, ctrlKey: event.ctrlKey,
      altKey: event.altKey };
    const act = () => {
      const selected = selectionRef.current;
      if (!selected || chartRef.current) return;
      if (draftController.getSnapshot().cell) {
        if (draftController.cellKey(key)) return;
        if (key.key === 'Enter') draftController.commit('cell', key.shiftKey ? 'up' : 'down');
        else if (key.key === 'Tab') draftController.commit('cell', key.shiftKey ? 'left' : 'right');
        else if (key.key === 'Escape') draftController.cancel();
        return;
      }
      const action = selectionKeyReducer(selected, key, limits(), undefined, hidden());
      if (action.type === 'move') {
        predict(activeRef.current, action.selection);
        reveal(activeRef.current, action.selection.focus);
        command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, action.selection) } });
      } else if (action.type === 'startEdit') openCell(action.initialInput);
      else if (action.type === 'clear') clear();
    };
    const type = selectionKeyReducer(selectionRef.current, key, limits(), undefined, hidden()).type;
    if (type === 'none') {
      if (key.key === 'Escape' && selectionNavigation.defer(navigationContext(), act, true)) event.preventDefault();
      return;
    }
    event.preventDefault();
    if (!selectionNavigation.defer(navigationContext(), act, type !== 'clear')) act();
  };

  const sheets = catalog;
  const defaultToolbar = useMemo(() => <PeerlessDefaultToolbar />, []);
  const toolbar = props.showToolbar === false ? null : props.toolbar === undefined ? props.readOnly ? null : defaultToolbar : props.toolbar;
  const renderDock = (placement: DockPlacement) => pluginHost.managed ? <PluginDock host={pluginHost.host}
    placement={placement} activations={pluginHost.activations.filter((activation) => activation.plugin.panel?.placement === placement)}
    available={dockArea} /> : null;
  const painted = sourceRef.current?.painted;
  return <XlsxCommandContext.Provider value={commands.store}><EditorChromeContext.Provider value={true}>
    <PeerlessChromeContext.Provider value={true}>
    <PeerlessWorkerInputs controller={draftController} presentation={presentation} run={run} active={active}
      selectedChart={selectedChart !== null} error={Boolean(error)} readOnly={Boolean(props.readOnly)} navigation={navigation}>
      <div ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
        onKeyDownCapture={(event) => {
          if (!fromPlugin(event) && event.key === 'Escape') { setFormatToken(null); commands.refresh(); }
        }}
        style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%', height: '100%',
          minWidth: 0, color: '#202124', background: '#ffffff', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }}>
        <PeerlessSelectionEffect store={presentation} onChange={applySelectionFormat} />
        {toolbar != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbar}</div>}
        <div ref={dockAreaRef} data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minHeight: 0 }}>
          {renderDock('left')}
          <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
          {isProposalsAvailable() && proposalsPanelOpen && <ProposalsPanel proposals={proposals} staleFor={staleFor}
            style={{ top: 4, right: 8, width: 'min(320px, calc(100% - 16px))', maxHeight: 'min(420px, calc(100% - 12px))' }} />}
          <div ref={scrollRef} data-testid="xlsx-scroll" tabIndex={0} onKeyDown={onKeyDown}
            onMouseDown={onMouseDown} onMouseMove={onMouseMove} onClick={onClick}
            onDoubleClick={(event) => {
              if (fromPlugin(event) || point(event).chart) return;
              const cell = point(event).cell;
              const frame = sourceRef.current?.painted?.displayList;
              if (cell && frame && hyperlinkAtCell(frame, cell.row, cell.col)) return;
              openCell();
            }}
            onMouseLeave={(event) => { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }}
            style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}>
            <PeerlessWorkerExtent store={presentation} zoom={zoom} />
            <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
              <canvas ref={canvasRef} style={{ display: 'block', position: 'absolute', top: 0, left: 0 }} />
              <PeerlessWorkerOverlay store={presentation} active={active} zoom={zoom} selectedChart={selectedChart}
                chartOffset={chartOffset} controller={draftController} run={run} disabled={Boolean(error) || Boolean(props.readOnly)} />
              <PeerlessWorkerPresence store={presentation} run={run} provider={props.collaboration?.provider}
                active={active} sheets={sheets} />
              {pluginHost.managed && <PluginOverlays host={pluginHost.host} activations={pluginHost.activations}
                layerRef={pluginHost.overlayLayerRef} width={(painted?.displayList.width ?? 0) * (painted?.request.zoom ?? zoom)}
                height={(painted?.displayList.height ?? 0) * (painted?.request.zoom ?? zoom)} />}
            </div>
          </div>
          {renderDock('bottom')}
          </div>
          {renderDock('right')}
        </div>
        <PeerlessWorkerAccessibility store={presentation} active={active} sheets={sheets} />
        {refusal && <div data-testid="xlsx-input-refusal" role="alert">
          {commandReason(refusal.code, { translate: t }).message}{refusal.detail ? `: ${refusal.detail}` : ''}
        </div>}
        {error && <div data-testid="xlsx-error" role="alert">{t('editor.openError')}: {error.message}</div>}
        {sheets.length > 0 && <div data-testid="xlsx-sheet-tabs" role="tablist" aria-label={t('editor.sheetTabsLabel')}
          style={{ display: 'flex', gap: 2, padding: '4px 6px', borderTop: '1px solid #e0e0e0', background: '#fafafa', overflowX: 'auto' }}>
          {sheets.map((sheet) => <button key={sheet.id} role="tab" aria-selected={sheet.index === active} disabled={Boolean(error)}
            onClick={() => navigate(sheet.index)}
            style={{ border: 'none', padding: '4px 12px', cursor: 'pointer', background: sheet.index === active ? '#ffffff' : 'transparent',
              borderBottom: sheet.index === active ? `2px solid ${BRAND}` : '2px solid transparent', fontWeight: sheet.index === active ? 600 : 400 }}>
            {sheet.name}
          </button>)}
        </div>}
      </div>
    </PeerlessWorkerInputs>
    </PeerlessChromeContext.Provider>
  </EditorChromeContext.Provider></XlsxCommandContext.Provider>;
}
