import type { NavigationContext } from '../selectionData';
import {
  copiedSelection, createSelectionDataNavigation, MAX_CLEAR_CELLS, mergeAnchorCell, sameSelection, selectionDataRequest, selectionFrameContains, selectionHidden, selectionMerges,
} from '../selectionData';
import {
  cellAtPoint, cellRect, chartRegionAtPoint, extendTo, hyperlinkAtCell, isProposalsAvailable,
  normalizeRange, paintDisplayList, parseHyperlinkLocation, safeExternalHyperlink, selectionAt,
  selectionKeyReducer, toTsv,
} from '@betteroffice/xlsx';
import type {
  CellAddr, CellInputEdit, CellRange, Proposal, Selection, WorkbookSheetView, WorkbookSheetSummary, XlsxRangeTarget,
} from '@betteroffice/xlsx';
import type { WorkbookCommandSelection } from '@betteroffice/xlsx/worker';
import { editorExportFaces, useDefaultFonts } from '../useDefaultFonts';
import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useReducer, useRef, useState } from 'react';
import type { KeyboardEvent, MouseEvent, SetStateAction, SyntheticEvent } from 'react';
import { flushSync } from 'react-dom';
import { createXlsxCommandController, XlsxCommandAdmissionError, XlsxHostRefusalError } from '../commands/createXlsxCommandStore';
import { commandForEvent, isMacPlatform } from '../commands/descriptors';
import { useContextMenu } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { xlsxContextMenuEntries } from '../contextMenu/menus';
import { commandReason, initialZoomScale } from '../commands/evaluate';
import type { XlsxCommandFailureCode } from '../commands/types';
import type { InputDraft, InputSeal } from '../commands/inputCoordinator';
import { inPluginChrome, pluginEventStore } from '../commands/pluginEvents';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { useSaveRequest } from '../commands/useSaveRequest';
import { usePeerlessWorkerXlsxCommands } from '../commands/usePeerlessWorkerXlsxCommands';
import { WorkerInputRefusal } from '../commands/workerInputErrors';
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
import {
  changedEdits, coverRanges, formulaMoves, GATE_CELL_LIMIT, inputRanges, withGatedInputs, writtenInputs,
} from '../plugins/editGate';
import type { XlsxPluginCellCommit, XlsxPluginCellRef, XlsxPluginNavigationFailureCode } from '../plugins/types';
import { PluginOverlays } from '../plugins/PluginOverlays';
import { PluginRowGutter } from '../plugins/PluginRowGutter';
import { ROW_GUTTER_WIDTH } from '../plugins/rowDecorations';
import { PluginDock, useDockArea, type DockPlacement } from '../plugins/PluginPanels';
import { useXlsxPluginHost, type XlsxPluginHostBinding } from '../plugins/useXlsxPluginHost';
import { ProposalsPanel } from '../proposals/ProposalsPanel';
import { cellInViewport, deriveLimits, extendScrollExtent, nearestCellScroll, positionAtPoint } from '../viewer/sessionGeometry';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { XlsxAsyncExportBridge } from './createAsyncEditorApi';
import { PeerlessWorkbookSession, type PeerlessSessionWorkbookProps } from './usePeerlessSessionWorkbook';
import { usePeerlessSessionWorkbook } from './usePeerlessSessionWorkbook';
import {
  PeerlessWorkerPaintSource, type WorkerFramePaint, type WorkerPaintRequest, type WorkerPaintResult,
} from './PeerlessWorkerPaintSource';

import { CellClipboard, clipboardHtml, newClipboardMarker, parseCellClipboardText, readCellClipboard, type CellPaste } from '../clipboard';
import { WorkerCellEcho } from './WorkerCellEcho';
import { WorkerDraftController, type WorkerNativeDraft } from './WorkerDraftController';
import { PeerlessWorkerInputs } from './PeerlessWorkerInputs';
import { WorkerPresentationStore } from './WorkerPresentationStore';
import {
  PeerlessSelectionEffect, PeerlessWorkerAccessibility, PeerlessWorkerExtent, PeerlessWorkerOverlay,
  PeerlessWorkerPresence,
} from './PeerlessWorkerSurface';
import { WorkbookFailure } from '../components/WorkbookFailure';

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
  const anchor = target.activeCell ?? start;
  const range = previous ? normalizeRange(previous) : null;
  if (previous && previous.anchor.row === anchor.row && previous.anchor.col === anchor.col &&
    range?.top === start.row && range.left === start.col &&
    range.bottom === end.row && range.right === end.col) return previous;
  if ((anchor.row !== start.row && anchor.row !== end.row) ||
    (anchor.col !== start.col && anchor.col !== end.col)) {
    return { ...selectionAt(anchor), range: { top: start.row, left: start.col, bottom: end.row, right: end.col } };
  }
  return { anchor: { ...anchor }, focus: {
    row: anchor.row === start.row ? end.row : start.row,
    col: anchor.col === start.col ? end.col : start.col,
  } };
}

const MAX_CLEAR_SCAN_CELLS = 10 * MAX_CLEAR_CELLS;

function rangeA1(range: CellRange): string {
  return `${address({ row: range.top, col: range.left })}:${address({ row: range.bottom, col: range.right })}`;
}

function handledRejection(error: unknown): Promise<never> {
  const rejected = Promise.reject(error);
  rejected.catch(() => {});
  return rejected;
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

const NEIGHBOR_FRAME_DELAY_MS = 400;
const NEIGHBOR_FRAME_EDIT_DELAY_MS = 1500;
const NEIGHBOR_FRAME_BUDGET_MS = 24;

interface NeighborFrame {
  sheetId: string;
  proposals: string;
  request: WorkerPaintRequest;
  paint: WorkerFramePaint;
}

function sameFrameRequest(left: WorkerPaintRequest, right: WorkerPaintRequest): boolean {
  return left.generation === right.generation && left.sheet === right.sheet && left.zoom === right.zoom &&
    left.dpr === right.dpr && left.width === right.width && left.height === right.height &&
    left.viewport.x === right.viewport.x && left.viewport.y === right.viewport.y &&
    left.viewport.width === right.viewport.width && left.viewport.height === right.viewport.height;
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
      <EditorToolbar.FormulaBar />
      {isProposalsAvailable() && <ToolbarCommandButton id="proposalsPanel" />}
    </div>
  </EditorToolbar>;
});

export function XlsxPeerlessWorkerEditor(props: PeerlessSessionWorkbookProps) {
  const { t } = useTranslation();
  const [cellClipboard] = useState(() => new CellClipboard());
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pluginBindingRef = useRef<XlsxPluginHostBinding | null>(null);
  const gateCommitRef = useRef<(draft: InputDraft, target: XlsxRangeTarget) => (() => Promise<string>) | null>(() => null);
  useLayoutEffect(() => () => pluginBindingRef.current?.host.close('unmounted'), []);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [commands] = useState(createXlsxCommandController);
  const [active, setActive, activeRef] = useSyncedState(0);
  const activeSheetId = useRef<string | undefined>(undefined);
  const [presentation] = useState(() => new WorkerPresentationStore());
  const { setSelection, selectionRef, setView, viewRef, setPainted } = presentation;
  const [catalog, setCatalog, catalogRef] = useSyncedState<WorkbookSheetSummary[]>([]);
  const [zoom, setZoom, zoomRef] = useSyncedState(initialZoomScale(props.initialZoom));
  const [refusal, setRefusal] = useState<{ code: XlsxCommandFailureCode; detail?: string; message?: string } | null>(null);
  const [selectedChart, setSelectedChart, chartRef] = useSyncedState<{ id: string; movable: boolean } | null>(null);
  const [chartOffset, setChartOffset] = useState<{ x: number; y: number } | null>(null);
  const [renaming, setRenaming] = useState<number | null>(null);
  const tabsRef = useRef<HTMLDivElement>(null);
  const pendingActivation = useRef<string | null>(null);
  const [, bumpActivation] = useReducer((count: number) => count + 1, 0);
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
  const emptySelection = useRef(selectionAt({ row: 0, col: 0 }));
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
  const nameNavigation = useRef<(cell: CellAddr) => void>(() => {});
  const sheetNavigation = useRef<(sheet: number) => void>(() => {});
  const sheetLayout = useRef(0);
  const navigateFromNameBox = useCallback((cell: CellAddr) => nameNavigation.current(cell), []);
  const exportInput = useRef<ReturnType<typeof createPeerlessWorkerInputCoordinator> | null>(null);
  const [bridge] = useState<WorkerEditorApiBridge & XlsxAsyncExportBridge>(() => ({
    coordinator: () => null, readOnly: () => propsRef.current.readOnly ?? false,
    allowHostProposals: () => propsRef.current.allowHostProposals ?? false, focus,
    exportCoordinator: () => exportInput.current, defaultFonts: () => propsRef.current.defaultFonts,
    editorDocument: () => pluginAccess, saved: (bytes) => propsRef.current.onSave?.(bytes),
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
  const paintWithDefaultFonts = useDefaultFonts(run, run?.session.call.decodeFontFace, props.defaultFonts,
    () => { if (run?.current && !run.failure) sourceRef.current?.schedule(); },
    () => { if (canvasRef.current) canvasRef.current.dataset.defaultFonts = 'settled'; });
  const sheetSelections = useMemo(() => new Map<string, Selection>(), [run]);
  const sheetScrolls = useMemo(() => new Map<string, { x: number; y: number }>(), [run]);
  const neighborFrames = useMemo(() => new Map<number, NeighborFrame>(), [run]);
  const neighborTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const neighborRevision = useRef(0);
  const neighborVersion = useRef<string | undefined>(undefined);
  const costlyNeighbors = useMemo(() => new Set<string>(), [run]);
  const scheduleNeighborFrames = useRef(() => {});
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
    if (error instanceof XlsxHostRefusalError) {
      setRefusal({ code: error.code, message: error.message });
    } else if (error instanceof PeerlessWorkerCommandRefusal && error.status.reason === 'read-only') {
      setRefusal({ code: 'read-only' });
    } else if (error instanceof WorkerInputRefusal) {
      setRefusal({ code: error.code, detail: error.code === 'sheet-protected' ? undefined : error.message });
    } else if (error instanceof XlsxCommandAdmissionError) setRefusal({ code: error.code });
    else setRefusal({ code: 'input-failed', detail: error instanceof Error ? error.message : String(error) });
  }, []);
  const rangeTarget = (sheet: number, selected: Selection): WorkbookCommandSelection => {
    const sheetId = ownerRef.current?.session.editState.sheets?.find((entry) => entry.index === sheet)?.id ??
      ownerRef.current?.session.state.sheets.find((entry) => entry.index === sheet)?.id;
    if (!sheetId) throw new XlsxCommandAdmissionError('editor-unavailable');
    const range = normalizeRange(selected);
    return { sheetId, range: { kind: 'rowCol', start: { row: range.top, col: range.left },
      end: { row: range.bottom, col: range.right } }, activeCell: { ...selected.anchor } };
  };
  const selectedTarget = () => {
    const owner = ownerRef.current;
    const sheets = owner?.session.editState.sheets ?? owner?.session.state.sheets;
    return selectionRef.current && sheets?.some((sheet) => sheet.index === activeRef.current)
      ? rangeTarget(activeRef.current, selectionRef.current) : undefined;
  };
  const hidden = () => selectionHidden(sourceRef.current?.painted?.displayList.grid);
  const navigationContext = () => {
    const owner = ownerRef.current;
    if (!owner?.current || owner.retiring || owner.failure) return null;
    const basis = owner.session.confirmedBasis;
    return { owner, sheet: activeRef.current, selection: selectionRef.current ?? emptySelection.current, navigation: navigation.current,
      version: `${basis.epoch}:${basis.version}:${owner.session.state.version}` };
  };
  const merges = (context: NavigationContext | null = navigationContext()) => {
    const painted = sourceRef.current?.painted;
    const basis = ownerRef.current?.session.confirmedBasis;
    const visible = painted?.request.sheet === context?.sheet && painted?.version === basis?.version
      ? selectionMerges(painted?.mergedRanges) : undefined;
    return context ? selectionNavigation.merges(context, visible) : selectionMerges(undefined);
  };
  const limits = () => deriveLimits(sourceRef.current?.painted?.displayList ?? null,
    (scrollRef.current?.clientHeight ?? 0) / zoomRef.current);
  const placeSelection = (sheet: number, next: Selection) => {
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
  const geometryKnown = (context: NavigationContext, cell: CellAddr, stable = false) => {
    const painted = sourceRef.current?.painted;
    const basis = ownerRef.current?.session.confirmedBasis;
    if (!stable && coordinator.pending && !coordinator.pendingSelectionGeometryStable) return false;
    if (selectionNavigation.known(context, cell)) return true;
    if (!painted || painted.request.sheet !== context.sheet || painted.version !== basis?.version) return false;
    return selectionFrameContains(painted.displayList.grid, cell);
  };
  const navigationOptions = () => ({
    context: navigationContext,
    activeCell: (selected: Selection) => selected.anchor,
    limits, hidden, merges,
    known: geometryKnown,
    query: (context: { sheet: number; selection: Selection }, direction: import('@betteroffice/xlsx').Direction | 'usedRange' | 'rowEnd' | 'merge' | undefined) =>
      ownerRef.current!.session.call.selectionData(context.sheet, context.selection.focus.row, context.selection.focus.col, direction),
    apply: (next: Selection, sheet = activeRef.current) => {
      placeSelection(sheet, next);
      reveal(sheet, next.focus);
      command({ kind: 'select', payload: { selection: rangeTarget(sheet, next) } });
      return true;
    },
  });
  const predict = (sheet: number, next: Selection, replace = false, stable = false) => {
    if (sheet === activeRef.current && sameSelection(selectionRef.current, next)) return;
    void selectionNavigation.select(next, { ...navigationOptions(), known: (context, cell) => geometryKnown(context, cell, stable), apply: (resolved) => {
      placeSelection(sheet, resolved); return true;
    } }, sheet, replace);
  };
  const reveal = (sheet: number, cell: CellAddr, previousZoom?: number) => {
    const owner = ownerRef.current;
    const frame = sourceRef.current?.painted;
    if (!owner?.current || owner.failure) return;
    const scroll = scrollRef.current;
    const rect = frame?.geometry && frame.request.sheet === sheet ? cellRect(frame.geometry, cell.row, cell.col) : null;
    const scale = zoomRef.current;
    const padding = 2 / scale;
    if (previousZoom === undefined && rect && scroll && frame?.request.zoom === scale &&
      frame.request.viewport.x === scroll.scrollLeft / scale && frame.request.viewport.y === scroll.scrollTop / scale &&
      (rect.x === 0 || rect.x >= padding) && (rect.y === 0 || rect.y >= padding) &&
      rect.x + rect.w + padding <= scroll.clientWidth / scale &&
      rect.y + rect.h + padding <= scroll.clientHeight / scale) return;
    const at = navigation.current;
    if (revealing.current?.generation === owner.generation && revealing.current.navigation === at) return;
    const request = { generation: owner.generation, navigation: at };
    revealing.current = request;
    const layout = sheetLayout.current;
    void owner.session.call.cellGeometry(sheet, cell.row, cell.col).then((geometry) => {
      const scroll = scrollRef.current;
      if (!owner.current || ownerRef.current !== owner || navigation.current !== at || sheetLayout.current !== layout || !scroll) return;
      if (previousZoom !== undefined && !cellInViewport(scroll, geometry.rect, previousZoom)) return;
      const view = viewRef.current;
      const id = sheetId(sheet);
      if (!view || view.sheet !== sheet) {
        if (id) sheetScrolls.set(id, geometry.scrollPosition);
        scroll.scrollLeft = geometry.scrollPosition.x * zoomRef.current;
        scroll.scrollTop = geometry.scrollPosition.y * zoomRef.current;
        sourceRef.current?.schedule();
        return;
      }
      const scale = zoomRef.current;
      const padding = 2 / scale;
      const extended = extendScrollExtent(scroll, view, {
        ...geometry.rect, w: geometry.rect.w + padding, h: geometry.rect.h + padding,
      }, scale);
      revealedExtent.current = { sheet, contentWidth: extended.contentWidth, contentHeight: extended.contentHeight };
      setView(extended);
      const target = nearestCellScroll(geometry.rect, cell, view, {
        x: scroll.scrollLeft / scale, y: scroll.scrollTop / scale,
        width: scroll.clientWidth / scale, height: scroll.clientHeight / scale,
      }, padding);
      if (id) sheetScrolls.set(id, target);
      scroll.scrollLeft = target.x * scale;
      scroll.scrollTop = target.y * scale;
      sourceRef.current?.schedule();
    }).catch((error) => {
      if (owner.current && ownerRef.current === owner && navigation.current === at && sheetLayout.current === layout) owner.fail(error);
    }).finally(() => { if (revealing.current === request) revealing.current = null; });
  };
  const revealedZoom = useRef(zoom);
  useLayoutEffect(() => {
    if (revealedZoom.current === zoom) return;
    const previousZoom = revealedZoom.current;
    revealedZoom.current = zoom;
    const selected = selectionRef.current;
    if (selected && !chartRef.current) reveal(activeRef.current, selected.anchor, previousZoom);
  }, [zoom]);
  const requests = useMemo(() => new Map<number, WorkerPaintRequest>(), [run]);
  const pipeCommands = useMemo(() => new Set<number>(), [run]);
  const localPaintSequence = useMemo(() => ({ current: 0 }), [run]);
  const selectionCommands = useMemo(() => new Set<number>(), [run]);
  const selectionRevisions = useMemo(() => new Map<number, number>(), [run]);
  const issued = useMemo(() => new Map<number, PeerlessWorkerCommandInput>(), [run]);
  const inputEvents = useRef<(event: PeerlessWorkerInputEvent, publish?: () => void) => void>(() => {});
  const coordinatorRef = useRef<ReturnType<typeof createPeerlessWorkerInputCoordinator> | null>(null);
  const draftController: WorkerDraftController = useMemo(() => new WorkerDraftController({
    defer: (task, pendingOnly) => pendingOnly || draftController.ownsFocus
      ? selectionNavigation.defer(navigationContext(), task, pendingOnly && !draftController.ownsFocus) :
      selectionNavigation.beforeInput({ ...navigationOptions(), apply: (next) => {
        placeSelection(activeRef.current, next); return true;
      } }, task, () => report(new XlsxCommandAdmissionError('editor-unavailable'))),
    restoreCell: (sheet, cell, apply) => {
      void selectionNavigation.select(selectionAt(cell), { ...navigationOptions(), apply: (next) => {
        placeSelection(sheet, next);
        apply(next.anchor);
        reveal(sheet, next.anchor);
        return true;
      } }, sheet);
    },
    generation: () => run?.generation ?? 0,
    identity: () => `${run?.generation}:${run?.session.commandIdentity.session}:${run?.session.commandIdentity.incarnation}:${run?.session.confirmedBasis.epoch}`,
    valueRevision: () => run?.session.confirmedBasis.version ?? '',
    navigation: () => navigation.current, sheet: () => activeRef.current,
    selection: () => {
      const selected = selectionRef.current;
      const anchor = selected && mergeAnchorCell(selected.anchor, merges());
      return selected && anchor !== selected.anchor ? { ...selected, anchor: anchor! } : selected;
    },
    target: (sheet, row, col) => rangeTarget(sheet, selectionAt({ row, col })),
    resolveSheet: (sheetId) => (run?.session.editState.sheets ?? run?.session.state.sheets)?.find((sheet) => sheet.id === sheetId)?.index,
    editable,
    prefill: (sheet, cell) => readWorkerCellInput(run!.session, sheet, address(cell)),
    select: (sheet, cell, move) => {
      const base = selectionAt(cell);
      if (sheet !== activeRef.current || !sameSelection(selectionRef.current, base)) predict(sheet, base);
      if (move) {
        let immediate = true;
        selectionNavigation.enqueue({ key: move === 'down' || move === 'up' ? 'Enter' : 'Tab',
          shiftKey: move === 'up' || move === 'left' }, { ...navigationOptions(), apply: (next) => {
          placeSelection(sheet, next); reveal(sheet, next.focus);
          if (!immediate) command({ kind: 'select', payload: { selection: rangeTarget(sheet, next) } });
          return true;
        } });
        immediate = false;
      }
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
    allowHostProposals: () => propsRef.current.allowHostProposals ?? false,
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
    gateCommit: (draft, target) => gateCommitRef.current(draft, target),
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
    readOnly: () => propsRef.current.readOnly ?? false,
    commands: () => commands,
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
    readOnly: props.readOnly ?? false, handle: pluginIdentity, selection: pluginSelection, proposals, staleFor,
    proposalsPending: () => coordinator.pending || proposalsPending.current !== null, canvasRef, t,
    positionAtPoint: (clientX, clientY) => {
      const owner = ownerRef.current;
      const frame = sourceRef.current?.painted;
      if (!owner?.current || owner.retiring || owner.failure || !frame || frame.request.generation !== owner.generation) return null;
      return positionAtPoint(frame.displayList, frame.request.sheet, canvasRef.current, clientX, clientY);
    } });
  pluginBindingRef.current = pluginHost;
  const editGate = pluginHost.host;

  /** Whether plugins let the user type into a cell of the active sheet; a protected sheet leaves it to protection. */
  const typingAllowed = useCallback((cell: CellAddr): boolean => {
    const owner = ownerRef.current;
    if (!owner?.ready || !editGate.gatesEdits() || owner.session.editState.selectionEditable === false) return true;
    const sheetId = owner.session.editState.sheets?.find((entry) => entry.index === activeRef.current)?.id ??
      owner.session.state.sheets.find((entry) => entry.index === activeRef.current)?.id;
    return !sheetId || editGate.openGate().cellEditable({ sheetId, row: cell.row, col: cell.col });
  }, [editGate]);

  /** Whether sheet protection refuses writing `ranges` of a protected sheet, which it does before plugins are asked. */
  const lockedBySheet = async (owner: PeerlessWorkbookSession, sheetId: string, ranges: readonly CellRange[]) => {
    const catalog = await owner.session.call.readCells({ ranges: [] });
    const entry = catalog.ok ? catalog.sheets.find((sheet) => sheet.sheetId === sheetId) : undefined;
    if (!catalog.ok || !entry || entry.editable) return false;
    const check = await owner.session.call.validateEdits({ expectVersion: catalog.version,
      steps: ranges.map(({ top, left, bottom, right }) => ({ op: 'setCellInputs' as const,
        inputs: Array.from({ length: bottom - top + 1 }, () => Array<string>(right - left + 1).fill('')),
        target: { sheetId, range: { kind: 'rowCol' as const, start: { row: top, col: left }, end: { row: bottom, col: right } } } })) });
    return !check.ok && check.failure.code === 'locked-target';
  };
  /** Asks the plugins about `edits` in one pass: the inputs to write, or a throw refusing them all. */
  const askPlugins = (sheetId: string, edits: readonly CellInputEdit[], source: XlsxPluginCellCommit['source']) => {
    const cells: XlsxPluginCellRef[] = edits.map(({ row, col }) => ({ sheetId, row, col }));
    if (edits.length > 0 && editGate.editGateStarting()) throw new XlsxHostRefusalError(t('commands.reasons.hostStarting'));
    const refusal = (message: string | null) => new XlsxHostRefusalError(message ?? t('commands.reasons.hostRefused'));
    const pass = editGate.openGate();
    if (!cells.every((cell) => pass.cellEditable(cell))) throw refusal(null);
    return edits.map((edit, index) => {
      const outcome = pass.beforeCellCommit({ ...cells[index], input: edit.input, source });
      if (!outcome.ok) throw refusal(outcome.message);
      return { ...edit, input: outcome.input };
    });
  };
  /** Runs typed writes through the plugins once the worker confirmed protection does not refuse them first. */
  const gateWrites = async (
    owner: PeerlessWorkbookSession, sheetId: string, edits: readonly CellInputEdit[],
    source: XlsxPluginCellCommit['source']
  ): Promise<CellInputEdit[]> => {
    if (edits.length === 0) return [...edits];
    if (await lockedBySheet(owner, sheetId, edits.map(({ row, col }) => ({ top: row, bottom: row, left: col, right: col })))) {
      return [...edits];
    }
    return askPlugins(sheetId, edits, source);
  };
  /**
   * A cut through the plugins: its cells holding content are asked as cleared. Every refusal,
   * sheet protection's first, throws before the clipboard is written.
   */
  const gateCut = async (
    owner: PeerlessWorkbookSession, sheetId: string, range: CellRange, cells: readonly (readonly { input: string }[])[]
  ): Promise<CellInputEdit[]> => {
    const content = cells.flatMap((row, dr) => row.flatMap(({ input }, dc) =>
      input === '' ? [] : [{ row: range.top + dr, col: range.left + dc, input: '' }]));
    if (content.length === 0) return content;
    if (content.length > GATE_CELL_LIMIT) {
      throw new XlsxHostRefusalError(t('commands.reasons.tooManyToCheck', { limit: GATE_CELL_LIMIT }));
    }
    if (await lockedBySheet(owner, sheetId, coverRanges(content))) {
      throw new WorkerInputRefusal(t('commands.reasons.sheetProtected'), false, 'sheet-protected');
    }
    return askPlugins(sheetId, content, 'cut');
  };
  /**
   * A paste through the plugins: each cell it changes is asked once, with the input written there,
   * after protection had its say; any refusal refuses the whole paste.
   */
  const gatePaste = async (
    owner: PeerlessWorkbookSession, sheet: number, sheetId: string, edits: CellInputEdit[]
  ): Promise<CellInputEdit[]> => {
    if (edits.length === 0 || !editGate.gatesEdits() || await lockedBySheet(owner, sheetId, coverRanges(edits))) return edits;
    const moves = formulaMoves(edits);
    const ranges = inputRanges(edits);
    const [moved, read] = await Promise.all([
      moves.items.length ? owner.session.call.translateFormulas(moves.items) : [],
      Promise.all(ranges.map(async (range) => (await owner.session.call.cellInputs(sheet, rangeA1(range))).cells)),
    ]);
    const changed = changedEdits(edits, writtenInputs(edits, moves, moved), ranges, read);
    if (changed.length > GATE_CELL_LIMIT) {
      throw new XlsxHostRefusalError(t('commands.reasons.tooManyToCheck', { limit: GATE_CELL_LIMIT }));
    }
    return withGatedInputs(edits, changed, askPlugins(sheetId, changed, 'paste'));
  };
  /**
   * The cells of `range` that hold content, read inside the used range in bounded chunks; null when
   * more than a clear may change hold content or the area to read is too large.
   */
  const contentCells = async (owner: PeerlessWorkbookSession, sheet: number, range: CellRange): Promise<CellInputEdit[] | null> => {
    const { usedRange: used } = await owner.session.call.selectionData(sheet, 0, 0, 'usedRange');
    if (!used) return [];
    const top = Math.max(range.top, used.top), bottom = Math.min(range.bottom, used.bottom);
    const left = Math.max(range.left, used.left), right = Math.min(range.right, used.right);
    if (top > bottom || left > right) return [];
    const width = right - left + 1;
    if ((bottom - top + 1) * width > MAX_CLEAR_SCAN_CELLS) return null;
    const rows = Math.max(1, Math.floor(MAX_CLEAR_CELLS / width));
    const content: CellInputEdit[] = [];
    for (let start = top; start <= bottom; start += rows) {
      const end = Math.min(bottom, start + rows - 1);
      const { cells } = await owner.session.call.cellInputs(sheet, rangeA1({ top: start, left, bottom: end, right }));
      cells.forEach((row, dr) => row.forEach((cell, dc) => {
        if (cell.input !== '') content.push({ row: start + dr, col: left + dc, input: '' });
      }));
      if (content.length > MAX_CLEAR_CELLS) return null;
    }
    return content;
  };
  gateCommitRef.current = (draft, target) => {
    const owner = ownerRef.current;
    if (!owner || !editGate.gatesEdits()) return null;
    return () => gateWrites(owner, target.sheetId, [{ row: draft.row, col: draft.col, input: draft.value }],
      draft.source === 'cell' ? 'cell-editor' : 'formula-bar').then(([edit]) => edit.input);
  };
  const [dockAreaRef, dockArea] = useDockArea(pluginHost.managed);
  useLayoutEffect(() => {
    const apply = async ([update, ...updates]: Uint8Array[]) => {
      const status = await coordinator.command({ kind: 'applyUpdate',
        payload: { update: update!, ...updates.length ? { updates } : {} } }, { after: null }).result;
      if (status.status !== 'applied' && status.status !== 'noop') throw new Error(status.reason ?? 'Remote update refused');
      const refused = (status.result as { refusedUpdate?: string } | undefined)?.refusedUpdate;
      if (refused) throw new Error(refused);
      return status;
    };
    return run?.session.bindCollaboration?.((update) => apply([update]), apply, async (payload) => {
      const status = await coordinator.command({ kind: 'collabPipe', payload }, { after: null }).result;
      if (status.status !== 'applied' && status.status !== 'noop') throw new Error(status.reason ?? 'Collaboration pipe refused');
      return status;
    });
  }, [run, coordinator]);
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
  const revealedExtent = useRef<{ sheet: number; contentWidth: number; contentHeight: number } | null>(null);
  const refreshSheetGeometry = useCallback((paint = false) => {
    const owner = ownerRef.current;
    if (!owner?.ready || owner.failure) return;
    const sheet = activeRef.current;
    const revision = ++geometryRevision.current;
    void owner.session.call.sheetView(sheet).then((next) => {
      if (ownerRef.current !== owner || !owner.current || owner.failure ||
        revision !== geometryRevision.current || activeRef.current !== next.sheet) return;
      const changedSheet = viewRef.current?.sheet !== next.sheet;
      const revealed = revealedExtent.current;
      if (revealed?.sheet !== next.sheet) revealedExtent.current = null;
      const view = revealed?.sheet === next.sheet ? {
        ...next,
        contentWidth: Math.max(next.contentWidth, revealed.contentWidth),
        contentHeight: Math.max(next.contentHeight, revealed.contentHeight),
      } : next;
      setView(view);
      if (changedSheet) restoreSheetScroll(view);
      if (paint || changedSheet) sourceRef.current?.schedule();
    }).catch((error) => {
      if (ownerRef.current === owner && owner.current && !owner.failure && revision === geometryRevision.current) owner.fail(error);
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
      if (command.kind === 'collabPipe') pipeCommands.add(event.seq);
      const revision = command.kind === 'editCells' &&
        JSON.stringify(command.payload.selection) === JSON.stringify(selectedTarget())
        ? navigation.current : event.selectionRevision ?? navigation.current;
      selectionRevisions.set(event.seq, revision);
      while (selectionRevisions.size > 256) selectionRevisions.delete(selectionRevisions.keys().next().value!);
      if (command.frame) {
        if (command.kind !== 'collabPipe') localPaintSequence.current = event.seq;
        requests.set(event.seq, { ...command.frame as WorkerPaintRequest,
          navigation: revision, coversThroughSeq: event.seq });
      }
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
      if (event.status?.committed) {
        cellClipboard.changed();
        setRefusal(null);
      }
      issued.delete(event.seq);
      const refreshed = command?.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'refreshProposals';
      if (command && !refreshed && command.kind !== 'select' && command.kind !== 'setActiveSheet' &&
        command.kind !== 'setReadOnly' && command.kind !== 'applyUpdate' && command.kind !== 'collabPipe' &&
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
        report(event.error ?? new PeerlessWorkerCommandRefusal(event.status ?? { seq: event.seq ?? 0, status: 'refused' }, event.target, event.draft));
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
    let remoteUpdate = false;
    for (const status of ack.statuses) {
      const result = status.result as { remoteUpdate?: boolean; proposals?: Proposal[] } | undefined;
      if (!result?.remoteUpdate) continue;
      remoteUpdate = true;
      cellClipboard.changed();
      proposalsMutation.current = Math.max(proposalsMutation.current, status.seq);
      if (result.proposals) {
        setProposals(result.proposals);
        proposalsLanded.current = { owner: run, revision: proposalsRevision.current, sequence: status.seq };
      }
      commands.refresh();
    }
    const latestSelection = Math.max(0, ...selectionCommands);
    const reconcile = !selectionNavigation.isPending(navigationContext()) && latestSelection > 0 && ack.throughSeq >= latestSelection &&
      selectionRevisions.get(latestSelection) === navigation.current &&
      !draftController.ownsFocus;
    if (ack.editState?.sheets && JSON.stringify(ack.editState.sheets) !== JSON.stringify(catalogRef.current)) {
      const sheets = ack.editState.sheets;
      setCatalog(sheets);
      const kept = sheets.find((entry) => entry.id === activeSheetId.current);
      if (!kept || kept.index !== activeRef.current) sheetLayout.current += 1;
      if (kept && kept.index !== activeRef.current) {
        const previous = activeRef.current;
        geometryRevision.current += 1;
        if (viewRef.current?.sheet === previous) setView({ ...viewRef.current, sheet: kept.index });
        if (revealedExtent.current?.sheet === previous) revealedExtent.current = { ...revealedExtent.current, sheet: kept.index };
        setActive(kept.index);
        sourceRef.current?.schedule();
      } else if (!kept && sheets.length > 0) {
        const acknowledged = reconcile && ack.selection ? ack.selection : undefined;
        const fallback = (acknowledged && sheets.find((entry) => entry.id === acknowledged.sheetId)) ??
          sheets.find((entry) => entry.index === run.session.editState.activeSheet) ??
          sheets.find((entry) => entry.index === Math.min(activeRef.current, sheets.length - 1)) ?? sheets[0]!;
        const restored = acknowledged?.sheetId === fallback.id ? commandSelection(acknowledged, selectionRef.current) : null;
        placeSelection(fallback.index, restored ?? sheetSelection(fallback.index));
        queueMicrotask(() => sheetNavigation.current(fallback.index));
      }
    }
    if (ack.selection && reconcile) {
      const target = ack.selection;
      const sheet = run.session.editState.sheets?.find((entry) => entry.id === target.sheetId)?.index;
      const restored = commandSelection(target, selectionRef.current);
      if (sheet !== undefined && restored) predict(sheet, restored);
    } else if (!ack.selection && ack.editState?.activeSheet !== undefined && reconcile && ack.editState.activeSheet !== activeRef.current) {
      predict(ack.editState.activeSheet, sheetSelection(ack.editState.activeSheet));
    }
    publish?.();
    const source = sourceRef.current;
    const frame = ack.frame;
    const frameRequest = frame && requests.get(frame.coversThroughSeq);
    let handled = false;
    if (frame && frameRequest && source) {
      const request = reconcile ? { ...frameRequest, navigation: navigation.current } : frameRequest;
      const localPaint = localPaintSequence.current > source.coversThroughSeq;
      handled = pipeCommands.has(frame.coversThroughSeq) && !localPaint
        ? source.queuePipeFrame(frame, request) : source.install(frame, request);
    }
    let pipe = false;
    for (const seq of pipeCommands) {
      if (seq > (source?.coversThroughSeq ?? 0) && seq <= ack.throughSeq) { pipe = true; break; }
    }
    const covered = pipe && !ack.frame && source?.advanceCoverage(ack.version, ack.throughSeq);
    if (!handled && !covered && (ack.frame || remoteUpdate ||
      pipe && ack.throughSeq >= coordinator.submittedThroughSeq)) source?.schedule();
    for (const seq of pipeCommands) if (seq <= (source?.coversThroughSeq ?? 0)) pipeCommands.delete(seq);
    for (const seq of requests.keys()) if (seq <= (sourceRef.current?.coversThroughSeq ?? 0)) requests.delete(seq);
    for (const seq of selectionCommands) if (seq <= ack.throughSeq) selectionCommands.delete(seq);
    for (const seq of selectionRevisions.keys()) if (seq <= ack.throughSeq) selectionRevisions.delete(seq);
    if (causalSelection.current !== null && causalSelection.current <= ack.throughSeq) causalSelection.current = null;
    if (ack.statuses.some((status) => status.committed && status.status === 'applied')) {
      mutation.current += 1;
      refreshSheetGeometry();
      propsRef.current.onDocumentChange?.({ version: ack.version });
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
    let cell = target?.selection?.anchor ?? { row: target?.row!, col: target?.col! };
    const valid = (value: CellAddr | undefined) => value && Number.isInteger(value.row) && Number.isInteger(value.col) &&
      value.row >= 0 && value.row < 1_048_576 && value.col >= 0 && value.col < 16_384;
    if (!sheet || !valid(cell) || target.selection && !valid(target.selection.focus)) return 'missing-target';
    let pendingNavigation: typeof pluginNavigation.current = null;
    try {
      const context = navigationContext();
      if (target.selection && context) {
        let resolved: Selection | undefined;
        const accepted = await selectionNavigation.select(target.selection, { ...navigationOptions(), apply: (next) => {
          if (!owner.current || owner.failure || !live()) return false;
          resolved = next;
          placeSelection(sheet.index, next);
          return true;
        } }, sheet.index);
        if (!accepted || !resolved || !owner.current || owner.failure || !live()) return 'unsupported';
        target = { ...target, selection: resolved };
        cell = resolved.anchor;
      }
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
          await new Promise<void>((done) => {
            if (!selectionNavigation.defer(navigationContext(), done, false, done)) done();
          });
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
        if (source && seq && painted && !painted.provisional && painted.version === expectVersion && request &&
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
    revealedExtent.current = null;
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
        canvas.dataset.defaultFonts = paintWithDefaultFonts(frame.displayList,
          () => paintDisplayList(context, frame.displayList, request.dpr * request.zoom)) ? 'settled' : 'pending';
        canvas.dataset.workerSequence = String(frame.sequence);
        canvas.dataset.commandCoverage = String(frame.coversThroughSeq ?? 0);
        canvas.dataset.paintSource = frame.source; canvas.dataset.workerZoom = String(request.zoom);
      },
      onInstalled(frame) {
        for (const seq of pipeCommands) if (seq <= source.coversThroughSeq) pipeCommands.delete(seq);
        for (const seq of requests.keys()) if (seq <= source.coversThroughSeq) requests.delete(seq);
        setPainted(frame);
        const sheetId = (run.session.editState.sheets ?? run.session.state.sheets).find((sheet) => sheet.index === frame.request.sheet)?.id;
        pluginHost.presentGrid(sheetId ? { frame: frame.displayList, version: frame.version,
          sheetId, sheet: frame.request.sheet, zoom: frame.request.zoom, viewport: frame.request.viewport } : null);
        if (frame.coversThroughSeq !== undefined) run.session.markPresented({
          ...run.session.commandIdentity, coversThroughSeq: frame.coversThroughSeq,
        });
        run.firstPaint();
      },
      onCovered(coversThroughSeq) {
        const canvas = canvasRef.current;
        if (canvas) canvas.dataset.commandCoverage = String(coversThroughSeq);
        presentation.advanceCoverage(coversThroughSeq);
        run.session.markPresented({ ...run.session.commandIdentity, coversThroughSeq });
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
      const selected = sheetSelection(sheet);
      setSelection(selectionAt({ row: 0, col: 0 }));
      if (!sameSelection(selectionRef.current, selected)) void selectionNavigation.select(selected,
        { ...navigationOptions(), apply: (resolved) => { placeSelection(sheet, resolved); return true; } }, sheet);
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
    neighborFrames.clear();
    neighborRevision.current += 1;
  }, [proposals, neighborFrames]);
  useEffect(() => {
    const source = sourceRef.current;
    if (!run || !source) return;
    const unregister = source.registerPaintTransaction((painted, publish) => {
      publish();
      const canvas = canvasRef.current;
      if (canvas) {
        canvas.dataset.paintedSheet = String(painted.request.sheet);
        canvas.dataset.provisional = painted.provisional ? 'true' : 'false';
      }
      if (!painted.provisional && painted.request.sheet === activeRef.current) scheduleNeighborFrames.current();
    });
    if (source.painted && !source.painted.provisional) scheduleNeighborFrames.current();
    return () => {
      unregister();
      neighborRevision.current += 1;
      if (neighborTimer.current !== null) clearTimeout(neighborTimer.current);
      neighborTimer.current = null;
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
      if (frame && request && (source?.coversThroughSeq ?? -1) >= coordinator.submittedThroughSeq &&
        samePaintRequest(frame, { ...request, navigation: frame.request.navigation })) return;
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
  const mode = useRef<{ owner: PeerlessWorkbookSession | null; readOnly: boolean; allowHostProposals: boolean }>(
    { owner: null, readOnly: false, allowHostProposals: false });
  useLayoutEffect(() => {
    if (!run?.ready || mode.current.owner === run && mode.current.readOnly === Boolean(props.readOnly) &&
      mode.current.allowHostProposals === Boolean(props.allowHostProposals)) return;
    const previous = mode.current;
    mode.current = { owner: run, readOnly: Boolean(props.readOnly), allowHostProposals: Boolean(props.allowHostProposals) };
    if (previous.owner !== run && Boolean(run.session.editState.readOnly) === Boolean(props.readOnly) && !props.allowHostProposals) return;
    if (props.readOnly) {
      flushNudgeRef.current(true);
      draftController.finish(); setFormatToken(null); chartDrag.current = null; nudge.current = null;
      setChartOffset(null);
    }
    void coordinator.setReadOnly(Boolean(props.readOnly), props.allowHostProposals).result.catch(report);
  }, [run, run?.ready, props.readOnly, props.allowHostProposals, coordinator, draftController]);
  useEffect(() => {
    if (!run?.failure) return;
    echoRef.current?.dispose(); sourceRef.current?.dispose(); coordinator.fail(run.failure); draftController.dispose();
    chartDrag.current = null; nudge.current = null;
  }, [run, run?.failure, coordinator, draftController]);

  const command = (input: PeerlessWorkerCommandInput) => {
    const frame = sourceRef.current?.captureCommandFrame() ?? undefined;
    const selectionRevision = navigation.current;
    if (input.kind === 'select' && !draftController.ownsFocus) {
      void coordinator.command(input, { frame, selectionRevision }).result.catch(report);
    } else void coordinator.runAfterPendingInput(() => coordinator.command(input, { frame, selectionRevision }).result).catch(report);
  };
  const clear = () => {
    const target = selectedTarget();
    const owner = ownerRef.current;
    const selected = selectionRef.current;
    if (!editable() || !target || !owner || !selected) return undefined;
    const cleared: PeerlessWorkerCommandInput = { kind: 'applyOps', payload: { toolbar: { id: 'clear', args: null, target }, selection: target } };
    if (!editGate.gatesEdits()) { command(cleared); return undefined; }
    const sheet = activeRef.current;
    const range = normalizeRange(selected);
    const gated = coordinator.runAfterPendingInput(async () => {
      const content = await contentCells(owner, sheet, range);
      if (!content) throw new WorkerInputRefusal(t('editor.clearExceedsLimit'));
      const edits = await gateWrites(owner, target.sheetId, content, 'clear-key');
      return edits.length > 0 ? edits : null;
    }, 'posted');
    const ticket = coordinator.clipboard(() => gated, (edits): PeerlessWorkerCommandInput => edits
      ? { kind: 'editCells', payload: { args: [sheet, edits], target, selection: target } } : cleared,
    { frame: sourceRef.current?.captureCommandFrame() ?? undefined, selectionRevision: navigation.current });
    void ticket.result.catch(report);
    return ticket.result;
  };
  const clipboardAllowed = (kind: 'copy' | 'cut' | 'paste') => {
    const owner = ownerRef.current;
    return Boolean(owner?.ready && !owner.failure && selectionRef.current && (kind === 'copy' || editable()));
  };
  const clipboardAccess = (kind: 'copy' | 'cut' | 'paste'): { read?: Promise<{ text: string; html?: string | null }>;
    write?: Promise<void>; resolve?: (blob: Blob) => void; reject?: (error: unknown) => void; refused?: { error: unknown };
    marker?: string } => {
    if (kind === 'paste') {
      const read = readCellClipboard();
      void read.catch(() => {});
      return { read };
    }
    const marker = newClipboardMarker();
    let resolve!: (blob: Blob) => void;
    let reject!: (error: unknown) => void;
    const payload = new Promise<Blob>((done, fail) => { resolve = done; reject = fail; });
    void payload.catch(() => {});
    let write: Promise<void> | undefined;
    let refused: { error: unknown } | undefined;
    try {
      if (typeof ClipboardItem === 'function' && typeof navigator.clipboard?.write === 'function') {
        write = navigator.clipboard.write([new ClipboardItem({ 'text/plain': payload, 'text/html': clipboardHtml({ marker }, payload) })]);
        void write.catch(() => {});
      } else if (kind === 'cut' && typeof navigator.clipboard?.writeText !== 'function') {
        throw new WorkerInputRefusal('Cut requires clipboard write access');
      }
    } catch (error) { refused = { error }; }
    return { write, resolve, reject, refused, marker };
  };
  const clipboard = (kind: 'copy' | 'cut' | 'paste', access: ReturnType<typeof clipboardAccess>): Promise<unknown> => {
    if (!clipboardAllowed(kind)) {
      const refusal = new XlsxCommandAdmissionError('editor-unavailable');
      access.reject?.(refusal);
      return handledRejection(refusal);
    }
    const owner = ownerRef.current!;
    const selected = selectionRef.current!;
    const sheet = activeRef.current;
    const at = navigation.current;
    const target = rangeTarget(sheet, selected);
    const range = normalizeRange(selected);
    if (kind === 'paste') {
      let pasted: CellPaste | null = null;
      const ticket = coordinator.clipboard(() => {
        const reading = access.read!;
        const writing = cellClipboard.pending;
        return coordinator.runAfterPendingInput(async () => {
          await writing;
          const { text, html } = await reading;
          pasted = cellClipboard.paste(owner, text, range, html, sheet);
          const rows = pasted ? [] : parseCellClipboardText(text);
          const edits: CellInputEdit[] = pasted?.edits ?? [];
          let lastCol = 0;
          if (!pasted) rows.forEach((row, dr) => {
            lastCol = Math.max(lastCol, row.length - 1);
            row.forEach((input, dc) => edits.push({ row: range.top + dr, col: range.left + dc, input }));
          });
          const next: Selection = { anchor: { row: range.top, col: range.left }, focus: {
            row: pasted?.bottom ?? range.top + Math.max(0, rows.length - 1),
            col: pasted?.right ?? range.left + lastCol,
          } };
          return { edits: await gatePaste(owner, sheet, target.sheetId, edits), next };
        }, true);
      }, ({ edits, next }) => {
        if (ownerRef.current === owner && navigation.current === at && !draftController.ownsFocus &&
          !selectionNavigation.isPending(navigationContext())) predict(sheet, next, false, true);
        return { kind: 'editCells', payload: { args: [sheet, edits], target, selection: rangeTarget(sheet, next) } };
      });
      const settled = ticket.result.then((status) => { if (pasted) cellClipboard.consume(pasted, status.status === 'applied' && status.committed === true); });
      void settled.catch(report);
      return settled;
    }
    const checkOwner = () => {
      if (ownerRef.current !== owner || !owner.current || owner.retiring) throw new XlsxCommandAdmissionError('document-replaced');
      if (owner.failure) throw owner.failure;
    };
    const { write, resolve, reject, refused, marker } = access;
    if (refused) { reject!(refused.error); report(refused.error); return handledRejection(refused.error); }
    const token = cellClipboard.begin(undefined, marker);
    const copied = coordinator.runAfterPendingInput(async () => {
      checkOwner();
      const range = await copiedSelection(owner.session, sheet, selected);
      const result = await owner.session.call.cellInputs(sheet, rangeA1(range));
      checkOwner();
      const cleared = kind === 'cut' && editGate.gatesEdits()
        ? await gateCut(owner, target.sheetId, range, result.cells) : null;
      return { range, cells: result.cells, cleared };
    }, kind === 'cut' ? true : 'posted');
    const written = copied.then(async ({ range, cells }) => {
      checkOwner();
      const text = toTsv(cells);
      if (write) { resolve!(new Blob([text], { type: 'text/plain' })); await write; }
      else await navigator.clipboard.writeText(text);
      checkOwner();
      if (kind === 'copy') cellClipboard.capture(token, { owner, sheet, range, cells, cut: false }, Boolean(write));
    });
    cellClipboard.pending = written.catch(() => {});
    if (kind === 'cut') {
      const ticket = coordinator.clipboard(() => written.then(() => copied), ({ cleared }): PeerlessWorkerCommandInput =>
        cleared?.length ? { kind: 'editCells', payload: { args: [sheet, cleared], target, selection: target } }
          : { kind: 'applyOps', payload: { toolbar: { id: 'clear', args: null, target }, selection: target } });
      void ticket.result.catch((error) => { reject!(error); report(error); });
      return ticket.result;
    }
    void written.catch((error) => { reject!(error); report(error); });
    return written;
  };
  const menuClipboard = (kind: 'copy' | 'cut' | 'paste' | 'pastePlainText'): Promise<unknown> => {
    const base = kind === 'pastePlainText' ? 'paste' : kind;
    if (!clipboardAllowed(base) || draftController.getSnapshot().cell) {
      return handledRejection(new XlsxCommandAdmissionError('editor-unavailable'));
    }
    let access: ReturnType<typeof clipboardAccess>;
    if (kind === 'pastePlainText') {
      const read = navigator.clipboard.readText().then((text) => ({ text, html: null }));
      void read.catch(() => {});
      access = { read };
    } else access = clipboardAccess(base);
    return clipboard(base, access);
  };
  const settledFor = (owner: NonNullable<typeof run>) => ownerRef.current === owner && owner.ready && owner.current &&
    !owner.retiring && !owner.failure && !coordinator.pending && !draftController.ownsFocus &&
    owner.session.throughSeq >= coordinator.submittedThroughSeq;
  const proposalState = () => `${proposalsRevision.current}:${proposalsMutation.current}`;
  const prefetchNeighborFrames = async (owner: NonNullable<typeof run>, revision: number) => {
    const version = owner.session.confirmedBasis.version;
    const proposalsAt = proposalState();
    neighborVersion.current = version;
    const live = () => revision === neighborRevision.current && settledFor(owner) &&
      owner.session.confirmedBasis.version === version && proposalState() === proposalsAt;
    const sheets = catalogRef.current;
    const at = sheets.findIndex((entry) => entry.index === activeRef.current);
    const neighbors = at < 0 ? [] : [sheets[at + 1], sheets[at - 1]].filter(
      (entry): entry is WorkbookSheetSummary => entry !== undefined);
    for (const index of [...neighborFrames.keys()]) {
      if (!neighbors.some((entry) => entry.index === index)) neighborFrames.delete(index);
    }
    for (const neighbor of neighbors) {
      const current = capture();
      if (!current || !live()) return;
      if (costlyNeighbors.has(neighbor.id)) continue;
      const cached = neighborFrames.get(neighbor.index);
      const position = sheetScrolls.get(neighbor.id);
      if (cached && cached.sheetId === neighbor.id && cached.paint.version === version && cached.proposals === proposalsAt &&
        (!position || position.x === cached.request.viewport.x && position.y === cached.request.viewport.y) &&
        cached.request.generation === current.generation && cached.request.zoom === current.zoom &&
        cached.request.dpr === current.dpr && cached.request.width === current.width &&
        cached.request.height === current.height) continue;
      neighborFrames.delete(neighbor.index);
      const view = await owner.session.call.sheetView(neighbor.index);
      if (!live() || view.sheet !== neighbor.index) return;
      const started = performance.now();
      const frame = await owner.session.call.frame({ x: position?.x ?? view.initialScrollX,
        y: position?.y ?? view.initialScrollY, width: current.viewport.width, height: current.viewport.height },
      { sheet: neighbor.index });
      if (performance.now() - started > NEIGHBOR_FRAME_BUDGET_MS) {
        costlyNeighbors.add(neighbor.id);
        continue;
      }
      if (!live() || frame.version !== version || frame.sheet !== neighbor.index) return;
      neighborFrames.set(neighbor.index, { sheetId: neighbor.id, proposals: proposalsAt,
        request: { ...current, sheet: neighbor.index, viewport: { ...frame.viewport } },
        paint: { displayList: frame.displayList, mergedRanges: frame.mergedRanges ?? [], version: frame.version } });
      paintWithDefaultFonts.preload(frame.displayList, live);
    }
  };
  scheduleNeighborFrames.current = () => {
    const owner = ownerRef.current;
    if (neighborTimer.current !== null) clearTimeout(neighborTimer.current);
    neighborTimer.current = null;
    const revision = ++neighborRevision.current;
    if (!owner) return;
    neighborTimer.current = setTimeout(() => {
      neighborTimer.current = null;
      if (revision === neighborRevision.current && settledFor(owner)) {
        void prefetchNeighborFrames(owner, revision).catch(() => {});
      }
    }, neighborVersion.current === undefined || neighborVersion.current === owner.session.confirmedBasis.version
      ? NEIGHBOR_FRAME_DELAY_MS : NEIGHBOR_FRAME_EDIT_DELAY_MS);
  };
  const paintNeighborFrame = (sheet: number) => {
    const owner = ownerRef.current;
    const source = sourceRef.current;
    const entry = neighborFrames.get(sheet);
    if (!owner || !source || !entry || entry.sheetId !== sheetId(sheet) || !settledFor(owner) ||
      entry.paint.version !== owner.session.confirmedBasis.version || entry.proposals !== proposalState()) return;
    const request = capture();
    if (request && sameFrameRequest(entry.request, request)) source.installProvisional(entry.paint, request);
  };
  const navigate = (sheet: number, next?: Selection) => {
    if (!run?.ready || run.failure || draftController.isComposing('cell') || draftController.isComposing('formula')) return;
    draftController.commit('cell'); draftController.commit('formula');
    const selected = next ?? sheetSelection(sheet);
    selectionNavigation.select(selected, { ...navigationOptions(), apply: (resolved) => {
      const target = rangeTarget(sheet, resolved);
      const switching = sheet !== activeRef.current;
      placeSelection(sheet, resolved);
      if (next) reveal(sheet, resolved.anchor);
      else if (switching) paintNeighborFrame(sheet);
      command({ kind: 'setActiveSheet', payload: { args: [sheet], target, selection: target } });
      return true;
    } }, sheet, true);
  };
  nameNavigation.current = (cell) => { navigate(activeRef.current, selectionAt(cell)); focus(); };
  sheetNavigation.current = (sheet) => { if (sheet === activeRef.current) navigate(sheet); };
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
    readOnly: () => propsRef.current.readOnly ?? false,
    allowHostProposals: () => propsRef.current.allowHostProposals ?? false, collaborative: () => propsRef.current.collaboration !== undefined,
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
      if (propsRef.current.downloadOnSave ?? !propsRef.current.onSave) {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: XLSX_MIME }));
        const anchor = document.createElement('a');
        anchor.href = url; anchor.download = propsRef.current.fileName ?? 'workbook.xlsx';
        document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
      }
    },
    exportFonts: (fonts) => editorExportFaces(fonts, run?.session.call.decodeFontFace, propsRef.current.defaultFonts),
    deliverPng(bytes) {
      const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' }));
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `${(propsRef.current.fileName ?? 'workbook.xlsx').replace(/\.xlsx$/i, '')}.png`;
      document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
    },
    focusGrid: focus, refuse: (reason) => setRefusal({ code: reason.code as XlsxCommandFailureCode }),
    menu: {
      clipboard: menuClipboard,
      clear: () => clear() ?? Promise.resolve(),
      beginRename: () => setRenaming(activeRef.current),
      activate: (name) => { pendingActivation.current = name.toLowerCase(); bumpActivation(); },
    },
  });
  useCommandShortcuts(commands, rootRef);
  useEffect(() => {
    const name = pendingActivation.current;
    const entry = name === null ? undefined : catalog.find((sheet) => sheet.name.toLowerCase() === name);
    if (!entry) return;
    pendingActivation.current = null;
    if (entry.index !== activeRef.current || entry.id !== activeSheetId.current) navigate(entry.index);
  });
  const selectAtPoint = (request: { x: number; y: number }) => {
    const { cell, chart } = point({ clientX: request.x, clientY: request.y });
    if (chart) { setSelectedChart({ id: chart.id, movable: chart.movable }); return null; }
    if (!cell) return null;
    const sheet = activeRef.current;
    const covers = () => {
      const range = selectionRef.current ? normalizeRange(selectionRef.current) : null;
      return !chartRef.current && range !== null && activeRef.current === sheet &&
        cell.row >= range.top && cell.row <= range.bottom && cell.col >= range.left && cell.col <= range.right;
    };
    if (!covers()) {
      setSelectedChart(null);
      const next = selectionAt(mergeAnchorCell(cell, merges()));
      predict(sheet, next, true);
      command({ kind: 'select', payload: { selection: rangeTarget(sheet, next) } });
    }
    return covers;
  };
  const menuEntries = (entries: OverflowMenuEntry[], ready: () => boolean) => ready() ? entries
    : entries.map((entry) => entry.kind === 'item' || entry.kind === 'submenu' ? { ...entry, disabled: true } : entry);
  const gridMenu = useContextMenu({
    surface: scrollRef,
    subscribe: commands.store.subscribe,
    keyboardPoint: () => {
      const selected = selectionRef.current;
      const frame = sourceRef.current?.painted;
      const canvas = canvasRef.current;
      const rect = selected && frame?.geometry ? cellRect(frame.geometry, selected.focus.row, selected.focus.col) : null;
      if (!rect || !canvas || !frame) return null;
      const box = canvas.getBoundingClientRect();
      return { x: box.left + (rect.x + rect.w / 2) * frame.request.zoom, y: box.top + (rect.y + rect.h) * frame.request.zoom };
    },
    resolve: (request) => {
      if (!run?.ready || run.failure || inPluginChrome(request.target)) return null;
      if (draftController.cellInput?.contains(request.target as Node)) return null;
      let ready = () => true;
      if (request.modality === 'pointer') {
        const covers = selectAtPoint(request);
        if (!covers) return null;
        ready = covers;
      } else if (chartRef.current) return null;
      return { label: t('editor.cellMenuLabel'), testId: 'xlsx-cell-menu',
        entries: () => menuEntries(xlsxContextMenuEntries('cell', commands.store, t), ready) };
    },
  });
  const tabMenu = useContextMenu({
    surface: tabsRef,
    subscribe: commands.store.subscribe,
    keyboardPoint: () => {
      const tab = tabsRef.current?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]');
      const box = tab?.getBoundingClientRect();
      return box ? { x: box.left, y: box.top } : null;
    },
    resolve: (request) => {
      if (!run?.ready || run.failure || renaming !== null) return null;
      const tab = request.target instanceof Element ? request.target.closest<HTMLElement>('[role="tab"]') : null;
      const index = Number(tab?.dataset.sheetIndex);
      if (!tab || !Number.isInteger(index)) return null;
      const id = sheetId(index);
      if (index !== activeRef.current || id !== activeSheetId.current) navigate(index);
      const ready = () => activeRef.current === index && activeSheetId.current === id;
      return { label: t('editor.sheetTabMenuLabel'), testId: 'xlsx-sheet-tab-menu',
        entries: () => menuEntries(xlsxContextMenuEntries('sheetTab', commands.store, t), ready) };
    },
  });
  const renameDone = useRef(false);
  const finishRename = (input: HTMLInputElement | null, commit: boolean) => {
    if (renameDone.current) return;
    renameDone.current = true;
    const name = input?.value ?? '';
    setRenaming(null);
    requestAnimationFrame(() => tabsRef.current?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.focus());
    if (commit) void commands.store.execute('renameSheet', { name });
  };
  useLayoutEffect(() => { if (renaming !== null) renameDone.current = false; }, [renaming]);
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

  const point = (event: Pick<MouseEvent, 'clientX' | 'clientY'>) => {
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
    if (!run?.ready || !editable() || !selected || !typingAllowed(selected.anchor)) return;
    reveal(activeRef.current, selected.anchor);
    draftController.open(seed);
  };
  const onMouseDown = (event: MouseEvent) => {
    if (fromPlugin(event) || !run?.ready || run.failure || draftController.cellInput?.contains(event.target as Node) ||
      draftController.isComposing('cell') || draftController.isComposing('formula')) return;
    draftController.commit('cell'); draftController.commit('formula'); flushNudge();
    if (event.button === 2 || (event.button === 0 && event.ctrlKey && isMacPlatform())) return;
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
    const next = event.shiftKey ? extendTo(selected!, cell, limits()) : selectionAt(mergeAnchorCell(cell, merges()));
    predict(activeRef.current, next, true);
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
      predict(activeRef.current, next, true);
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
    const afterNavigation = (task: () => void, cancel: () => void) => {
      if (!selectionNavigation.beforeInput({ ...navigationOptions(), apply: (next) => {
        placeSelection(activeRef.current, next); return true;
      } }, task, cancel)) task();
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
    const movement = selectionKeyReducer(selectionRef.current, event, limits(), undefined, hidden(), merges()).type === 'move';
    if (request || movement) {
      const key = { key: event.key, ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, altKey: event.altKey };
      const act = () => {
        if (draftController.getSnapshot().cell) {
          if (draftController.cellKey(key)) return;
          if (key.key === 'Enter') draftController.commit('cell', key.shiftKey ? 'up' : 'down');
          else if (key.key === 'Tab') draftController.commit('cell', key.shiftKey ? 'left' : 'right');
          return;
        }
        selectionNavigation.enqueue(key, navigationOptions());
      };
      if (selectionNavigation.defer(navigationContext(), act, true)) { event.preventDefault(); return; }
      event.preventDefault();
      selectionNavigation.enqueue(event, navigationOptions());
      return;
    }
    const key = { key: event.key, shiftKey: event.shiftKey, metaKey: event.metaKey, ctrlKey: event.ctrlKey,
      altKey: event.altKey };
    const extending = key.shiftKey && key.key !== 'Tab' && key.key !== 'Enter';
    const from = (selected: Selection) => extending ? selected : selectionAt(selected.anchor);
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
      const action = selectionKeyReducer(from(selected), key, limits(), undefined, hidden(), merges());
      if (action.type === 'move') {
        predict(activeRef.current, action.selection);
        reveal(activeRef.current, action.selection.focus);
        command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, action.selection) } });
      } else if (action.type === 'startEdit') openCell(action.initialInput);
      else if (action.type === 'clear') clear();
    };
    const type = selectionKeyReducer(from(selectionRef.current), key, limits(), undefined, hidden(), merges()).type;
    if (type === 'none') {
      if (key.key === 'Escape' && selectionNavigation.defer(navigationContext(), act, true)) event.preventDefault();
      return;
    }
    event.preventDefault();
    if (!selectionNavigation.defer(navigationContext(), act, false)) act();
  };

  const sheets = catalog;
  const defaultToolbar = useMemo(() => <PeerlessDefaultToolbar />, []);
  const reviewToolbar = props.allowHostProposals && isProposalsAvailable() ? <ToolbarCommandButton id="proposalsPanel" /> : null;
  const toolbar = props.showToolbar === false ? null : props.toolbar === undefined ? props.readOnly ? reviewToolbar : defaultToolbar : props.toolbar;
  const renderDock = (placement: DockPlacement) => pluginHost.managed ? <PluginDock host={pluginHost.host}
    placement={placement} activations={pluginHost.activations.filter((activation) => activation.plugin.panel?.placement === placement)}
    available={dockArea} /> : null;
  const painted = sourceRef.current?.painted;
  return <XlsxCommandContext.Provider value={commands.store}><EditorChromeContext.Provider value={true}>
    <PeerlessChromeContext.Provider value={true}>
    <PeerlessWorkerInputs onNavigate={navigateFromNameBox} controller={draftController} presentation={presentation} run={run} active={active}
      selectedChart={selectedChart !== null} error={Boolean(error)} readOnly={Boolean(props.readOnly)} navigation={navigation}
      cellEditable={typingAllowed} gateRevision={pluginHost.activations}>
      <div dir="ltr" ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
        onKeyDownCapture={(event) => {
          if (!fromPlugin(event) && event.key === 'Escape' && !(event.target instanceof Element && event.target.closest('[role="menu"]'))) {
            cellClipboard.cancelCut(); setFormatToken(null); commands.refresh();
          }
        }}
        style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%', height: '100%',
          minWidth: 0, color: '#202124', background: '#ffffff', fontFamily: 'ui-sans-serif, system-ui, sans-serif', ...props.style }}>
        {props.file && !run?.ready && !error && props.loadingIndicator}
        <PeerlessSelectionEffect store={presentation} onChange={applySelectionFormat} />
        {toolbar != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbar}</div>}
        <div ref={dockAreaRef} data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minHeight: 0 }}>
          {renderDock('left')}
          <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
          {isProposalsAvailable() && proposalsPanelOpen && <ProposalsPanel proposals={proposals} staleFor={staleFor}
            style={{ top: 4, right: 8, width: 'min(320px, calc(100% - 16px))', maxHeight: 'min(420px, calc(100% - 12px))' }} />}
          {pluginHost.rowGutter && <PluginRowGutter host={pluginHost.host} index={pluginHost.rowGutter}
            painted={pluginHost.painted} zoom={zoom} />}
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
            style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none',
              ...(pluginHost.rowGutter ? { marginLeft: ROW_GUTTER_WIDTH * zoom } : {}) }}>
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
          {error && sheets.length === 0 && <WorkbookFailure kind="open" details={error.message} />}
          {!props.file && !error && <div data-testid="xlsx-empty" role="status" style={{ position: 'absolute', inset: 0,
            display: 'grid', placeItems: 'center', color: '#5f6368', pointerEvents: 'none' }}>{t('editor.openPrompt')}</div>}
        </div>
        <PeerlessWorkerAccessibility store={presentation} active={active} sheets={sheets} />
        {refusal && <div data-testid="xlsx-input-refusal" role="alert">
          {refusal.message ?? <>{commandReason(refusal.code, { translate: t }).message}{refusal.detail ? `: ${refusal.detail}` : ''}</>}
        </div>}
        {error && sheets.length > 0 && <WorkbookFailure kind="stopped" details={error.message} />}
        {gridMenu.element}{tabMenu.element}
        {sheets.length > 0 && <div ref={tabsRef} data-testid="xlsx-sheet-tabs" role="tablist" aria-label={t('editor.sheetTabsLabel')}
          style={{ display: 'flex', gap: 2, padding: '4px 6px', borderTop: '1px solid #e0e0e0', background: '#fafafa', overflowX: 'auto' }}>
          {sheets.map((sheet) => renaming === sheet.index ? <input key={sheet.id} data-testid="xlsx-sheet-rename"
            aria-label={t('editor.renameSheetLabel')} defaultValue={sheet.name} maxLength={31} autoFocus
            onFocus={(event) => event.currentTarget.select()}
            onKeyDown={(event) => {
              if (event.key === 'F6') event.stopPropagation();
              if (event.key === 'Enter') { event.preventDefault(); finishRename(event.currentTarget, true); }
              else if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); finishRename(event.currentTarget, false); }
            }}
            onBlur={(event) => finishRename(event.currentTarget, true)}
            style={{ border: `1px solid ${BRAND}`, padding: '3px 8px', font: 'inherit', width: Math.max(80, sheet.name.length * 9) }} />
            : <button key={sheet.id} role="tab" aria-selected={sheet.index === active} disabled={Boolean(error)}
            data-sheet-index={sheet.index}
            onClick={() => navigate(sheet.index)}
            onDoubleClick={() => { if (commands.store.getState('renameSheet').enabled) void commands.store.execute('renameSheet', null); }}
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
