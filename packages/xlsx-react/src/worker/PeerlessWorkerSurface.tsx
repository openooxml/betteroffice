import { memo, useCallback, useEffect, useMemo } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { buildA11yGrid, cellRect, normalizeRange, rangeRect } from '@betteroffice/xlsx';
import type { WorkbookSheetSummary } from '@betteroffice/xlsx';
import type { CollaborationAwareness } from '@betteroffice/xlsx/collaboration';
import { useTranslation } from '../i18n';
import { expandRangeToMergedCells, scaledRect } from '../viewer/sessionGeometry';
import { PeerlessWorkerCellInput, usePeerlessReadOnly } from './PeerlessWorkerInputs';
import { cellEditorColors } from '../viewer/cellEditorColors';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import type { WorkerDraftController } from './WorkerDraftController';
import { WorkerPresentationStore, useWorkerPresentation } from './WorkerPresentationStore';
import { PresenceStrip, RemoteSelections } from '../presence/Presence';
import { ACTIVE_CELL_WIDTH, marqueeBox, marqueeFills, MARQUEE_COLOR, MARQUEE_WIDTH, SELECTION_FILL } from '../selectionOutline';
import { usePeerlessPresence } from './usePeerlessPresence';
import { A11yGridTable } from '../components/A11yGridCell';
import { SheetGrid } from '../components/SheetGrid';
import type { SheetGridHeaderHandlers } from '../components/SheetGrid';

const selectSelection = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.selection;
const selectPainted = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.painted;
const selectView = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.view;
const hidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, border: 0,
  overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap',
};

/**
 * The peerless paint lives in the presentation store, not in this shell's
 * state. Subscribe here so the header bands advance with the same painted
 * frame the canvas is showing instead of freezing at the first paint.
 */
export const PeerlessSheetGrid = memo(function PeerlessSheetGrid({ store, active, leadingInset = 0, headerHandlers, selectAllLabel, unhideLabels, children }: {
  store: WorkerPresentationStore; active: number; leadingInset?: number;
  headerHandlers?: SheetGridHeaderHandlers;
  selectAllLabel?: string;
  unhideLabels?: { col: string; row: string }; children: ReactNode;
}) {
  const painted = useWorkerPresentation(store, selectPainted);
  // the headers follow the same gated sheet as the paint: a stale selection from
  // the sheet just left must not light tracks on the one now shown.
  const selection = useWorkerPresentation(store, selectSelection);
  const paintedSheet = painted?.request.sheet === active ? active : null;
  return <SheetGrid frame={painted?.request.sheet === active ? painted.displayList : null}
    zoom={painted?.request.zoom ?? 1} leadingInset={leadingInset} headerHandlers={headerHandlers}
    sheet={paintedSheet}
    selectAllLabel={selectAllLabel}
    unhideLabels={unhideLabels}
    selection={painted?.request.sheet === active ? selection : null}>{children}</SheetGrid>;
});
export const PeerlessWorkerExtent = memo(function PeerlessWorkerExtent({ store, zoom }: {
  store: WorkerPresentationStore; zoom: number;
}) {
  const view = useWorkerPresentation(store, selectView);
  return <div style={{ position: 'absolute', top: 0, left: 0,
    width: view ? view.contentWidth * zoom : '100%', height: view ? view.contentHeight * zoom : '100%' }} />;
});

export const PeerlessWorkerOverlay = memo(function PeerlessWorkerOverlay({ store, active, zoom, selectedChart,
  chartOffset, controller, disabled, run }: {
  store: WorkerPresentationStore;
  active: number;
  zoom: number;
  selectedChart: { id: string; movable: boolean } | null;
  chartOffset: { x: number; y: number } | null;
  controller: WorkerDraftController;
  disabled: boolean;
  run: PeerlessWorkbookSession | null;
}) {
  const unavailable = usePeerlessReadOnly(run, disabled);
  const painted = useWorkerPresentation(store, selectPainted);
  const selection = useWorkerPresentation(store, selectSelection);
  const scale = painted?.request.zoom ?? zoom;
  const visible = painted?.request.sheet === active ? selection : null;
  const grid = painted?.displayList.grid;
  const expanded = visible ? expandRangeToMergedCells(normalizeRange(visible), painted?.mergedRanges ?? []) : null;
  const selectionRect = grid && expanded ? rangeRect(grid, expanded) : null;
  const focused = visible ? expandRangeToMergedCells({ top: visible.anchor.row, left: visible.anchor.col,
    bottom: visible.anchor.row, right: visible.anchor.col }, painted?.mergedRanges ?? []) : null;
  const focusRect = grid && focused ? rangeRect(grid, focused) : null;
  const chartRect = selectedChart ? painted?.displayList.charts?.find((chart) => chart.id === selectedChart.id)?.rect : null;
  const rect = useCallback((row: number, col: number, sheet: number) =>
    painted?.geometry && painted.request.sheet === sheet ? cellRect(painted.geometry, row, col) : null, [painted]);
  const colors = useCallback((sheet: number, bounds: { x: number; y: number; w: number; h: number }) =>
    cellEditorColors(painted?.request.sheet === sheet ? painted.displayList : null, bounds), [painted]);
  // The chart box keeps the outline: it is a movable frame rather than a cell
  // edge and has no snapped interior to straddle.
  const outline = (bounds: NonNullable<typeof selectionRect>, color: string, width: 1 | 2 = 2): CSSProperties => {
    const scaled = scaledRect(bounds, scale);
    return { position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      outline: `${width}px solid ${color}`, outlineOffset: -width / 2 };
  };
  // Cell selections ride the same snapped pixel grid as the canvas gridlines:
  // the range ring is one pixel on the snapped rect and the active-cell ring is
  // two pixels straddling its boundaries, never an outline that rasterizes
  // asymmetrically at a fractional zoom.
  const marquee = (bounds: NonNullable<typeof selectionRect>, width = MARQUEE_WIDTH): CSSProperties => {
    const box = marqueeBox(scaledRect(bounds, scale), width);
    return { position: 'absolute', left: box.left, top: box.top, width: box.width, height: box.height,
      boxShadow: `inset 0 0 0 ${width}px ${MARQUEE_COLOR}` };
  };
  const fills = selectionRect && !selectedChart
    ? marqueeFills(scaledRect(selectionRect, scale), focusRect ? scaledRect(focusRect, scale) : null) : [];
  return <div data-testid="xlsx-overlay-host"
    style={{ position: 'absolute', top: 0, left: 0, width: 0, height: 0, pointerEvents: 'none' }}>
    {fills.map((fill, index) => <div key={index} data-testid="xlsx-selection-fill"
      style={{ position: 'absolute', left: fill.x, top: fill.y, width: fill.w, height: fill.h,
        background: SELECTION_FILL }} />)}
    {selectionRect && !selectedChart && <div data-testid="xlsx-selection"
      style={marquee(selectionRect)} />}
    {focusRect && !selectedChart && <div data-testid="xlsx-active-cell"
      style={marquee(focusRect, ACTIVE_CELL_WIDTH)} />}
    {chartRect && <div data-testid="xlsx-chart-selection" data-chart-id={selectedChart?.id} aria-hidden
      style={{ ...outline(chartRect, '#217346'), transform: `translate(${chartOffset?.x ?? 0}px, ${chartOffset?.y ?? 0}px)` }} />}
    <PeerlessWorkerCellInput controller={controller} zoom={scale} disabled={unavailable} rect={rect} colors={colors} />
  </div>;
});

export const PeerlessWorkerAccessibility = memo(function PeerlessWorkerAccessibility({ store, active, sheets, onFollowLink }: {
  store: WorkerPresentationStore; active: number; sheets: WorkbookSheetSummary[];
  onFollowLink?: (row: number, col: number) => void;
}) {
  const { t } = useTranslation();
  const painted = useWorkerPresentation(store, selectPainted);
  const selection = useWorkerPresentation(store, selectSelection);
  const frame = painted?.displayList;
  const visible = painted?.request.sheet === active ? selection : null;
  const sheetName = painted ? sheets[painted.request.sheet]?.name ?? '' : '';
  const grid = useMemo(() => frame ? buildA11yGrid(frame, visible, sheetName, {
    gridLabel: t('a11y.gridLabel'), rowHeaderLabel: t('a11y.rowHeaderLabel'), columnHeaderLabel: t('a11y.columnHeaderLabel'),
    cellLabel: t('a11y.cellLabel'), cellLabelSelected: t('a11y.cellLabelSelected'), emptyCellLabel: t('a11y.emptyCellLabel'),
    emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
  }) : null, [frame, visible, sheetName, t]);
  if (!grid) return null;
  return <PeerlessAccessibilityGrid grid={grid} onFollowLink={onFollowLink} />;
});

const PeerlessAccessibilityGrid = memo(function PeerlessAccessibilityGrid({ grid, onFollowLink }: {
  grid: ReturnType<typeof buildA11yGrid>; onFollowLink?: (row: number, col: number) => void;
}) {
  return <>
    <A11yGridTable grid={grid} style={hidden} onFollowLink={onFollowLink} />
    {grid.charts.map((chart, index) => <div key={`${index}:${chart.label}`} style={hidden} role="img" aria-label={chart.label} />)}
  </>;
});

export function PeerlessSelectionEffect({ store, onChange }: { store: WorkerPresentationStore; onChange(): void }) {
  const selection = useWorkerPresentation(store, selectSelection);
  useEffect(onChange, [selection, onChange]);
  return null;
}

export function PeerlessWorkerPresence({ store, run, provider, active, sheets }: {
  store: WorkerPresentationStore; run: PeerlessWorkbookSession | null;
  provider: CollaborationAwareness | null | undefined; active: number; sheets: WorkbookSheetSummary[];
}) {
  const painted = useWorkerPresentation(store, selectPainted);
  const selection = useWorkerPresentation(store, selectSelection);
  const peers = usePeerlessPresence(run, provider, active, selection, sheets);
  const sheetIds = sheets.map((sheet) => sheet.id);
  return <>
    <div style={{ position: 'absolute', top: 8, left: 0, width: painted?.request.width ?? 0,
      display: 'flex', justifyContent: 'flex-end', pointerEvents: 'none' }}>
      <PresenceStrip peers={peers} sheetIds={sheetIds} sheetNames={sheets.map((sheet) => sheet.name)} activeSheet={active} />
    </div>
    {painted?.request.sheet === active && <div style={{ pointerEvents: 'none' }}>
      <RemoteSelections peers={peers} sheetIds={sheetIds} activeSheet={active}
        grid={painted.displayList.grid} zoom={painted.request.zoom} mergedRanges={painted.mergedRanges ?? []} />
    </div>}
  </>;
}
