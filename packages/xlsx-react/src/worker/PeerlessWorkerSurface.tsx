import { memo, useCallback, useEffect, useMemo } from 'react';
import type { CSSProperties } from 'react';
import { buildA11yGrid, cellRect, normalizeRange, rangeRect } from '@betteroffice/xlsx';
import type { WorkbookSheetSummary } from '@betteroffice/xlsx';
import type { CollaborationAwareness } from '@betteroffice/xlsx/collaboration';
import { useTranslation } from '../i18n';
import { expandRangeToMergedCells, scaledRect } from '../viewer/sessionGeometry';
import { PeerlessWorkerCellInput, usePeerlessReadOnly } from './PeerlessWorkerInputs';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import type { WorkerDraftController } from './WorkerDraftController';
import { WorkerPresentationStore, useWorkerPresentation } from './WorkerPresentationStore';
import { PresenceStrip, RemoteSelections } from '../presence/Presence';
import { usePeerlessPresence } from './usePeerlessPresence';

const selectSelection = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.selection;
const selectPainted = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.painted;
const selectView = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.view;
const hidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, border: 0,
  overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap',
};

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
  const outline = (bounds: NonNullable<typeof selectionRect>, border: string, background?: string): CSSProperties => {
    const scaled = scaledRect(bounds, scale);
    return { position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      boxSizing: 'border-box', border, background };
  };
  return <div data-testid="xlsx-overlay-host"
    style={{ position: 'absolute', top: 0, left: 0, width: 0, height: 0, pointerEvents: 'none' }}>
    {selectionRect && !selectedChart && <div data-testid="xlsx-selection"
      style={outline(selectionRect, '1px solid #217346', 'rgba(33, 115, 70, 0.12)')} />}
    {focusRect && !selectedChart && <div style={outline(focusRect, '2px solid #217346')} />}
    {chartRect && <div data-testid="xlsx-chart-selection" data-chart-id={selectedChart?.id} aria-hidden
      style={{ ...outline(chartRect, '2px solid #217346'), transform: `translate(${chartOffset?.x ?? 0}px, ${chartOffset?.y ?? 0}px)` }} />}
    <PeerlessWorkerCellInput controller={controller} zoom={scale} disabled={unavailable} rect={rect} />
  </div>;
});

export const PeerlessWorkerAccessibility = memo(function PeerlessWorkerAccessibility({ store, active, sheets }: {
  store: WorkerPresentationStore; active: number; sheets: WorkbookSheetSummary[];
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
  return <PeerlessAccessibilityGrid grid={grid} />;
});

const PeerlessAccessibilityGrid = memo(function PeerlessAccessibilityGrid({ grid }: { grid: ReturnType<typeof buildA11yGrid> }) {
  return <>
    <div style={hidden} role="grid" aria-label={grid.label}>
      <div role="row"><span role="columnheader" />{grid.columnHeaders.map((header) =>
        <span key={header.col} role="columnheader">{header.label}</span>)}</div>
      {grid.rows.map((row) => <div key={row.row} role="row"><span role="rowheader">{row.header}</span>
        {row.cells.map((cell) => <span key={cell.col} role="gridcell" aria-selected={cell.selected}>{cell.label}</span>)}
      </div>)}
    </div>
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
