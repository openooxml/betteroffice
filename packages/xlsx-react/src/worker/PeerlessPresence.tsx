import type { Selection, WorkbookSheetSummary } from '@betteroffice/xlsx';
import type { TFunction } from '@betteroffice/xlsx-i18n';
import type { CollaborationAwareness } from '@betteroffice/xlsx/collaboration';
import { PresenceStrip, RemoteSelections, type PresenceGeometry } from '../presence/PresenceView';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import { usePeerlessPresence } from './usePeerlessPresence';
import type { WorkerPaintResult } from './PeerlessWorkerPaintSource';

export function PeerlessPresence({ painted, selection, run, provider, active, sheets, t, geometry }: {
  painted: WorkerPaintResult | null; selection: Selection | null; run: PeerlessWorkbookSession | null;
  provider: CollaborationAwareness; active: number; sheets: WorkbookSheetSummary[]; t: TFunction;
  geometry: PresenceGeometry & { normalizeRange: typeof import('@betteroffice/xlsx')['normalizeRange'] };
}) {
  const peers = usePeerlessPresence(run, provider, active, selection, sheets, geometry.normalizeRange);
  const sheetIds = sheets.map((sheet) => sheet.id);
  return <>
    <div style={{ position: 'absolute', top: 8, left: 0, width: painted?.request.width ?? 0,
      display: 'flex', justifyContent: 'flex-end', pointerEvents: 'none' }}>
      <PresenceStrip t={t} peers={peers} sheetIds={sheetIds} sheetNames={sheets.map((sheet) => sheet.name)} activeSheet={active} />
    </div>
    {painted?.request.sheet === active && <div style={{ pointerEvents: 'none' }}>
      <RemoteSelections geometry={geometry} peers={peers} sheetIds={sheetIds} activeSheet={active}
        grid={painted.displayList.grid} zoom={painted.request.zoom} mergedRanges={painted.mergedRanges ?? []} />
    </div>}
  </>;
}
