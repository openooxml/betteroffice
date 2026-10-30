import type { WorkbookSheetSummary } from '@betteroffice/xlsx';
import type { CollaborationAwareness } from '@betteroffice/xlsx/collaboration';
import { PresenceStrip, RemoteSelections } from '../presence/Presence';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import { usePeerlessPresence } from './usePeerlessPresence';
import { WorkerPresentationStore, useWorkerPresentation } from './WorkerPresentationStore';

const selectSelection = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.selection;
const selectPainted = (state: ReturnType<WorkerPresentationStore['getSnapshot']>) => state.painted;

export function PeerlessPresence({ store, run, provider, active, sheets }: {
  store: WorkerPresentationStore; run: PeerlessWorkbookSession | null;
  provider: CollaborationAwareness; active: number; sheets: WorkbookSheetSummary[];
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
