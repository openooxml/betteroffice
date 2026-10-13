import type { WorkbookSession } from '@betteroffice/xlsx';
import type { PeerlessWorkerInputCoordinator } from '../commands/peerlessWorkerInputCoordinator';

export function readWorkerCellInput(session: WorkbookSession, sheet: number, a1: string): Promise<string> {
  return session.call.cellInputs(sheet, a1).then((result) => result.cells[0]?.[0]?.input ?? '');
}

export function holdWorkerCellInputRead(coordinator: PeerlessWorkerInputCoordinator, read: () => Promise<string>,
  receive: (value: string) => void, applied: Promise<boolean>, cancelled: () => boolean,
  refused: (error: unknown) => void): Promise<void> {
  let failure: { error: unknown } | undefined;
  return coordinator.runAfterPendingInput(async () => {
    if (cancelled()) return;
    try {
      const loaded = read().then((value) => { if (!cancelled()) receive(value); });
      await Promise.race([loaded.then(() => applied), applied]);
    } catch (error) { failure = { error }; throw error; }
  }, 'settled').catch((error) => {
    if (failure) throw failure.error;
    refused(error);
  });
}
