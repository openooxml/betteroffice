import { expect, spyOn, test } from 'bun:test';
import type { WorkbookSession } from '@betteroffice/xlsx';
import { createWorkbookSessionRecovery, workbookSessionRecovery } from '../../../xlsx/src/session/clientRecovery';
import { testSessionCommands } from '../commands/testing';
import { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

test.each([20, 120, 40])('idle checkpoint waits for twice a %s ms encode or 150 ms', async (encodeMs) => {
  const unexpected = async () => { throw new Error('Unexpected workbook call'); };
  const session: WorkbookSession = {
    ...testSessionCommands(),
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, sheets: [], activeSheet: 0 },
    call: { version: unexpected, readCells: unexpected, findText: unexpected, validateEdits: unexpected,
      applyEdits: unexpected, command: unexpected, frame: unexpected, sheetView: unexpected,
      cellGeometry: unexpected, cellInputs: unexpected, captureFormat: unexpected, sheets: unexpected,
      calculationStatus: unexpected, save: unexpected,
      collaborationStateVector: unexpected, collaborationStateAsUpdate: unexpected,
      renderPng: unexpected, renderRangePng: unexpected, exportStructured: unexpected,
      exportMarkdown: unexpected, printDisplayList: unexpected },
    save: unexpected, on: () => () => {}, onFailure: () => () => {}, failure: undefined, dispose: async () => {},
  };
  const recovery = createWorkbookSessionRecovery(new Uint8Array([1]), {}, async () => ({
    state: new Uint8Array([1]), epoch: 1, version: 0, throughSeq: 64, encodeMs,
  }), 1);
  for (let seq = 1; seq <= 128; seq += 1) recovery.admit({ session: 'book', incarnation: 1, seq,
    basis: { epoch: 1, version: 0 }, kind: 'setReadOnly', payload: { payload: { readOnly: false } } });
  await recovery.recovery.capture(() => true);
  workbookSessionRecovery.set(session, recovery.recovery);
  let now = 0;
  const timers = new Map<number, { callback(): void; delay: number }>();
  let nextTimer = 0;
  const options = { changed: () => {}, onError: () => {}, onReady: () => {}, isCurrent: () => true,
    checkpointScheduler: {
      now: () => now,
      schedule(callback: () => void, delay: number) {
        const id = ++nextTimer;
        timers.set(id, { callback, delay });
        return () => { timers.delete(id); };
      },
    } };
  const run = new PeerlessWorkbookSession(session, 1, options);
  const capture = spyOn(recovery.recovery, 'capture');
  const fire = (elapsed: number) => {
    now += elapsed;
    const entry = [...timers.entries()][0];
    if (!entry) throw new Error('Missing checkpoint timer');
    timers.delete(entry[0]);
    entry[1].callback();
  };
  try {
    run.connectRecoveryInput({ draft: null, pending: false }, { compositionId: undefined, subscribe: () => () => {} });
    run.firstPaint();
    expect([...timers.values()].map((timer) => timer.delay)).toEqual([Math.max(150, 2 * encodeMs)]);
    now += 10;
    run.connectRecoveryInput({ draft: null, pending: false }, { compositionId: undefined, subscribe: () => () => {} });
    expect([...timers.values()].map((timer) => timer.delay)).toEqual([Math.max(150, 2 * encodeMs)]);
    fire(Math.max(150, 2 * encodeMs) - 1);
    expect(capture).not.toHaveBeenCalled();
    expect([...timers.values()].map((timer) => timer.delay)).toEqual([1]);
    fire(1);
    expect(capture).toHaveBeenCalledTimes(1);
    await recovery.recovery.capture(() => true);
  } finally { capture.mockRestore(); await run.dispose(); }
});
