import { beforeAll, expect, test } from 'bun:test';
import { isClientMessage } from '../../../../shared/office-session';
import { hydratePeer } from './client';
import { createWorkbookEditPeer, type WorkbookEditPeer } from './editPeer';
import type { WorkbookReplayEnvelope } from './replay';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

let fixture: Uint8Array;
let wasmBytes: Uint8Array<ArrayBuffer>;

beforeAll(async () => {
  ({ fixture, wasmBytes } = await loadWorkbookSessionFixtures());
});

function deferred() {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => { resolve = done; });
  return { promise, resolve };
}

test('waits for input queued during edit-peer hydration and flush through the final acknowledgement', async () => {
  const acknowledgement = [deferred(), deferred()];
  const entered = [deferred(), deferred()];
  const session = await createTestWorkbookSession(fixture, (transport) => ({
    ...transport,
    listen: (listener) => transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
        const { sequence } = message.args[0] as WorkbookReplayEnvelope;
        entered[sequence - 1].resolve();
        void acknowledgement[sequence - 1].promise.then(() => listener(message));
      } else listener(message);
    }),
  }), { wasm: wasmBytes.buffer, retainPeerHydration: true });
  const hydration = hydratePeer(session);
  let edits: WorkbookEditPeer | undefined;
  let flushed = false;
  const flush = hydration.then((peer) => {
    edits = createWorkbookEditPeer({ session, peer });
    return edits.flush();
  });
  void flush.then(() => { flushed = true; }).catch(() => {});
  const input = hydration.then(() => {
    if (!edits) throw new Error('Missing hydrated edit peer');
    expect(edits.editCell(0, 0, 0, 'queued during hydration').applied).toBe(true);
  });
  const peer = await hydration;
  try {
    await input;
    await entered[0].promise;
    await session.call.version();
    if (!edits) throw new Error('Missing hydrated edit peer');
    expect(flushed).toBe(false);
    expect(edits.sentSequence).toBe(1);
    expect(edits.acknowledgedSequence).toBe(0);
    expect(edits.editCell(0, 0, 1, 'queued during acknowledgement').applied).toBe(true);
    await entered[1].promise;
    acknowledgement[0].resolve();
    await session.call.version();
    expect(flushed).toBe(false);
    expect(edits.sentSequence).toBe(2);
    expect(edits.acknowledgedSequence).toBe(1);
    acknowledgement[1].resolve();
    await flush;
    expect(flushed).toBe(true);
    expect(edits.acknowledgedSequence).toBe(edits.sentSequence);
    expect(edits.acknowledgedSequence).toBe(2);
    expect(peer.version()).toBe(await session.call.version());
    expect(peer.save()).toEqual(await session.save());
  } finally {
    for (const held of acknowledgement) held.resolve();
    edits?.dispose();
    peer.dispose();
    await session.dispose();
  }
});
