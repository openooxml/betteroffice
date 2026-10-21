import { expect, mock, spyOn, test } from 'bun:test';
import { CollaborationProvider, type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { encodeSyncStep2, encodeUpdate } from '@betteroffice/docx/collaboration/protocol';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';
import { pipeOpBuffers, type PipeOp } from '../../../../../../shared/collaboration-pipe/protocol';

type Merge = { update?: ArrayBuffer; updates?: ArrayBuffer[]; provider?: number; ops?: PipeOp[] };

function transferringReplica() {
  const merges: Merge[] = [];
  const submit = mock((_kind: string, payload: Merge) => {
    const buffers = payload.ops ? pipeOpBuffers(payload.ops) : payload.update ? [payload.update] : payload.updates!;
    merges.push(structuredClone(payload, { transfer: buffers }));
    return { id: merges.length, seq: merges.length, ack: Promise.resolve({ seq: merges.length, status: 'applied' as const }),
      presented: Promise.resolve('presented' as const) };
  });
  const client = { onCollabUpdates: () => () => {}, onCollabPipe: () => () => {}, documentRead: async () => ({ value: Uint8Array.of(0) }) } as unknown as ResidentEngineWorkerClient;
  const session = { submit, flush: async () => {} } as unknown as DocxWorkerEditSession;
  return { merges, replica: createPeerlessCollaborationReplica(client, { session }, 41) };
}

async function until(done: () => boolean, ms = 0) {
  const end = Date.now() + ms;
  for (let tick = 0; (tick < 100 || Date.now() < end) && !done(); tick++) await new Promise(resolve => setTimeout(resolve, 0));
  expect(done()).toBe(true);
}

test('local submits pass through unchanged and the adapter returns the merge inference', async () => {
  let receive!: Parameters<ResidentEngineWorkerClient['onCollabUpdates']>[0];
  const client = { onCollabUpdates(listener: typeof receive) { receive = listener; return () => {}; },
    onCollabPipe: () => () => {} } as unknown as ResidentEngineWorkerClient;
  let resolve!: (outcome: { status: 'applied' }) => void;
  const ticket = { id: 1, seq: 1, ack: new Promise<{ status: 'applied' }>(done => { resolve = done; }),
    presented: Promise.resolve('presented' as const) };
  const submit = mock((..._args: Parameters<DocxWorkerEditSession['submit']>) => ticket as ReturnType<DocxWorkerEditSession['submit']>);
  const session = { submit } as unknown as DocxWorkerEditSession;
  const replica = createPeerlessCollaborationReplica(client, { session }, 41);
  expect(session.submit).toBe(submit);
  const merged = replica.async.applyUpdate(new Uint8Array([1]));
  session.submit('insertText', { text: 'queued' });
  expect(submit.mock.calls[1]).toEqual(['insertText', { text: 'queued' }]);
  const insertion = { clientId: 42, story: 'body', paraId: 'p', endOffset: 4 };
  receive({ updates: [], inferences: [{ seq: 1, insertion }] });
  resolve({ status: 'applied' });
  expect(await merged).toEqual(insertion);
  replica.retire();
});

test('a remote frame travels from the transport to the worker session undecoded, with one copy', async () => {
  const { merges, replica } = transferringReplica();
  let receive!: (event: CollaborationTransportEvent) => void;
  const provider = new CollaborationProvider(replica, { connect() {}, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; } });
  try {
    provider.connect();
    receive({ type: 'open' });
    receive({ type: 'message', data: encodeSyncStep2(Uint8Array.of(0, 0)) });
    await until(() => merges.length === 1);
    expect(merges[0]!.ops!.map(op => op.op === 'event' ? op.event.type : op.op)).toEqual(['create', 'connect', 'open', 'message']);
    const frame = encodeUpdate(Uint8Array.of(1, 2, 3));
    const slice = spyOn(Uint8Array.prototype, 'slice');
    let copies = -1;
    try {
      receive({ type: 'message', data: frame });
      await until(() => merges.length === 2, 1000);
      copies = slice.mock.calls.length;
    } finally { slice.mockRestore(); }
    console.log('[collab-update-copies]', copies);
    expect(copies).toBe(1);
    const [op] = merges[1]!.ops!;
    expect(op?.op === 'event' && op.event.type === 'message' ? [...new Uint8Array(op.event.data)] : null).toEqual([...frame]);
  } finally { provider.destroy(); replica.retire(); }
});

test('a host merge through adapter.async keeps its own buffers usable', async () => {
  const { merges, replica } = transferringReplica();
  const update = Uint8Array.of(1, 2);
  const chunks = [Uint8Array.of(3), Uint8Array.of(4, 5)];
  await replica.async.applyUpdate(update);
  await replica.async.applyUpdateChunks!(chunks);
  expect([...update, ...chunks[0]!, ...chunks[1]!]).toEqual([1, 2, 3, 4, 5]);
  expect(merges.map(merge => merge.update ? [...new Uint8Array(merge.update)] : merge.updates!.map(chunk => [...new Uint8Array(chunk)])))
    .toEqual([[1, 2], [[3], [4, 5]]]);
  replica.retire();
});
