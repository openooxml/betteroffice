import { expect, test } from 'bun:test';
import { CollaborationProvider, collaborationPipeFor, type CollaborationTransportEvent } from '@betteroffice/pptx';
import { presentationCollaborationPipes } from '@betteroffice/pptx/worker';
import { createPresentationCollaborationPipe } from '../../../pptx/src/session/commandCollaboration';
import { createPresentationCollaboration } from '../../../pptx/src/session/collaboration';
import type { PresentationHandle } from '../../../pptx/src/wasm/loader';
import type { PresentationCommandClient, PresentationCommandIntent } from '@betteroffice/pptx/worker';
import { createPeerlessReplica, retirePeerlessReplica } from './peerlessCollaboration';
import type { PeerlessPresentation } from './usePeerlessPresentation';

async function settle() { for (let turn = 0; turn < 24; turn++) await Promise.resolve(); }

function presentation() {
  const dispatched: { kind: string; payload: { bytes: ArrayBuffer | readonly ArrayBuffer[]; undoGroup: string } }[] = [];
  const run = {
    client: {} as PresentationCommandClient,
    require() {},
    ready: Promise.resolve(),
    store: { async dispatch(intent: (typeof dispatched)[number]) { dispatched.push(intent); return { outcomes: [{ status: 'applied' }] }; } },
  } as unknown as PeerlessPresentation;
  return { run, dispatched };
}

const bytes = (value: ArrayBuffer | readonly ArrayBuffer[]) =>
  (Array.isArray(value) ? value : [value]).map((buffer: ArrayBuffer) => [...new Uint8Array(buffer)]);

test('chunked remote updates dispatch once, in order, as remote history', async () => {
  const { run, dispatched } = presentation();
  await createPeerlessReplica(run, 7).applyUpdateChunks([Uint8Array.of(1, 2), Uint8Array.of(3), Uint8Array.of(4)]);
  expect(dispatched.map((intent) => [intent.kind, intent.payload.undoGroup])).toEqual([['applyRemoteUpdate', 'remote']]);
  expect(bytes(dispatched[0]!.payload.bytes)).toEqual([[1, 2], [3], [4]]);
});

test('a single remote update keeps its one-buffer payload', async () => {
  const { run, dispatched } = presentation();
  await createPeerlessReplica(run, 7).applyUpdate(Uint8Array.of(5));
  expect(dispatched[0]!.payload.bytes).toBeInstanceOf(ArrayBuffer);
  expect(bytes(dispatched[0]!.payload.bytes)).toEqual([[5]]);
});

test('chunked remote updates refuse empty or oversized batches before dispatch', async () => {
  const { run, dispatched } = presentation();
  const replica = createPeerlessReplica(run, 7);
  await expect(replica.applyUpdateChunks([])).rejects.toThrow(RangeError);
  await expect(replica.applyUpdateChunks([Uint8Array.of(1), new Uint8Array()])).rejects.toThrow(RangeError);
  const half = new Uint8Array(8 * 1024 * 1024 + 1);
  await expect(replica.applyUpdateChunks([half, half])).rejects.toThrow(RangeError);
  expect(dispatched).toEqual([]);
});

test('pipe posts dispatch through the view store and retire with the replica', async () => {
  const run = { client: {} as PresentationCommandClient, require() {} } as PeerlessPresentation;
  const dispatched: PresentationCommandIntent[] = [];
  Object.assign(run, { store: { async dispatch(intent: PresentationCommandIntent) {
    dispatched.push(intent);
    return { outcomes: [{ status: 'noop' }] };
  } } });
  const router = createPresentationCollaborationPipe();
  presentationCollaborationPipes.set(run.client, router);
  const replica = createPeerlessReplica(run, 7);
  const channel = collaborationPipeFor(replica)!;
  expect(channel).toBeDefined();
  const losses: string[] = [];
  const output: unknown[] = [];
  const endpoint = channel.open((out) => output.push(...out), (reason) => losses.push(reason));
  const data = Uint8Array.of(1, 2, 3).buffer;
  await endpoint.post([{ op: 'event', seq: 1, event: { type: 'message', data } }], [data]);
  expect(dispatched).toMatchObject([{ kind: 'collabPipe', payload: {
    provider: 1, undoGroup: 'remote', ops: [{ op: 'event', event: { data } }],
  } }]);
  router.receive([{ provider: 1, out: [{ do: 'status', through: 1, status: 'connected', synced: true }] }]);
  expect(output).toEqual([{ do: 'status', through: 1, status: 'connected', synced: true }]);
  retirePeerlessReplica(replica);
  await settle();
  expect(losses).toEqual(['Collaboration replica is retired']);
  await expect(endpoint.post([{ op: 'connect', seq: 2 }], [])).rejects.toThrow('retired');
  endpoint.close();
  presentationCollaborationPipes.delete(run.client);
});

for (const inFlight of [false, true]) test(`retirement destroys the worker provider before closing with inFlight=${inFlight}`, async () => {
  const updates = new Set<unknown>();
  const host = createPresentationCollaboration({ clientId: 7, version: () => 'v0',
    encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
    applyUpdate() {}, onUpdate(listener: Parameters<PresentationHandle['onUpdate']>[0]) {
      updates.add(listener); return () => { updates.delete(listener); };
    },
  } as unknown as PresentationHandle);
  const router = createPresentationCollaborationPipe();
  const run = { client: {} as PresentationCommandClient, require() {} } as PeerlessPresentation;
  let release!: () => void;
  let pending: Promise<void> | undefined;
  const commands: PresentationCommandIntent[] = [];
  Object.assign(run, { store: { async dispatch(intent: PresentationCommandIntent) {
    commands.push(intent);
    if (pending) await pending;
    if (intent.kind !== 'collabPipe') throw new Error('Unexpected command');
    host.run(intent.payload);
    const output = host.drain();
    if (output) router.receive(output.batches);
    return { outcomes: [{ status: 'noop' }] };
  } } });
  presentationCollaborationPipes.set(run.client, router);
  const replica = createPeerlessReplica(run, 7);
  let receive!: (event: CollaborationTransportEvent) => void;
  const provider = new CollaborationProvider(replica, {
    connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; },
  });
  try {
    provider.connect();
    await settle();
    expect(updates.size).toBe(1);
    if (inFlight) {
      pending = new Promise<void>((resolve) => { release = resolve; });
      provider.setCursor({ slideId: 'slide' });
      await settle();
    }
    const before = commands.length;
    provider.destroy();
    retirePeerlessReplica(replica);
    expect(commands.slice(before)).toMatchObject([{ kind: 'collabPipe', payload: { ops: [{ op: 'destroy' }] } }]);
    if (inFlight) {
      expect(updates.size).toBe(1);
      release();
    }
    await settle();
    expect(updates.size).toBe(0);
    expect(provider.status).toBe('destroyed');
    const endpoint = collaborationPipeFor(replica)!.open(() => {}, () => {});
    try { await expect(endpoint.post([{ op: 'connect', seq: 99 }], [])).rejects.toThrow('retired'); }
    finally { endpoint.close(); }
  } finally { release?.(); provider.destroy(); host.dispose(); presentationCollaborationPipes.delete(run.client); }
});
