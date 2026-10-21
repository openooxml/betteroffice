import { expect, test } from 'bun:test';
import { createPeerlessReplica } from './peerlessCollaboration';
import type { PeerlessPresentation } from './usePeerlessPresentation';

function presentation() {
  const dispatched: { kind: string; payload: { bytes: ArrayBuffer | readonly ArrayBuffer[]; undoGroup: string } }[] = [];
  const run = {
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
