import type { AsyncCollaborationReplica } from '@betteroffice/pptx';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function createPeerlessReplica(run: PeerlessPresentation, clientId: number): AsyncCollaborationReplica & {
  applyUpdateChunks(updates: readonly Uint8Array[]): Promise<void>;
} {
  const read = async (method: 'encodeStateVector' | 'encodeStateAsUpdate', vector?: Uint8Array) => {
    run.require();
    await run.store.fence();
    run.require();
    const result = await run.client.access!(method, ...(method === 'encodeStateAsUpdate' ? [vector?.slice()] : []));
    run.require();
    return result;
  };
  const apply = async (updates: readonly Uint8Array[]) => {
    run.require();
    let total = 0;
    for (const update of updates) {
      if (!(update instanceof Uint8Array) || update.byteLength === 0) throw new RangeError('Invalid collaboration update');
      total += update.byteLength;
    }
    if (updates.length === 0 || total > 16 * 1024 * 1024) throw new RangeError('Invalid collaboration update');
    const bytes = updates.map((update) => update.slice().buffer);
    const ack = await run.store.dispatch({ kind: 'applyRemoteUpdate', payload: {
      bytes: bytes.length === 1 ? bytes[0]! : bytes, inputTime: performance.now(), undoGroup: 'remote' } });
    run.require();
    const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
    if (refused) throw new Error(`Collaboration update refused: ${refused.reason}`);
    await run.ready;
    run.require();
  };
  return {
    clientId,
    encodeStateVector: () => read('encodeStateVector'),
    encodeStateAsUpdate: (vector) => read('encodeStateAsUpdate', vector),
    applyUpdate: (update) => apply([update]),
    applyUpdateChunks: (updates) => apply(updates),
    onUpdate(listener) { run.require(); return run.client.subscribeUpdate!(listener); },
  };
}
