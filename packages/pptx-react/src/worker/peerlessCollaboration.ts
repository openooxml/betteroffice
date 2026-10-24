import { type AsyncCollaborationReplica } from '@betteroffice/pptx';
import { attachCollaborationPipe } from '@betteroffice/pptx/collaboration';
import { presentationCollaborationPipes } from '@betteroffice/pptx/worker';
import type { PeerlessPresentation } from './usePeerlessPresentation';

const retirements = new WeakMap<AsyncCollaborationReplica, () => void>();

export function retirePeerlessReplica(replica: AsyncCollaborationReplica): void {
  retirements.get(replica)?.();
  retirements.delete(replica);
}

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
  const replica = {
    clientId,
    encodeStateVector: () => read('encodeStateVector'),
    encodeStateAsUpdate: (vector?: Uint8Array) => read('encodeStateAsUpdate', vector),
    applyUpdate: (update: Uint8Array) => apply([update]),
    applyUpdateChunks: (updates: readonly Uint8Array[]) => apply(updates),
    onUpdate(listener) { run.require(); return run.client.subscribeUpdate!(listener); },
  } satisfies AsyncCollaborationReplica & { applyUpdateChunks(updates: readonly Uint8Array[]): Promise<void> };
  const router = presentationCollaborationPipes.get(run.client);
  if (router) {
    let retired = false;
    const channel = router.channel(async (payload) => {
      run.require();
      const ack = await run.store.dispatch({ kind: 'collabPipe', payload: {
        ...payload, inputTime: performance.now(), undoGroup: 'remote' } });
      run.require();
      const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
      if (refused) throw new Error(`Collaboration pipe refused: ${refused.reason}`);
    });
    const endpoints = new Map<(reason: string) => void, { endpoint: ReturnType<typeof channel.open>; seq: number }>();
    attachCollaborationPipe(replica, { open(receive, lost) {
      if (retired) queueMicrotask(() => lost('Collaboration replica is retired'));
      const endpoint = channel.open(receive, lost);
      const state = { endpoint, seq: 0 };
      endpoints.set(lost, state);
      return { post(ops, transfer) {
        if (retired) return Promise.reject(new Error('Collaboration replica is retired'));
        for (const op of ops) state.seq = Math.max(state.seq, op.seq);
        return endpoint.post(ops, transfer);
      }, close() { endpoints.delete(lost); endpoint.close(); } };
    } });
    retirements.set(replica, () => {
      retired = true;
      for (const [lost, state] of endpoints) {
        void state.endpoint.post([{ op: 'destroy', seq: state.seq + 1 }], []).catch(() => {}).then(() => {
          if (!endpoints.delete(lost)) return;
          state.endpoint.close();
          lost('Collaboration replica is retired');
        });
      }
    });
  }
  return replica;
}
