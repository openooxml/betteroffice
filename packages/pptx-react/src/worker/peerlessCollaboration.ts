import type { AsyncCollaborationReplica } from '@betteroffice/pptx';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function createPeerlessReplica(run: PeerlessPresentation, clientId: number): AsyncCollaborationReplica {
  const read = async (method: 'encodeStateVector' | 'encodeStateAsUpdate', vector?: Uint8Array) => {
    run.require();
    await run.store.fence();
    run.require();
    const result = await run.client.access!(method, ...(method === 'encodeStateAsUpdate' ? [vector?.slice()] : []));
    run.require();
    return result;
  };
  return {
    clientId,
    encodeStateVector: () => read('encodeStateVector'),
    encodeStateAsUpdate: (vector) => read('encodeStateAsUpdate', vector),
    async applyUpdate(update) {
      run.require();
      if (!(update instanceof Uint8Array) || update.byteLength === 0 || update.byteLength > 16 * 1024 * 1024) {
        throw new RangeError('Invalid collaboration update');
      }
      const ack = await run.store.dispatch({ kind: 'applyRemoteUpdate', payload: {
        bytes: update.slice().buffer, inputTime: performance.now(), undoGroup: 'remote' } });
      run.require();
      const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
      if (refused) throw new Error(`Collaboration update refused: ${refused.reason}`);
      await run.ready;
      run.require();
    },
    onUpdate(listener) { run.require(); return run.client.subscribeUpdate!(listener); },
  };
}
