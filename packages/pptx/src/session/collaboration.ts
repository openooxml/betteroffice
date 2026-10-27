import { createCollaborationPipeHost } from '../../../../shared/collaboration-pipe/worker';
import { CollaborationProvider, providerArrival, providerLeave } from '../collaboration/provider';
import { CollaborationError, type PptxPresenceCursor, type PptxPresencePeer } from '../collaboration/types';
import type { PresentationHandle } from '../wasm/loader';

type WorkerProvider = CollaborationProvider & { onPeers: CollaborationProvider['onPresence'] };

export function createPresentationCollaboration(opened: PresentationHandle, onOutput?: () => void) {
  opened.setCollaborative?.();
  return createCollaborationPipeHost<WorkerProvider, PptxPresenceCursor, PptxPresencePeer>({
    clientId: opened.clientId,
    read: { encodeStateVector: () => opened.encodeStateVector(),
      encodeStateAsUpdate: (vector) => opened.encodeStateAsUpdate(vector) },
    onUpdate: (listener) => opened.onUpdate(listener),
    merge(updates) {
      for (const update of updates) {
        const before = opened.version();
        try { opened.applyUpdate(update); }
        catch (error) {
          if (!(error instanceof Error) || error instanceof WebAssembly.RuntimeError || opened.version() !== before) throw error;
          return { failed: error };
        }
      }
      return { applied: null };
    },
    createProvider(replica, transport, options) {
      const provider = new CollaborationProvider(replica, transport, options);
      return Object.assign(provider, { onPeers: (listener: Parameters<CollaborationProvider['onPresence']>[0]) =>
        provider.onPresence(listener) });
    },
    arrive: providerArrival,
    leave: providerLeave,
    transportError: (error) => new CollaborationError(error.code, error.message),
    onOutput,
  });
}
