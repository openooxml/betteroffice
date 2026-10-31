import { createCollaborationPipeHost } from '../../../../shared/collaboration-pipe/worker';
import { CollaborationProvider, providerArrival, providerLeave } from '../collaboration/provider';
import { CollaborationError, type AwarenessCursor, type AwarenessPeer } from '../collaboration/types';
import type { WorkbookHandle } from '../wasm/loader';

type WorkerCollaborationProvider = CollaborationProvider & { onPeers: CollaborationProvider['onAwareness'] };

export function createWorkbookCollaborationHost(opened: WorkbookHandle, onOutput: () => void) {
  return createCollaborationPipeHost<WorkerCollaborationProvider, AwarenessCursor, AwarenessPeer>({
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
    createProvider: (replica, transport, providerOptions) => {
      const provider = new CollaborationProvider(replica, transport, providerOptions);
      return Object.assign(provider, { onPeers: (listener: Parameters<CollaborationProvider['onAwareness']>[0]) =>
        provider.onAwareness(listener) });
    },
    arrive: providerArrival,
    leave: providerLeave,
    transportError: (error) => new CollaborationError(error.code, error.message),
    onOutput,
  });
}
