import { createAsyncCollaborationReplica, type AsyncReplicaAdapter, type CollaborationTextInsertion,
  type CollaborationUpdateOrigin } from '@betteroffice/docx/collaboration';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import { flushAdmittedInput, type PeerlessInputAdmission } from './inputAdmission';

export function createPeerlessCollaborationReplica<X>(client: ResidentEngineWorkerClient,
  binding: { admission?: PeerlessInputAdmission; session: DocxWorkerEditSession<X> }, clientId: number): AsyncReplicaAdapter & { retire(): void } {
  const { session } = binding;
  let retired = false;
  const listeners = new Set<(update: Uint8Array, origin: CollaborationUpdateOrigin) => void>();
  const inferences = new Map<number, CollaborationTextInsertion>();
  const requireActive = () => { if (retired) throw new Error('Collaboration replica is retired'); };
  const fenced = async <T>(read: () => Promise<{ value: T }>): Promise<T> => {
    requireActive(); await flushAdmittedInput(binding); requireActive();
    const reply = await read(); requireActive(); return reply.value;
  };
  const off = client.onCollabUpdates(event => {
    if (retired) return;
    for (const inference of event.inferences) inferences.set(inference.seq, inference.insertion);
    while (inferences.size > 256) inferences.delete(inferences.keys().next().value!);
    for (const update of event.updates) for (const listener of [...listeners]) {
      try { listener(update, 'local'); } catch {}
    }
  });
  const adapter = createAsyncCollaborationReplica({
    clientId,
    encodeStateVector: () => fenced(() => client.documentRead({ kind: 'collabStateVector' })),
    encodeStateAsUpdate: (stateVector) => fenced(() => client.documentRead({ kind: 'collabDiff', stateVector: stateVector?.slice() })),
    async applyUpdate(update) {
      requireActive();
      const ticket = session.submit('applyRemoteUpdate', { update: update.slice().buffer }, { lane: 'collab', fence: true });
      const outcome = await ticket.ack;
      requireActive();
      if (outcome.status === 'refused' || outcome.status === 'cancelled') throw new Error(`Collaboration merge ${outcome.status}: ${outcome.reason}`);
      const inference = ticket.seq === undefined ? undefined : inferences.get(ticket.seq);
      if (ticket.seq !== undefined) inferences.delete(ticket.seq);
      return inference ?? null;
    },
    onUpdate(listener) { requireActive(); listeners.add(listener); return () => { listeners.delete(listener); }; },
  });
  return Object.assign(adapter, { retire() {
    if (retired) return;
    retired = true; off(); listeners.clear(); inferences.clear();
  } });
}
