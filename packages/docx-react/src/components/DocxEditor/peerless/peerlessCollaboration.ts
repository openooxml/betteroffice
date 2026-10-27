import { attachCollaborationPipe, createAsyncCollaborationReplica, type AsyncReplicaAdapter, type CollaborationTextInsertion,
  type CollaborationUpdateOrigin } from '@betteroffice/docx/collaboration';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/worker';
import { flushAdmittedInput, type PeerlessInputAdmission } from './inputAdmission';
import type { PipeOutput } from '../../../../../../shared/collaboration-pipe/protocol';

const owned = (update: Uint8Array): ArrayBuffer =>
  update.buffer instanceof ArrayBuffer && update.byteOffset === 0 && update.byteLength === update.buffer.byteLength
    ? update.buffer : update.slice().buffer;

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
  const merged = async (ticket: ReturnType<typeof session.submit>) => {
    const outcome = await ticket.ack;
    requireActive();
    if (outcome.status === 'refused' || outcome.status === 'cancelled') throw new Error(`Collaboration merge ${outcome.status}: ${outcome.reason}`);
    return ticket;
  };
  const adapter = createAsyncCollaborationReplica({
    clientId,
    encodeStateVector: () => fenced(() => client.documentRead({ kind: 'collabStateVector' })),
    encodeStateAsUpdate: (stateVector) => fenced(() => client.documentRead({ kind: 'collabDiff', stateVector: stateVector?.slice() })),
    encodeStateAsUpdateChunks: (stateVector, maxChunkBytes) =>
      fenced(() => client.documentRead({ kind: 'collabDiffChunks', stateVector: stateVector.slice(), maxChunkBytes })),
    async applyUpdate(update) {
      requireActive();
      const ticket = await merged(session.submit('applyRemoteUpdate', { update: owned(update) }, { lane: 'collab', fence: true }));
      const inference = ticket.seq === undefined ? undefined : inferences.get(ticket.seq);
      if (ticket.seq !== undefined) inferences.delete(ticket.seq);
      return inference ?? null;
    },
    async applyUpdateChunks(updates) {
      requireActive();
      await merged(session.submit('applyRemoteUpdateChunks', { updates: updates.map(owned) }, { lane: 'collab', fence: true }));
      return null;
    },
    onUpdate(listener) { requireActive(); listeners.add(listener); return () => { listeners.delete(listener); }; },
  });
  const pipes = new Map<number, { receive: (out: readonly PipeOutput[]) => void; lost: (reason: string) => void }>();
  let nextPipe = 1;
  const offPipe = client.onCollabPipe(batches => {
    if (retired) return;
    for (const batch of batches) pipes.get(batch.provider)?.receive(batch.out);
  });
  attachCollaborationPipe(adapter, {
    open(receive, lost) {
      const id = nextPipe++;
      if (retired) queueMicrotask(() => lost('Collaboration replica is retired'));
      else pipes.set(id, { receive, lost });
      return {
        async post(ops) {
          requireActive();
          await merged(session.submit('collabPipe', { provider: id, ops }, { lane: 'collab', fence: true }));
        },
        close() { pipes.delete(id); },
      };
    },
  });
  return Object.assign(adapter, { retire() {
    if (retired) return;
    retired = true; off(); offPipe(); listeners.clear(); inferences.clear();
    const lost = [...pipes.values()]; pipes.clear();
    for (const pipe of lost) { try { pipe.lost('Collaboration replica is retired'); } catch {} }
  } });
}
