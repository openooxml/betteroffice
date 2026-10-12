import { useCallback, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createAsyncCollaborationReplica, type AsyncReplicaAdapter, type CollaborationTextInsertion,
  type CollaborationUpdateOrigin } from '@betteroffice/docx/collaboration';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxEditorCollaborationOptions } from '../types';

export function createPeerlessCollaborationReplica<X>(client: ResidentEngineWorkerClient, session: DocxWorkerEditSession<X>, clientId: number): AsyncReplicaAdapter & { retire(): void } {
  let retired = false;
  const listeners = new Set<(update: Uint8Array, origin: CollaborationUpdateOrigin) => void>();
  const inferences = new Map<number, CollaborationTextInsertion>();
  const requireActive = () => { if (retired) throw new Error('Collaboration replica is retired'); };
  const originalSubmit = session.submit;
  let merge: ReturnType<typeof session.submit> | undefined;
  const submit: typeof session.submit = (kind, payload, options) => {
    const ticket = originalSubmit.call(session, kind, payload, kind !== 'applyRemoteUpdate' && merge && (!options?.after || options.after.id < merge.id)
      ? { ...options, after: merge } : options);
    if (kind === 'applyRemoteUpdate') {
      merge = ticket;
      const clear = () => { if (merge === ticket) merge = undefined; };
      void ticket.ack.then(clear, clear);
    }
    return ticket;
  };
  session.submit = submit;
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
    async encodeStateVector() { requireActive(); const reply = await client.documentRead({ kind: 'collabStateVector' }); requireActive(); return reply.value; },
    async encodeStateAsUpdate(stateVector) { requireActive(); const reply = await client.documentRead({ kind: 'collabDiff', stateVector: stateVector?.slice() }); requireActive(); return reply.value; },
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
    retired = true; off(); listeners.clear(); inferences.clear(); merge = undefined;
    if (session.submit === submit) session.submit = originalSubmit;
  } });
}

export function usePeerlessCollaboration(options: DocxEditorCollaborationOptions | undefined, source: unknown) {
  const enabled = options !== undefined;
  const openOptions = useMemo(() => enabled ? { clientId: options?.clientId ?? crypto.getRandomValues(new Uint32Array(1))[0]!,
    initialUpdate: options?.initialUpdate } : undefined, [enabled, options?.clientId, options?.initialUpdate]);
  const room = useMemo(() => ({ source, openOptions }), [source, openOptions]);
  const opening = useRef<{ client: ResidentEngineWorkerClient; room: typeof room } | null>(null);
  const [connection, setConnection] = useState<{ client: ResidentEngineWorkerClient; binding: PeerlessEditSessionBinding;
    replica: ReturnType<typeof createPeerlessCollaborationReplica> } | null>(null);
  const current = useRef<typeof connection>(null);
  const publication = useRef<{ replica: AsyncReplicaAdapter; handler: NonNullable<DocxEditorCollaborationOptions['onReplica']> } | null>(null);
  const retire = useCallback(() => {
    current.current?.replica.retire(); current.current = null; opening.current = null;
    const published = publication.current; publication.current = null;
    published?.handler(null);
    setConnection(null);
  }, []);
  useLayoutEffect(() => retire, [source, openOptions, retire]);
  useLayoutEffect(() => {
    if (!connection || connection !== current.current || !options?.onReplica) return;
    const published = { replica: connection.replica, handler: options.onReplica };
    publication.current = published; published.handler(published.replica);
    return () => { if (publication.current === published) { publication.current = null; published.handler(null); } };
  }, [connection, options?.onReplica]);
  const begin = useCallback((client: ResidentEngineWorkerClient) => { retire(); opening.current = { client, room }; }, [room, retire]);
  const ready = useCallback((client: ResidentEngineWorkerClient, binding: PeerlessEditSessionBinding) => {
    if (!openOptions || opening.current?.client !== client || opening.current.room !== room || current.current?.binding === binding) return;
    const next = { client, binding, replica: createPeerlessCollaborationReplica(client, binding.session, openOptions.clientId) };
    current.current = next; setConnection(next);
  }, [openOptions, room]);
  return { openOptions, begin, ready, retire, connection };
}
