import { createElement, useCallback, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { AsyncReplicaAdapter } from '@betteroffice/docx/collaboration';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxEditorCollaborationOptions } from '../types';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';
import { PeerlessRemotePresence } from './PeerlessRemotePresence';

type Connection = { client: ResidentEngineWorkerClient; binding: PeerlessEditSessionBinding;
  replica: ReturnType<typeof createPeerlessCollaborationReplica> };

/** @internal Opens, publishes and retires the worker editor's collaboration replica. */
export function useWorkerCollaboration(options: DocxEditorCollaborationOptions | undefined, source: unknown) {
  const enabled = options !== undefined;
  const openOptions = useMemo(() => enabled ? { clientId: options?.clientId ?? crypto.getRandomValues(new Uint32Array(1))[0]!,
    initialUpdate: options?.initialUpdate } : undefined, [enabled, options?.clientId, options?.initialUpdate]);
  const reopenKey = useMemo(() => ({ source, openOptions }), [source, openOptions]);
  const opening = useRef<{ client: ResidentEngineWorkerClient; room: typeof reopenKey } | null>(null);
  const [connection, setConnection] = useState<Connection | null>(null);
  const current = useRef<Connection | null>(null);
  const publication = useRef<{ replica: AsyncReplicaAdapter; handler: NonNullable<DocxEditorCollaborationOptions['onReplica']> } | null>(null);
  const retire = useCallback(() => {
    current.current?.replica.retire(); current.current = null; opening.current = null;
    const published = publication.current; publication.current = null;
    published?.handler(null);
    setConnection(null);
  }, []);
  useLayoutEffect(() => retire, [reopenKey, retire]);
  useLayoutEffect(() => {
    if (!connection || connection !== current.current || !options?.onReplica) return;
    const published = { replica: connection.replica, handler: options.onReplica };
    publication.current = published; published.handler(published.replica);
    return () => { if (publication.current === published) { publication.current = null; published.handler(null); } };
  }, [connection, options?.onReplica]);
  const begin = useCallback((client: ResidentEngineWorkerClient) => { retire(); opening.current = { client, room: reopenKey }; }, [reopenKey, retire]);
  const ready = useCallback((client: ResidentEngineWorkerClient | null, binding: PeerlessEditSessionBinding) => {
    if (!client || !openOptions || opening.current?.client !== client || opening.current.room !== reopenKey ||
      current.current?.binding === binding) return;
    const next = { client, binding, replica: createPeerlessCollaborationReplica(client, binding.session, openOptions.clientId) };
    current.current = next; setConnection(next);
  }, [openOptions, reopenKey]);
  const overlay = (pages: readonly [number, number]) => connection && options?.presence
    ? createElement(PeerlessRemotePresence, { binding: connection.binding, client: connection.client, presence: options.presence, pages }) : null;
  return { openOptions, reopenKey, begin, ready, retire, overlay, connection };
}
