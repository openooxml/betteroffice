import { createElement, useCallback, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { AsyncReplicaAdapter } from '@betteroffice/docx/collaboration';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxEditorCollaborationOptions } from '../types';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';
import { PeerlessRemotePresence } from './PeerlessRemotePresence';

type Connection = { client: ResidentEngineWorkerClient; binding: PeerlessEditSessionBinding;
  replica: ReturnType<typeof createPeerlessCollaborationReplica>; stop: () => void };

const heldClocks = new Map<number, number>();
const hold = (clientId: number, clock: number) => heldClocks.set(clientId, Math.max(heldClocks.get(clientId) ?? 0, clock));
let warnedFreshId = false;

/** @internal Opens, publishes and retires the worker editor's collaboration replica. */
export function useWorkerCollaboration(options: DocxEditorCollaborationOptions | undefined, source: unknown) {
  const enabled = options !== undefined;
  const openOptions = useMemo(() => enabled ? { clientId: options?.clientId, initialUpdate: options?.initialUpdate } : undefined,
    [enabled, options?.clientId, options?.initialUpdate]);
  const reopenKey = useMemo(() => ({ source, openOptions }), [source, openOptions]);
  const opening = useRef<{ client: ResidentEngineWorkerClient; room: typeof reopenKey; clientId?: number } | null>(null);
  const [connection, setConnection] = useState<Connection | null>(null);
  const current = useRef<Connection | null>(null);
  const publication = useRef<{ replica: AsyncReplicaAdapter; handler: NonNullable<DocxEditorCollaborationOptions['onReplica']> } | null>(null);
  const retire = useCallback(() => {
    current.current?.stop(); current.current?.replica.retire(); current.current = null; opening.current = null;
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
  const begin = useCallback((client: ResidentEngineWorkerClient) => {
    retire();
    const pinned = openOptions?.clientId;
    const collaboration = openOptions && (pinned === undefined
      ? { clientId: crypto.getRandomValues(new Uint32Array(1))[0]!, initialUpdate: openOptions.initialUpdate }
      : { clientId: pinned, initialUpdate: openOptions.initialUpdate, clientClock: heldClocks.get(pinned) });
    opening.current = { client, room: reopenKey, clientId: collaboration?.clientId };
    return collaboration;
  }, [openOptions, reopenKey, retire]);
  const opened = useCallback((client: ResidentEngineWorkerClient, { clientId, clientClock = 0 }: { clientId?: number; clientClock?: number }) => {
    const open = opening.current;
    if (open?.client !== client || open.clientId === undefined || clientId === undefined) return;
    if (clientId === openOptions?.clientId) hold(clientId, clientClock);
    if (clientId === open.clientId) return;
    if (open.clientId === openOptions?.clientId && !warnedFreshId) {
      warnedFreshId = true;
      console.warn(`[DocxEditor] collaboration.clientId ${open.clientId} would reuse clocks this page already used; this session uses a fresh client id.`);
    }
    opening.current = { ...open, clientId };
  }, [openOptions]);
  const ready = useCallback((client: ResidentEngineWorkerClient | null, binding: PeerlessEditSessionBinding) => {
    const open = opening.current;
    if (!client || !openOptions || open?.client !== client || open.room !== reopenKey || open.clientId === undefined ||
      current.current?.binding === binding) return;
    const clientId = open.clientId;
    const stop = clientId === openOptions.clientId ? client.onCollabUpdates(({ clock }) => {
      if (clock !== undefined) hold(clientId, clock);
    }) : () => {};
    const next = { client, binding, replica: createPeerlessCollaborationReplica(client, binding, clientId), stop };
    current.current = next; setConnection(next);
  }, [openOptions, reopenKey]);
  const overlay = (pages: readonly [number, number]) => connection && options?.presence
    ? createElement(PeerlessRemotePresence, { binding: connection.binding, client: connection.client, presence: options.presence, pages }) : null;
  return { openOptions, reopenKey, begin, opened, ready, retire, overlay, connection };
}
