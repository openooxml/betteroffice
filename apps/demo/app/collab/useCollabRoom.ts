"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";
import { createRoomTransport, type RoomTransport } from "./createRoomTransport";
import type {
  CollaborationProvider,
  CollaborationProviderFactory,
  CollaborationReplica,
  CollaborationStatus,
} from "./types";

export const COLLAB_RELAY_ORIGIN =
  process.env.NEXT_PUBLIC_COLLAB_RELAY_ORIGIN ??
  "https://betteroffice-collaboration-relay.elia7.workers.dev";

function createClientId(): number {
  const value = crypto.getRandomValues(new Uint32Array(1))[0] & 0x7fffffff;
  return value || 1;
}

export function useDemoRoom(): string | null {
  const pathname = usePathname();
  const router = useRouter();
  const searchParams = useSearchParams();
  const room = searchParams.get("room");
  const generatedRoom = useRef<string | null>(null);

  useEffect(() => {
    if (room) return;
    generatedRoom.current ??= crypto.randomUUID();
    const next = new URLSearchParams(searchParams.toString());
    next.set("room", generatedRoom.current);
    router.replace(`${pathname}?${next.toString()}`, { scroll: false });
  }, [pathname, room, router, searchParams]);

  return room;
}

export interface CollabRoomState<
  Provider extends CollaborationProvider = CollaborationProvider,
> {
  clientId: number | null;
  status: CollaborationStatus;
  synced: boolean;
  peerCount: number | null;
  error: string | null;
  provider: Provider | null;
  onReplica(replica: CollaborationReplica | null): void;
}

export function useCollabRoom<Provider extends CollaborationProvider>(
  relayOrigin: string,
  roomId: string | null,
  createProvider: CollaborationProviderFactory<Provider>,
): CollabRoomState<Provider> {
  const [clientId, setClientId] = useState<number | null>(null);
  const [status, setStatus] =
    useState<CollaborationStatus>("disconnected");
  const [synced, setSynced] = useState(false);
  const [peerCount, setPeerCount] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [provider, setProvider] = useState<Provider | null>(null);
  const providerRef = useRef<Provider | null>(null);
  const transportRef = useRef<RoomTransport | null>(null);
  const cleanupRef = useRef<Array<() => void>>([]);

  useEffect(() => setClientId(createClientId()), []);

  const teardown = useCallback(() => {
    for (const cleanup of cleanupRef.current.splice(0)) cleanup();
    const provider = providerRef.current;
    const transport = transportRef.current;
    providerRef.current = null;
    setProvider(null);
    transportRef.current = null;
    void provider?.destroy();
    void transport?.disconnect();
    setStatus("disconnected");
    setSynced(false);
    setPeerCount(null);
  }, []);

  useEffect(
    () => teardown,
    [createProvider, relayOrigin, roomId, teardown],
  );

  const onReplica = useCallback(
    (replica: CollaborationReplica | null) => {
      teardown();
      if (!replica || !roomId) return;

      setError(null);
      setStatus("connecting");
      const transport = createRoomTransport(relayOrigin, roomId);
      const provider = createProvider(replica, transport);
      transportRef.current = transport;
      providerRef.current = provider;
      setProvider(provider);
      cleanupRef.current.push(
        transport.onPeerCount(setPeerCount),
        provider.onStatus((change) => {
          setStatus(change.status);
          setSynced(change.synced);
        }),
      );
      if (provider.onError) {
        cleanupRef.current.push(
          provider.onError((nextError) => setError(nextError.message)),
        );
      }
      void provider.connect();
    },
    [createProvider, relayOrigin, roomId, teardown],
  );

  return { clientId, status, synced, peerCount, error, provider, onReplica };
}
