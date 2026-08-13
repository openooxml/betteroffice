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

const IDENTITY_KEY = "betteroffice:collaboration-user";
const ADJECTIVES = [
  "Bright",
  "Calm",
  "Clever",
  "Kind",
  "Merry",
  "Nimble",
  "Quiet",
  "Swift",
] as const;
const ANIMALS = [
  "Badger",
  "Falcon",
  "Fox",
  "Koala",
  "Otter",
  "Panda",
  "Robin",
  "Tiger",
] as const;

export interface DemoCollaborationUser {
  name: string;
}

function generatedIdentity(): DemoCollaborationUser {
  const values = crypto.getRandomValues(new Uint32Array(2));
  return {
    name: `${ADJECTIVES[values[0] % ADJECTIVES.length]} ${
      ANIMALS[values[1] % ANIMALS.length]
    }`,
  };
}

export function useDemoIdentity(): DemoCollaborationUser | null {
  const [user, setUser] = useState<DemoCollaborationUser | null>(null);
  useEffect(() => {
    let identity: DemoCollaborationUser | null = null;
    try {
      const stored = sessionStorage.getItem(IDENTITY_KEY);
      if (stored) {
        const parsed = JSON.parse(stored) as Partial<DemoCollaborationUser>;
        if (typeof parsed.name === "string" && parsed.name.trim()) {
          identity = { name: parsed.name.trim() };
        }
      }
    } catch {}
    identity ??= generatedIdentity();
    try {
      sessionStorage.setItem(IDENTITY_KEY, JSON.stringify(identity));
    } catch {}
    setUser(identity);
  }, []);
  return user;
}

/** `enabled` false while the tab holds an unshared document: no room, none minted. */
export function useDemoRoom(enabled = true): string | null {
  const pathname = usePathname();
  const router = useRouter();
  const searchParams = useSearchParams();
  const room = searchParams.get("room");
  const generatedRoom = useRef<string | null>(null);

  useEffect(() => {
    if (!enabled || room) return;
    generatedRoom.current ??= crypto.randomUUID();
    const next = new URLSearchParams(searchParams.toString());
    next.set("room", generatedRoom.current);
    router.replace(`${pathname}?${next.toString()}`, { scroll: false });
  }, [enabled, pathname, room, router, searchParams]);

  return enabled ? room : null;
}

/** Drops the room from the URL, so the link stops advertising a document it no longer shows. */
export function useLeaveRoom(): () => void {
  const pathname = usePathname();
  const router = useRouter();
  const searchParams = useSearchParams();

  return useCallback(() => {
    const next = new URLSearchParams(searchParams.toString());
    next.delete("room");
    const query = next.toString();
    router.replace(query ? `${pathname}?${query}` : pathname, { scroll: false });
  }, [pathname, router, searchParams]);
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

  // Per room: a client id only identifies a replica within one room, and
  // carrying one across rooms collides with item ids the new peers already hold.
  useEffect(() => setClientId(roomId ? createClientId() : null), [roomId]);

  const teardown = useCallback(() => {
    for (const cleanup of cleanupRef.current.splice(0)) cleanup();
    const provider = providerRef.current;
    const transport = transportRef.current;
    providerRef.current = null;
    setProvider(null);
    transportRef.current = null;
    setProvider(null);
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
