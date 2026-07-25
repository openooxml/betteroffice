export type CollaborationStatus =
  | "disconnected"
  | "connecting"
  | "connected"
  | "destroyed";

export type CollaborationTransportEvent =
  | { type: "open" }
  | { type: "message"; data: Uint8Array }
  | { type: "close"; reason?: string }
  | { type: "error"; error: unknown }
  | { type: "drain" };

export interface CollaborationTransport {
  connect(): void | Promise<void>;
  disconnect(): void | Promise<void>;
  send(data: Uint8Array): boolean;
  onEvent(
    listener: (event: CollaborationTransportEvent) => void,
  ): () => void;
}

export interface CollaborationReplica {
  readonly clientId: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): unknown;
  onUpdate(
    listener: (update: Uint8Array, origin: "local" | "remote") => void,
  ): () => void;
}

export interface CollaborationProvider {
  connect(): void | Promise<void>;
  destroy(): void | Promise<void>;
  onStatus(
    listener: (change: {
      status: CollaborationStatus;
      synced: boolean;
    }) => void,
  ): () => void;
  onError?(listener: (error: Error) => void): () => void;
}

export type CollaborationProviderFactory<
  Provider extends CollaborationProvider = CollaborationProvider,
> = (
  replica: CollaborationReplica,
  transport: CollaborationTransport,
) => Provider;
