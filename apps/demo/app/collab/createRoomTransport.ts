import type {
  CollaborationTransport,
  CollaborationTransportEvent,
} from "./types";

const MAX_BUFFERED_BYTES = 1024 * 1024;
const INITIAL_RETRY_MS = 250;
const MAX_RETRY_MS = 5000;

type TransportListener = (event: CollaborationTransportEvent) => void;
type PeerCountListener = (count: number | null) => void;

export interface RoomTransport extends CollaborationTransport {
  onPeerCount(listener: PeerCountListener): () => void;
}

function roomUrl(relayOrigin: string, roomId: string): string {
  const url = new URL(relayOrigin);
  if (url.protocol === "https:") url.protocol = "wss:";
  if (url.protocol === "http:") url.protocol = "ws:";
  url.pathname = `/room/${encodeURIComponent(roomId)}`;
  url.search = "";
  url.hash = "";
  return url.toString();
}

export function createRoomTransport(
  relayOrigin: string,
  roomId: string,
): RoomTransport {
  const listeners = new Set<TransportListener>();
  const peerListeners = new Set<PeerCountListener>();
  let socket: WebSocket | null = null;
  let wantsConnection = false;
  let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  let drainTimer: ReturnType<typeof setInterval> | null = null;
  let retryMs = INITIAL_RETRY_MS;
  let peerCount: number | null = null;

  const emit = (event: CollaborationTransportEvent) => {
    for (const listener of [...listeners]) listener(event);
  };

  const emitPeerCount = (count: number | null) => {
    peerCount = count;
    for (const listener of [...peerListeners]) listener(count);
  };

  const clearReconnect = () => {
    if (reconnectTimer) clearTimeout(reconnectTimer);
    reconnectTimer = null;
  };

  const clearDrain = () => {
    if (drainTimer) clearInterval(drainTimer);
    drainTimer = null;
  };

  const startDrain = () => {
    if (drainTimer) return;
    drainTimer = setInterval(() => {
      if (!socket || socket.readyState !== WebSocket.OPEN) {
        clearDrain();
        return;
      }
      if (socket.bufferedAmount <= MAX_BUFFERED_BYTES) {
        clearDrain();
        emit({ type: "drain" });
      }
    }, 50);
  };

  const scheduleReconnect = () => {
    if (!wantsConnection || reconnectTimer) return;
    reconnectTimer = setTimeout(() => {
      reconnectTimer = null;
      open();
    }, retryMs);
    retryMs = Math.min(retryMs * 2, MAX_RETRY_MS);
  };

  const open = () => {
    if (
      !wantsConnection ||
      socket?.readyState === WebSocket.OPEN ||
      socket?.readyState === WebSocket.CONNECTING
    ) {
      return;
    }

    const next = new WebSocket(roomUrl(relayOrigin, roomId));
    next.binaryType = "arraybuffer";
    socket = next;
    next.onopen = () => {
      if (socket !== next) return;
      retryMs = INITIAL_RETRY_MS;
      emit({ type: "open" });
    };
    next.onmessage = (event: MessageEvent<ArrayBuffer | string>) => {
      if (socket !== next) return;
      if (typeof event.data === "string") {
        try {
          const message = JSON.parse(event.data) as {
            type?: string;
            count?: unknown;
          };
          if (
            message.type === "peers" &&
            typeof message.count === "number" &&
            Number.isSafeInteger(message.count)
          ) {
            emitPeerCount(message.count);
          }
        } catch {}
        return;
      }
      emit({ type: "message", data: new Uint8Array(event.data) });
    };
    next.onerror = (error) => {
      if (socket === next) emit({ type: "error", error });
    };
    next.onclose = (event) => {
      if (socket !== next) return;
      socket = null;
      clearDrain();
      emitPeerCount(null);
      emit({ type: "close", reason: event.reason });
      scheduleReconnect();
    };
  };

  return {
    connect() {
      wantsConnection = true;
      clearReconnect();
      open();
    },
    disconnect() {
      wantsConnection = false;
      clearReconnect();
      clearDrain();
      emitPeerCount(null);
      const current = socket;
      socket = null;
      current?.close(1000, "Disconnected");
    },
    send(data) {
      if (!socket || socket.readyState !== WebSocket.OPEN) return false;
      if (socket.bufferedAmount > MAX_BUFFERED_BYTES) {
        startDrain();
        return false;
      }
      const owned = new ArrayBuffer(data.byteLength);
      new Uint8Array(owned).set(data);
      socket.send(owned);
      return true;
    },
    onEvent(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    onPeerCount(listener) {
      peerListeners.add(listener);
      listener(peerCount);
      return () => peerListeners.delete(listener);
    },
  };
}
