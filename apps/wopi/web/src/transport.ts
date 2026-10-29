import type { CollaborationTransport, CollaborationTransportEvent } from "../../../../shared/editor-host/collaboration/types";
import { decodeSequenceAck, encodeSequencedFrame, type SocketSaveBarrier } from "../../src/relay/sequence";
import { BASELINE_GENERATION_PROTOCOL } from "../../src/baseline";
import { classifyFrame, SYNC_STEP_2, TOP_LEVEL_SYNC } from "../../src/relay/frames";

/** WebSocket constructor surface, so tests can fake it. */
export interface WebSocketLike {
  binaryType: string;
  readyState: number;
  bufferedAmount: number;
  onopen: (() => void) | null;
  onmessage: ((event: { data: unknown }) => void) | null;
  onerror: ((event: unknown) => void) | null;
  onclose: ((event: { code?: number; reason?: string }) => void) | null;
  send(data: ArrayBuffer): void;
  close(code?: number, reason?: string): void;
}

export interface WopiTransportOptions {
  handle: string;
  baselineGeneration: number;
  origin?: string;
  checkSession?: () => Promise<boolean>;
  createSocket?: (url: string, protocols: string[]) => WebSocketLike;
}

export interface WopiTransport extends CollaborationTransport {
  hasPendingResync(): boolean;
  captureSaveBarrier(): SocketSaveBarrier | null;
  hasUnacknowledgedUpdates(): boolean;
  noteLocalUpdate(): void;
}

const OPEN = 1;
const CONNECTING = 0;
const MAX_BUFFERED_BYTES = 1024 * 1024;
const INITIAL_RETRY_MS = 250;
const MAX_RETRY_MS = 5000;

/** The /collab endpoint for a page origin. */
export function collabUrl(origin: string): string {
  const url = new URL(origin);
  url.protocol = url.protocol === "https:" ? "wss:" : "ws:";
  url.pathname = "/collab";
  url.search = "";
  url.hash = "";
  return url.toString();
}

/** Collaboration transport over /collab with the session handle as subprotocol. */
export function createWopiTransport(options: WopiTransportOptions): WopiTransport {
  const origin = options.origin ?? (typeof location !== "undefined" ? location.origin : "http://localhost");
  const createSocket = options.createSocket ?? ((url: string, protocols: string[]) => new WebSocket(url, protocols) as unknown as WebSocketLike);
  const url = collabUrl(origin);
  const protocols = ["bo-wopi.v1", options.handle, `${BASELINE_GENERATION_PROTOCOL}${options.baselineGeneration}`];
  const listeners = new Set<(event: CollaborationTransportEvent) => void>();
  let socket: WebSocketLike | null = null;
  let connectionId: string | null = null;
  let seq = 0;
  let resynced = false;
  let wantsConnection = false;
  let rejoinRequired = false;
  let unacknowledged = false;
  let pendingUpdateSeq: number | null = null;
  let localVersion = 0;
  let pendingUpdateVersion = 0;
  let retryMs = INITIAL_RETRY_MS;
  let drainTimer: ReturnType<typeof setTimeout> | null = null;
  let reconnectTimer: ReturnType<typeof setTimeout> | null = null;

  const emit = (event: CollaborationTransportEvent): void => {
    for (const listener of [...listeners]) listener(event);
  };

  const stopDrain = (): void => {
    if (drainTimer !== null) clearTimeout(drainTimer);
    drainTimer = null;
  };

  const watchDrain = (): void => {
    if (drainTimer !== null) return;
    drainTimer = setTimeout(() => {
      drainTimer = null;
      if (!wantsConnection || !socket || socket.readyState !== OPEN) return;
      if (socket.bufferedAmount > MAX_BUFFERED_BYTES) watchDrain();
      else emit({ type: "drain" });
    }, 50);
  };

  const scheduleReconnect = (): void => {
    if (!wantsConnection || reconnectTimer !== null) return;
    reconnectTimer = setTimeout(() => {
      reconnectTimer = null;
      open();
    }, retryMs);
    retryMs = Math.min(retryMs * 2, MAX_RETRY_MS);
  };

  const open = (): void => {
    if (!wantsConnection || socket?.readyState === OPEN || socket?.readyState === CONNECTING) return;
    const next = createSocket(url, protocols);
    next.binaryType = "arraybuffer";
    socket = next;
    connectionId = crypto.randomUUID();
    seq = 0;
    resynced = false;
    let opened = false;
    next.onopen = () => {
      if (socket !== next) return;
      opened = true;
      retryMs = INITIAL_RETRY_MS;
      next.send(encodeSequencedFrame({ connectionId: connectionId!, seq: 0 }).buffer as ArrayBuffer);
      emit({ type: "open" });
    };
    const receive = (data: Uint8Array): void => {
      const ack = decodeSequenceAck(data);
      if (ack) {
        if (ack.connectionId === connectionId && ack.seq <= seq && pendingUpdateSeq !== null && ack.seq >= pendingUpdateSeq && pendingUpdateVersion === localVersion) {
          unacknowledged = false;
          pendingUpdateSeq = null;
        }
        return;
      }
      emit({ type: "message", data });
    };
    next.onmessage = (event) => {
      if (socket !== next) return;
      if (typeof event.data === "string") return;
      if (event.data instanceof ArrayBuffer) {
        receive(new Uint8Array(event.data));
      } else if (event.data instanceof Uint8Array) {
        receive(event.data);
      }
    };
    next.onclose = (event) => {
      if (socket !== next) return;
      stopDrain();
      socket = null;
      connectionId = null;
      seq = 0;
      pendingUpdateSeq = null;
      if (event.code === 4409 || event.code === 4401) {
        rejoinRequired = true;
        wantsConnection = false;
      }
      emit({ type: "close", code: event.code, reason: event.reason });
      if (!opened && wantsConnection && options.checkSession) {
        void options.checkSession().then((active) => {
          if (!wantsConnection || socket !== null) return;
          if (!active) {
            rejoinRequired = true;
            wantsConnection = false;
            emit({ type: "close", code: 4401, reason: "session ended" });
          } else scheduleReconnect();
        }).catch(() => scheduleReconnect());
      } else scheduleReconnect();
    };
  };

  return {
    connect() {
      if (rejoinRequired) return;
      wantsConnection = true;
      if (reconnectTimer !== null) {
        clearTimeout(reconnectTimer);
        reconnectTimer = null;
      }
      open();
    },
    disconnect() {
      wantsConnection = false;
      stopDrain();
      if (reconnectTimer !== null) {
        clearTimeout(reconnectTimer);
        reconnectTimer = null;
      }
      const current = socket;
      socket = null;
      connectionId = null;
      seq = 0;
      current?.close(1000, "Disconnected");
    },
    send(data) {
      if (!socket || socket.readyState !== OPEN) return false;
      if (socket.bufferedAmount > MAX_BUFFERED_BYTES) {
        watchDrain();
        return false;
      }
      const containsUpdate = classifyFrame(data).messages.some((message) =>
        message.type === "update" || (message.type === "sync-step-2" && unacknowledged));
      if (containsUpdate) {
        unacknowledged = true;
        pendingUpdateSeq = seq + 1;
        pendingUpdateVersion = localVersion;
      }
      const frame = encodeSequencedFrame({ connectionId: connectionId!, seq: seq + 1 }, data);
      socket.send(frame.buffer as ArrayBuffer);
      seq += 1;
      if (data[0] === TOP_LEVEL_SYNC && data[1] === SYNC_STEP_2) resynced = true;
      if (socket.bufferedAmount > MAX_BUFFERED_BYTES) watchDrain();
      return true;
    },
    noteLocalUpdate() { localVersion += 1; unacknowledged = true; },
    hasUnacknowledgedUpdates() { return unacknowledged; },
    hasPendingResync() {
      return socket !== null && !resynced;
    },
    captureSaveBarrier() {
      if (!socket) return null;
      if (socket.readyState !== OPEN || !connectionId) throw new Error("notSynced");
      return { connectionId, seq };
    },
    onEvent(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  };
}
