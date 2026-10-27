import { timingSafeEqual } from "node:crypto";

import { MAX_COLLABORATION_FRAME_BYTES } from "../../../../shared/collaboration-limits";
import type { FileSession } from "../session/fileSession";
import type { SessionRegistry } from "../session/registry";
import type { UserSession } from "../session/userSession";
import {
  awarenessEntries,
  classifyFrame,
  encodeAwarenessRemoval,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdateFrame,
  stripReservedSyncEntries,
} from "./frames";

/** Minimal socket surface the relay needs. */
export interface RelaySocket {
  send(data: Uint8Array): unknown;
  close(code?: number, reason?: string): unknown;
  readonly readyState?: number;
  data?: RelaySocketData;
}

/** Data attached at upgrade time. */
export interface RelaySocketData {
  handle: string;
}

export interface ResolvedHandle {
  session: FileSession;
  user: UserSession;
}

const AWARENESS_RATE_CAPACITY = 30;
const AWARENESS_REFILL_PER_SECOND = 30;

interface SocketState {
  ws: RelaySocket;
  session: FileSession;
  user: UserSession;
  awareness: Map<number, number>;
  awarenessFrames: Map<number, Uint8Array>;
  bucketTokens: number;
  bucketAt: number;
}

function safeReason(reason: string): string {
  return reason.slice(0, 120);
}

/** Finds the file and user session for a handle with a constant time compare. */
export function findByHandle(registry: SessionRegistry, handle: string): ResolvedHandle | null {
  const candidate = Buffer.from(handle);
  let found: ResolvedHandle | null = null;
  for (const session of registry.list()) {
    for (const user of session.users) {
      const expected = Buffer.from(user.id);
      if (expected.length === candidate.length && timingSafeEqual(expected, candidate)) {
        found = { session, user };
      }
    }
  }
  return found;
}

/** Authoritative WebSocket rooms, one per file session. */
export class CollabRelay {
  #registry: SessionRegistry;
  #byUser = new Map<string, SocketState>();
  #bySocket = new Map<RelaySocket, SocketState>();
  #rooms = new Map<string, Set<SocketState>>();

  constructor(registry: SessionRegistry) {
    this.#registry = registry;
  }

  get registry(): SessionRegistry {
    return this.#registry;
  }

  /** Resolves a session handle, or null. */
  resolve(handle: string): ResolvedHandle | null {
    return findByHandle(this.#registry, handle);
  }

  /** Registers a socket, replacing any previous socket for that user. */
  attach(ws: RelaySocket, resolved: ResolvedHandle): SocketState {
    const previous = this.#byUser.get(resolved.user.id);
    if (previous && previous.ws !== ws) {
      this.#remove(previous, false);
      try {
        previous.ws.close(1000, "replaced");
      } catch {}
    }
    const state: SocketState = {
      ws,
      session: resolved.session,
      user: resolved.user,
      awareness: new Map(),
      awarenessFrames: new Map(),
      bucketTokens: AWARENESS_RATE_CAPACITY,
      bucketAt: Date.now(),
    };
    this.#byUser.set(resolved.user.id, state);
    this.#bySocket.set(ws, state);
    let room = this.#rooms.get(resolved.session.fileKey);
    if (!room) {
      room = new Set();
      this.#rooms.set(resolved.session.fileKey, room);
    }
    room.add(state);
    this.#heartbeat(state);
    return state;
  }

  #remove(state: SocketState, broadcastRemoval: boolean): void {
    if (this.#byUser.get(state.user.id) === state) this.#byUser.delete(state.user.id);
    this.#bySocket.delete(state.ws);
    const room = this.#rooms.get(state.session.fileKey);
    if (room) {
      room.delete(state);
      if (room.size === 0) this.#rooms.delete(state.session.fileKey);
    }
    if (broadcastRemoval && state.awareness.size > 0) {
      const removal = encodeAwarenessRemoval(
        [...state.awareness.entries()].map(([clientId, clock]) => ({ clientId, clock })),
      );
      this.#broadcast(state, removal);
    }
  }

  /** Handles socket close and error: the user counts as disconnected. */
  detach(ws: RelaySocket): void {
    const state = this.#bySocket.get(ws);
    if (!state) return;
    const isCurrent = this.#byUser.get(state.user.id) === state;
    this.#remove(state, isCurrent);
    if (!isCurrent) return;
    try {
      void state.session.disconnect(state.user.id).catch(() => {});
    } catch {}
  }

  /** Counts a socket ping or pong as a heartbeat. */
  ping(ws: RelaySocket): void {
    const state = this.#bySocket.get(ws);
    if (state) this.#heartbeat(state);
  }

  #heartbeat(state: SocketState): void {
    try {
      state.session.heartbeat(state.user.id);
    } catch {}
  }

  #close(state: SocketState, code: number, reason: string): void {
    try {
      state.ws.close(code, safeReason(reason));
    } catch {}
  }

  #send(state: SocketState, frame: Uint8Array): void {
    try {
      state.ws.send(frame.slice());
    } catch {}
  }

  #broadcast(from: SocketState, frame: Uint8Array): void {
    const room = this.#rooms.get(from.session.fileKey);
    if (!room) return;
    for (const peer of [...room]) {
      if (peer !== from) this.#send(peer, frame);
    }
  }

  #consumeAwareness(state: SocketState): boolean {
    const now = Date.now();
    const elapsed = (now - state.bucketAt) / 1000;
    if (elapsed > 0) {
      state.bucketTokens = Math.min(
        AWARENESS_RATE_CAPACITY,
        state.bucketTokens + elapsed * AWARENESS_REFILL_PER_SECOND,
      );
      state.bucketAt = now;
    }
    if (state.bucketTokens < 1) return false;
    state.bucketTokens -= 1;
    return true;
  }

  /** Handles one incoming frame. */
  handleMessage(ws: RelaySocket, data: Uint8Array | ArrayBuffer | string): void {
    const state = this.#bySocket.get(ws);
    if (!state) return;
    if (typeof data === "string") {
      this.#close(state, 1003, "binary frames only");
      return;
    }
    const bytes = data instanceof Uint8Array ? data.slice() : new Uint8Array(data.slice(0));
    if (bytes.byteLength > MAX_COLLABORATION_FRAME_BYTES) {
      this.#close(state, 1009, "frame too large");
      return;
    }
    this.#heartbeat(state);
    const classified = classifyFrame(bytes);
    if (classified.kind === "invalid") {
      this.#close(state, 1002, "malformed collaboration frame");
      return;
    }
    if (classified.kind === "oversize-awareness") {
      this.#close(state, 1009, "awareness too large");
      return;
    }
    if (classified.kind === "auth") {
      this.#close(state, 1008, "auth messages are server only");
      return;
    }
    if (classified.hasAwareness && !this.#consumeAwareness(state)) {
      this.#close(state, 1008, "awareness rate exceeded");
      return;
    }
    const replica = state.session.replica;
    for (const message of classified.messages) {
      if (message.type === "sync-step-1") {
        if (!replica) {
          this.#close(state, 1011, "session is not open");
          return;
        }
        try {
          this.#send(state, encodeSyncStep2(replica.encodeStateAsUpdate(stripReservedSyncEntries(message.stateVector))));
          this.#send(state, encodeSyncStep1(replica.encodeStateVector()));
        } catch {
          this.#close(state, 1011, "sync failed");
          return;
        }
      } else if (message.type === "sync-step-2" || message.type === "update") {
        if (!state.user.canWrite || state.session.viewOnly) {
          this.#close(state, 4403, "read only");
          return;
        }
        let result: { accepted: true } | { accepted: false; reason: string };
        try {
          result = state.session.applyUpdate(state.user.id, message.update.slice());
        } catch {
          this.#close(state, 4400, "update rejected, resync");
          return;
        }
        if (!result.accepted) {
          const readOnly = /read only|token expired|not connected|not open/i.test(result.reason);
          this.#close(state, readOnly ? 4403 : 4400, result.reason || "update rejected, resync");
          return;
        }
        this.#broadcast(state, encodeUpdateFrame(message.update));
      } else if (message.type === "awareness") {
        for (const entry of awarenessEntries(message.update)) {
          state.awareness.set(entry.clientId, entry.clock);
          state.awarenessFrames.set(entry.clientId, message.update.slice());
        }
        this.#broadcast(state, bytesForAwareness(message.update));
      } else if (message.type === "query-awareness") {
        const room = this.#rooms.get(state.session.fileKey);
        if (room) {
          for (const peer of room) {
            if (peer === state) continue;
            for (const payload of peer.awarenessFrames.values()) {
              this.#send(state, bytesForAwareness(payload));
            }
          }
        }
      }
    }
  }

  /** Bun WebSocket handlers for this relay. */
  get websocket() {
    return {
      open: (ws: RelaySocket) => {
        const handle = ws.data?.handle;
        const resolved = handle ? this.resolve(handle) : null;
        if (!resolved) {
          try {
            ws.close(4401, "unauthorized");
          } catch {}
          return;
        }
        this.attach(ws, resolved);
      },
      message: (ws: RelaySocket, message: Uint8Array | ArrayBuffer | string) => {
        this.handleMessage(ws, message);
      },
      close: (ws: RelaySocket) => {
        this.detach(ws);
      },
      ping: (ws: RelaySocket) => {
        this.ping(ws);
      },
      pong: (ws: RelaySocket) => {
        this.ping(ws);
      },
    };
  }
}

function bytesForAwareness(update: Uint8Array): Uint8Array {
  const parts = [encodeVarUintLocal(1), encodeVarUintLocal(update.byteLength), update];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.byteLength;
  }
  return out;
}

function encodeVarUintLocal(value: number): Uint8Array {
  const bytes: number[] = [];
  let remaining = value;
  while (remaining >= 128) {
    bytes.push((remaining % 128) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return Uint8Array.from(bytes);
}
