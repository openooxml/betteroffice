import { timingSafeEqual } from "node:crypto";

import { MAX_COLLABORATION_FRAME_BYTES } from "../../../../shared/collaboration-limits";
import type { FileSession } from "../session/fileSession";
import type { SessionRegistry } from "../session/registry";
import type { UserSession } from "../session/userSession";
import {
  classifyFrame,
  encodeAwarenessFrameRaw,
  encodeAwarenessRemoval,
  encodeAwarenessUpdateRaw,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdateFrame,
  stripReservedSyncEntries,
  type ParsedAwarenessEntry,
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
  baselineGeneration: number;
}

export interface ResolvedHandle {
  session: FileSession;
  user: UserSession;
}

const AWARENESS_RATE_CAPACITY = 30;
const AWARENESS_REFILL_PER_SECOND = 30;
/** Client ids one socket may keep in awareness; a page announces only its own. */
export const MAX_AWARENESS_CLIENTS_PER_SOCKET = 8;

interface SocketState {
  ws: RelaySocket;
  session: FileSession;
  user: UserSession;
  generation: number;
  awareness: Map<number, number>;
  awarenessFrames: Map<number, Uint8Array>;
  bucketTokens: number;
  bucketAt: number;
}

function safeReason(reason: string): string {
  return reason.slice(0, 120);
}

function isEmptyUpdate(update: Uint8Array): boolean {
  return update.length === 0 || (update.length === 2 && update[0] === 0 && update[1] === 0);
}

export function isSessionUserAuthorized(session: FileSession, user: UserSession): boolean {
  return session.getUser(user.id) === user && session.isUserTokenValid(user.id);
}

/** Finds the file and user session for a handle with a constant time compare. */
export function findByHandle(registry: SessionRegistry, handle: string): ResolvedHandle | null {
  const candidate = Buffer.from(handle);
  let found: ResolvedHandle | null = null;
  for (const session of registry.list()) {
    for (const user of session.users) {
      const expected = Buffer.from(user.id);
      if (expected.length === candidate.length && timingSafeEqual(expected, candidate) && isSessionUserAuthorized(session, user)) {
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
  #subscribed = new Set<FileSession>();

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
  attach(ws: RelaySocket, resolved: ResolvedHandle, baselineGeneration: number | null): SocketState | null {
    if (baselineGeneration !== resolved.session.baselineGeneration) {
      try { ws.close(4409, "rejoin from source"); } catch {}
      return null;
    }
    const generation = resolved.session.connect(resolved.user.id);
    const previous = this.#byUser.get(resolved.user.id);
    if (previous && previous.ws !== ws) {
      this.#remove(previous, true);
      try {
        previous.ws.close(1000, "replaced");
      } catch {}
    }
    if (!this.#subscribed.has(resolved.session)) {
      this.#subscribed.add(resolved.session);
      resolved.session.on("closed", () => this.closeSession(resolved.session, 1000, "session closed"));
      resolved.session.on("userInvalidated", ({ userId }) => this.closeUser(userId, 4401, "unauthorized"));
      resolved.session.on("rejoinRequired", () => this.#closeRoom(resolved.session, 4409, "rejoin from source"));
      resolved.session.on("replicaUpdate", ({ update }) => {
        this.#broadcastAll(resolved.session, encodeUpdateFrame(update));
      });
    }
    const state: SocketState = {
      ws,
      session: resolved.session,
      user: resolved.user,
      generation,
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

  /** Closes one user's socket and removes it from the relay. */
  closeUser(userId: string, code = 1000, reason = "session closed"): void {
    const state = this.#byUser.get(userId);
    if (!state) return;
    this.#remove(state, true);
    this.#close(state, code, reason);
  }

  /** Closes every socket in one file session room. */
  closeSession(session: FileSession, code = 1000, reason = "session closed"): void {
    this.#subscribed.delete(session);
    this.#closeRoom(session, code, reason);
  }

  #closeRoom(session: FileSession, code: number, reason: string): void {
    const room = this.#rooms.get(session.fileKey);
    if (!room) return;
    for (const state of [...room]) {
      this.#remove(state, false);
      this.#close(state, code, reason);
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
      void state.session.disconnect(state.user.id, state.generation).catch(() => {});
    } catch {}
  }

  /** Counts a socket ping or pong as a heartbeat. */
  ping(ws: RelaySocket): void {
    const state = this.#bySocket.get(ws);
    if (state) this.#heartbeat(state);
  }

  #heartbeat(state: SocketState): void {
    if (!this.#authorize(state)) return;
    try {
      state.session.heartbeat(state.user.id, state.generation);
    } catch {}
  }

  #close(state: SocketState, code: number, reason: string): void {
    try {
      state.ws.close(code, safeReason(reason));
    } catch {}
  }

  #send(state: SocketState, frame: Uint8Array): boolean {
    if (!this.#authorize(state)) return false;
    try {
      const sent = state.ws.send(frame.slice());
      if (sent !== 0 && sent !== -1) return true;
    } catch {}
    this.#close(state, 1013, "delivery dropped, reconnect");
    this.detach(state.ws);
    return false;
  }

  #authorize(state: SocketState): boolean {
    if (!this.#bySocket.has(state.ws)) return false;
    if (isSessionUserAuthorized(state.session, state.user)) return true;
    this.closeUser(state.user.id, 4401, "unauthorized");
    return false;
  }

  #broadcast(from: SocketState, frame: Uint8Array): void {
    const room = this.#rooms.get(from.session.fileKey);
    if (!room) return;
    for (const peer of [...room]) {
      if (peer !== from) this.#send(peer, frame);
    }
  }

  #broadcastAll(session: FileSession, frame: Uint8Array): void {
    const room = this.#rooms.get(session.fileKey);
    if (!room) return;
    for (const peer of [...room]) this.#send(peer, frame);
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
    if (!this.#authorize(state)) return;
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
    const canWrite =
      state.user.canWrite && !state.session.viewOnly && state.session.isUserTokenValid(state.user.id);
    for (const message of classified.messages) {
      if (!this.#bySocket.has(ws)) return;
      if (message.type === "sync-step-1") {
        if (!replica) {
          this.#close(state, 1011, "session is not open");
          return;
        }
        try {
          if (!this.#send(state, encodeSyncStep2(replica.encodeStateAsUpdate(stripReservedSyncEntries(message.stateVector))))) return;
          if (canWrite) this.#send(state, encodeSyncStep1(replica.encodeStateVector()));
        } catch {
          this.#close(state, 1011, "sync failed");
          return;
        }
      } else if (message.type === "sync-step-2" || message.type === "update") {
        if (isEmptyUpdate(message.update)) continue;
        if (!canWrite) {
          this.#close(state, 4403, "read only");
          return;
        }
        const before = replica?.version;
        let result: ReturnType<FileSession["applyUpdate"]>;
        try {
          result = state.session.applyUpdate(state.user.id, message.update.slice());
        } catch {
          this.#close(state, 4400, "update rejected, resync");
          return;
        }
        if (result.accepted && replica && replica.version !== before) {
          this.#broadcast(state, encodeUpdateFrame(message.update));
        }
        for (const followUp of result.followUps ?? []) {
          this.#broadcastAll(state.session, encodeUpdateFrame(followUp));
        }
        if (!result.accepted) {
          if (result.reason === "causalGap") {
            if (replica) {
              try {
                if (!this.#send(state, encodeSyncStep1(replica.encodeStateVector()))) return;
              } catch {}
            }
            continue;
          }
          const readOnly = /read only|token expired|not connected|not open/i.test(result.reason);
          this.#close(state, readOnly ? 4403 : 4400, result.reason || "update rejected, resync");
          return;
        }
      } else if (message.type === "awareness") {
        const room = this.#rooms.get(state.session.fileKey);
        const ownedByOther = new Set<number>();
        if (room) {
          for (const peer of room) {
            if (peer === state) continue;
            for (const clientId of peer.awareness.keys()) ownedByOther.add(clientId);
          }
        }
        const accepted: ParsedAwarenessEntry[] = [];
        for (const entry of message.entries) {
          if (ownedByOther.has(entry.clientId)) continue;
          if (entry.isNull) {
            state.awareness.delete(entry.clientId);
            state.awarenessFrames.delete(entry.clientId);
          } else {
            if (!state.awareness.has(entry.clientId) && state.awareness.size >= MAX_AWARENESS_CLIENTS_PER_SOCKET) {
              this.#close(state, 1008, "too many awareness clients");
              this.detach(state.ws);
              return;
            }
            state.awareness.set(entry.clientId, entry.clock);
            state.awarenessFrames.set(entry.clientId, encodeAwarenessUpdateRaw([entry]));
          }
          accepted.push(entry);
        }
        if (accepted.length > 0) this.#broadcast(state, encodeAwarenessFrameRaw(accepted));
      } else if (message.type === "query-awareness") {
        const room = this.#rooms.get(state.session.fileKey);
        if (room) {
          const sent = new Set<number>();
          for (const peer of room) {
            if (peer === state) continue;
            for (const [clientId, payload] of peer.awarenessFrames) {
              if (sent.has(clientId)) continue;
              sent.add(clientId);
              if (!this.#send(state, bytesForAwareness(payload))) return;
            }
          }
        }
      }
    }
  }

  /** Bun WebSocket handlers for this relay. */
  get websocket() {
    return {
      maxPayloadLength: MAX_COLLABORATION_FRAME_BYTES + 4096,
      open: (ws: RelaySocket) => {
        const handle = ws.data?.handle;
        const resolved = handle ? this.resolve(handle) : null;
        if (!resolved) {
          try {
            ws.close(4401, "unauthorized");
          } catch {}
          return;
        }
        this.attach(ws, resolved, ws.data?.baselineGeneration ?? null);
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
