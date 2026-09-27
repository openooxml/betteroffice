import { DurableObject } from "cloudflare:workers";
import { MAX_COLLABORATION_FRAME_BYTES } from "../../../shared/collaboration-limits";
import {
  classifyFrame,
  RetainedUpdateLog,
  RoomCapacityError,
} from "./retention";
import { persistMutation, readRoom } from "./persistence";

interface Env {
  ROOMS: DurableObjectNamespace<CollaborationRoom>;
}

const MAX_RETAINED_COUNT = 512;
const ROOM_TTL_MS = 24 * 60 * 60 * 1000;
const TTL_REFRESH_SLACK_MS = 60 * 60 * 1000;

type PeerMessage = { type: "peers"; count: number };

function sendIfOpen(socket: WebSocket, data: Uint8Array | string): void {
  if (socket.readyState !== WebSocket.OPEN) return;
  try { socket.send(data); } catch { closeSocket(socket, 1011, "Connection unavailable"); }
}

function closeSocket(socket: WebSocket, code = 1000, reason = ""): void {
  try { socket.close(code, reason); } catch {}
}

function isWebSocketRequest(request: Request): boolean {
  return request.headers.get("Upgrade")?.toLowerCase() === "websocket";
}

function copyBytes(message: ArrayBuffer | ArrayBufferView): Uint8Array {
  if (message instanceof ArrayBuffer) return new Uint8Array(message.slice(0));
  return new Uint8Array(
    message.buffer.slice(message.byteOffset, message.byteOffset + message.byteLength),
  );
}

export class CollaborationRoom extends DurableObject<Env> {
  private updates = new RetainedUpdateLog(
    MAX_RETAINED_COUNT,
    MAX_COLLABORATION_FRAME_BYTES,
  );
  private persist = Promise.resolve();
  private expiresAt: number | null = null;
  private failed = false;
  private pendingBytes = 0;
  private pendingCount = 0;

  constructor(state: DurableObjectState, env: Env) {
    super(state, env);
    state.blockConcurrencyWhile(async () => {
      this.expiresAt = await state.storage.getAlarm();
      const stored = await readRoom(state.storage);
      const repair = this.updates.restore(stored.entries, stored.checkpoint);
      const compacted = stored.legacy ? this.updates.checkpoint() : null;
      const mutation = compacted ? { ...compacted, deletes: [...(repair?.deletes ?? []), ...compacted.deletes] } : repair;
      if (mutation || stored.legacy) {
        await persistMutation(state.storage, mutation ?? { puts: [], deletes: [] }, stored.legacy);
      }
    });
  }

  async fetch(request: Request): Promise<Response> {
    if (!isWebSocketRequest(request)) {
      return new Response("WebSocket upgrade required", { status: 426 });
    }

    await this.persist;
    if (this.failed) return new Response("Room storage unavailable", { status: 503 });
    const pair = new WebSocketPair();
    const client = pair[0];
    const server = pair[1];
    this.ctx.acceptWebSocket(server);
    await this.enqueue(() => this.refreshExpiry());
    if (this.failed) return new Response("Room storage unavailable", { status: 503 });
    this.updates.replay((update) => sendIfOpen(server, update));
    sendIfOpen(server, this.updates.syncRequest());
    this.broadcastPeerCount();
    return new Response(null, { status: 101, webSocket: client });
  }

  webSocketMessage(
    socket: WebSocket,
    message: ArrayBuffer | string,
  ): void {
    if (typeof message === "string") {
      socket.close(1003, "Binary frames only");
      return;
    }

    const bytes = copyBytes(message);
    if (bytes.byteLength > MAX_COLLABORATION_FRAME_BYTES) {
      socket.close(1009, `Frame exceeds ${MAX_COLLABORATION_FRAME_BYTES} bytes`);
      return;
    }

    const kind = classifyFrame(bytes);
    if (kind === "invalid") {
      socket.close(1002, "Malformed collaboration frame");
      return;
    }
    if (kind === "auth") {
      socket.close(1008, "Auth messages are server-only");
      return;
    }

    if (this.pendingCount >= 1024 || this.pendingBytes + bytes.length > MAX_COLLABORATION_FRAME_BYTES) {
      socket.close(1013, "Room is busy; reconnect to synchronize");
      return;
    }
    this.pendingBytes += bytes.length;
    this.pendingCount++;
    this.enqueue(async () => {
      try {
        let responses: Uint8Array[];
        let mutation;
        try {
          responses = this.updates.responses(bytes);
          mutation = kind === "document" ? this.updates.retain(bytes) : null;
        } catch (error) {
          socket.close(error instanceof RoomCapacityError ? 1009 : 1002,
            error instanceof RoomCapacityError ? "Room checkpoint limit reached" : "Invalid document update");
          return;
        }
        if (mutation) await persistMutation(this.ctx.storage, mutation);
        await this.refreshExpiry();
        for (const response of responses) sendIfOpen(socket, response);
        for (const peer of this.ctx.getWebSockets()) if (peer !== socket) sendIfOpen(peer, bytes.slice());
      } finally { this.pendingBytes -= bytes.length; this.pendingCount--; }
    });
  }

  webSocketClose(
    socket: WebSocket,
    code: number,
    reason: string,
    _wasClean: boolean,
  ): void {
    closeSocket(socket);
    this.broadcastPeerCount();
  }

  webSocketError(socket: WebSocket, _error: unknown): void {
    socket.close(1011, "WebSocket error");
    this.broadcastPeerCount();
  }

  /** Wipes the room once it has been idle for a full TTL. */
  async alarm(): Promise<void> {
    await this.enqueue(async () => {
      if (this.ctx.getWebSockets().some(socket => socket.readyState === WebSocket.OPEN)) {
        this.expiresAt = Date.now() + ROOM_TTL_MS;
        await this.ctx.storage.setAlarm(this.expiresAt);
        return;
      }
      await this.ctx.storage.deleteAll();
      this.updates.clear();
      this.expiresAt = null;
    });
    if (this.failed) throw new Error("Room storage unavailable");
  }

  async checkpoint(): Promise<void> {
    await this.enqueue(async () => {
      const mutation = this.updates.checkpoint();
      if (mutation) await persistMutation(this.ctx.storage, mutation);
    });
    if (this.failed) throw new Error("Room storage unavailable");
  }

  private async refreshExpiry(): Promise<void> {
    const deadline = Date.now() + ROOM_TTL_MS;
    if (this.expiresAt !== null && deadline - this.expiresAt < TTL_REFRESH_SLACK_MS) return;
    await this.ctx.storage.setAlarm(deadline);
    this.expiresAt = deadline;
  }

  private enqueue(action: () => Promise<void>): Promise<void> {
    const task = this.persist.then(async () => {
      if (this.failed) throw new Error("Room storage unavailable");
      await action();
    });
    this.persist = task.catch(error => {
      this.failed = true;
      for (const socket of this.ctx.getWebSockets()) closeSocket(socket, 1011, "Room storage unavailable; reconnect");
      console.error("Relay persistence failed", error);
    });
    this.ctx.waitUntil(this.persist);
    return this.persist;
  }

  private broadcastPeerCount(): void {
    const peers = this.ctx.getWebSockets().filter(socket => socket.readyState === WebSocket.OPEN);
    const message: PeerMessage = { type: "peers", count: peers.length };
    const payload = JSON.stringify(message);
    for (const peer of peers) sendIfOpen(peer, payload);
  }
}

export default {
  fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/") {
      return new Response("ok");
    }

    const match = url.pathname.match(/^\/room\/([^/]+)$/);
    if (!match) return new Response("Not found", { status: 404 });
    if (!isWebSocketRequest(request)) {
      return new Response("WebSocket upgrade required", { status: 426 });
    }

    let roomId: string;
    try {
      roomId = decodeURIComponent(match[1]);
    } catch {
      return new Response("Invalid room", { status: 400 });
    }
    if (!roomId || roomId.length > 128) {
      return new Response("Invalid room", { status: 400 });
    }

    const room = env.ROOMS.get(env.ROOMS.idFromName(roomId));
    return room.fetch(request);
  },
} satisfies ExportedHandler<Env>;
