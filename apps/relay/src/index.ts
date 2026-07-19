import { DurableObject } from "cloudflare:workers";

interface Env {
  ROOMS: DurableObjectNamespace<CollaborationRoom>;
}

const MAX_RETAINED_BYTES = 16 * 1024 * 1024;
const MAX_RETAINED_COUNT = 512;
const LOG_KEY = "updates";

type PeerMessage = { type: "peers"; count: number };

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
  private updates: Uint8Array[] = [];
  private retainedBytes = 0;
  private persist = Promise.resolve();

  constructor(state: DurableObjectState, env: Env) {
    super(state, env);
    state.blockConcurrencyWhile(async () => {
      this.updates = (await state.storage.get<Uint8Array[]>(LOG_KEY)) ?? [];
      this.retainedBytes = this.updates.reduce(
        (total, update) => total + update.byteLength,
        0,
      );
    });
  }

  async fetch(request: Request): Promise<Response> {
    if (!isWebSocketRequest(request)) {
      return new Response("WebSocket upgrade required", { status: 426 });
    }

    const pair = new WebSocketPair();
    const client = pair[0];
    const server = pair[1];
    this.ctx.acceptWebSocket(server);
    for (const update of this.updates) server.send(update.slice());
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
    if (bytes.byteLength > MAX_RETAINED_BYTES) {
      socket.close(1009, "Frame too large");
      return;
    }

    this.retain(bytes);
    for (const peer of this.ctx.getWebSockets()) {
      if (peer !== socket) peer.send(bytes.slice());
    }
  }

  webSocketClose(
    socket: WebSocket,
    code: number,
    reason: string,
    _wasClean: boolean,
  ): void {
    socket.close(code, reason);
    this.broadcastPeerCount();
  }

  webSocketError(socket: WebSocket, _error: unknown): void {
    socket.close(1011, "WebSocket error");
    this.broadcastPeerCount();
  }

  private retain(update: Uint8Array): void {
    this.updates.push(update.slice());
    this.retainedBytes += update.byteLength;
    while (
      this.updates.length > MAX_RETAINED_COUNT ||
      this.retainedBytes > MAX_RETAINED_BYTES
    ) {
      const removed = this.updates.shift();
      if (removed) this.retainedBytes -= removed.byteLength;
    }
    const snapshot = this.updates.map((entry) => entry.slice());
    this.persist = this.persist.then(() => this.ctx.storage.put(LOG_KEY, snapshot));
    this.ctx.waitUntil(this.persist);
  }

  private broadcastPeerCount(): void {
    const peers = this.ctx.getWebSockets();
    const message: PeerMessage = { type: "peers", count: peers.length };
    const payload = JSON.stringify(message);
    for (const peer of peers) peer.send(payload);
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
