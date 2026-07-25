import { DurableObject } from "cloudflare:workers";
import { MAX_COLLABORATION_FRAME_BYTES } from "../../../shared/collaboration-limits";
import {
  classifyFrame,
  RetainedUpdateLog,
  type LogMutation,
  type RetainedEntry,
} from "./retention";

interface Env {
  ROOMS: DurableObjectNamespace<CollaborationRoom>;
}

const MAX_RETAINED_COUNT = 512;
const UPDATE_PREFIX = "update:";
const SEQ_DIGITS = 16;
const LEGACY_LOG_KEY = "updates";
const ROOM_TTL_MS = 24 * 60 * 60 * 1000;
const TTL_REFRESH_SLACK_MS = 60 * 60 * 1000;

type PeerMessage = { type: "peers"; count: number };

/** Zero-padded so storage's lexicographic key order is replay order. */
function updateKey(seq: number): string {
  return UPDATE_PREFIX + String(seq).padStart(SEQ_DIGITS, "0");
}

function parseSeq(key: string): number | null {
  const digits = key.slice(UPDATE_PREFIX.length);
  if (digits.length !== SEQ_DIGITS || !/^\d+$/.test(digits)) return null;
  const seq = Number(digits);
  return Number.isSafeInteger(seq) ? seq : null;
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

  constructor(state: DurableObjectState, env: Env) {
    super(state, env);
    state.blockConcurrencyWhile(async () => {
      this.expiresAt = await state.storage.getAlarm();
      const stored = await state.storage.list<Uint8Array>({
        prefix: UPDATE_PREFIX,
      });
      const entries: RetainedEntry[] = [];
      const unusable: string[] = [];
      for (const [key, bytes] of stored) {
        const seq = parseSeq(key);
        if (seq === null || !(bytes instanceof Uint8Array)) {
          unusable.push(key);
          continue;
        }
        entries.push({ seq, bytes });
      }
      if ((await state.storage.get(LEGACY_LOG_KEY)) !== undefined) {
        unusable.push(LEGACY_LOG_KEY);
      }

      const repair = this.updates.restore(entries);
      if (unusable.length > 0) await state.storage.delete(unusable);
      if (repair) await this.writeMutation(repair);
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
    this.refreshExpiry();
    this.updates.replay((update) => server.send(update));
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

    this.refreshExpiry();
    if (kind === "document") {
      const mutation = this.updates.retain(bytes);
      if (mutation) this.persistUpdates(mutation);
    }
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

  /** Wipes the room once it has been idle for a full TTL. */
  async alarm(): Promise<void> {
    await this.persist;
    if (this.ctx.getWebSockets().length > 0) {
      this.expiresAt = Date.now() + ROOM_TTL_MS;
      await this.ctx.storage.setAlarm(this.expiresAt);
      return;
    }

    this.updates.clear();
    this.expiresAt = null;
    await this.ctx.storage.deleteAll();
  }

  /** setAlarm is a storage write, so only rewrite a materially stale deadline. */
  private refreshExpiry(): void {
    const deadline = Date.now() + ROOM_TTL_MS;
    if (
      this.expiresAt !== null &&
      deadline - this.expiresAt < TTL_REFRESH_SLACK_MS
    ) {
      return;
    }
    this.expiresAt = deadline;
    this.persist = this.persist.then(() => this.ctx.storage.setAlarm(deadline));
    this.ctx.waitUntil(this.persist);
  }

  private persistUpdates(mutation: LogMutation): void {
    this.persist = this.persist.then(() => this.writeMutation(mutation));
    this.ctx.waitUntil(this.persist);
  }

  private async writeMutation(mutation: LogMutation): Promise<void> {
    if (mutation.puts.length === 1) {
      const [entry] = mutation.puts;
      await this.ctx.storage.put(updateKey(entry.seq), entry.bytes);
    } else if (mutation.puts.length > 1) {
      const batch: Record<string, Uint8Array> = {};
      for (const entry of mutation.puts) batch[updateKey(entry.seq)] = entry.bytes;
      await this.ctx.storage.put(batch);
    }
    if (mutation.deletes.length > 0) {
      await this.ctx.storage.delete(mutation.deletes.map(updateKey));
    }
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
