import { expect, mock, test } from 'bun:test';
import { createAsyncCollaborationReplica } from '../../../packages/docx/src/collaboration/asyncReplica';
import { CollaborationProvider } from '../../../packages/docx/src/collaboration/provider';
import type {
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
} from '../../../packages/docx/src/collaboration/types';
import { createYrsSession, type YrsSession } from '../../../packages/docx/src/yrs';

mock.module('cloudflare:workers', () => ({
  DurableObject: class {
    constructor(protected readonly ctx: unknown, protected readonly env: unknown) {}
  },
}));
const { CollaborationRoom } = await import('../src/index');

class Socket {
  readyState = 1;
  private readonly inbox: Uint8Array[] = [];
  private deliver: ((data: Uint8Array) => void) | null = null;

  send(data: Uint8Array | string): void {
    if (typeof data === 'string') return;
    const frame = data.slice();
    if (this.deliver) queueMicrotask(() => this.deliver!(frame));
    else this.inbox.push(frame);
  }

  close(): void {
    this.readyState = 3;
  }

  attach(deliver: (data: Uint8Array) => void): void {
    this.deliver = deliver;
    for (const frame of this.inbox.splice(0)) deliver(frame);
  }
}

(globalThis as { WebSocketPair?: unknown }).WebSocketPair = function () {
  return { 0: new Socket(), 1: new Socket() };
};

function createRoom() {
  const sockets: Socket[] = [];
  const rows = new Map<string, unknown>();
  const pending: Promise<unknown>[] = [];
  const storage = {
    get: async (key: string | string[]) =>
      Array.isArray(key) ? new Map(key.filter(k => rows.has(k)).map(k => [k, rows.get(k)])) : rows.get(key),
    list: async ({ prefix }: { prefix: string }) =>
      new Map([...rows].filter(([key]) => key.startsWith(prefix)).sort()),
    put: async (key: string | Record<string, unknown>, value?: unknown) => {
      for (const [name, bytes] of typeof key === 'string' ? [[key, value] as const] : Object.entries(key)) rows.set(name, bytes);
    },
    delete: async (key: string | string[]) => {
      for (const name of [key].flat()) rows.delete(name);
    },
  };
  let initialization = Promise.resolve();
  const state = {
    storage: {
      ...storage,
      transaction: (run: (tx: typeof storage) => Promise<void>) => run(storage),
      getAlarm: async () => null,
      setAlarm: async () => {},
    },
    blockConcurrencyWhile: (initialize: () => Promise<void>) => { initialization = initialize(); },
    acceptWebSocket: (socket: Socket) => sockets.push(socket),
    getWebSockets: () => sockets,
    waitUntil: (promise: Promise<unknown>) => pending.push(promise),
  };
  const room = new CollaborationRoom(state as never, {} as never);
  return { room, sockets, pending, ready: () => initialization };
}
type Room = ReturnType<typeof createRoom>;

class RelayTransport implements CollaborationTransport {
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();
  private socket: Socket | null = null;

  constructor(private readonly relay: Room) {}

  async connect(): Promise<void> {
    await this.relay.room.fetch(new Request('https://relay.test/room/a', { headers: { Upgrade: 'websocket' } }));
    this.socket = this.relay.sockets.at(-1)!;
    this.emit({ type: 'open' });
    this.socket.attach(data => this.emit({ type: 'message', data }));
  }

  disconnect(): void {}

  send(data: Uint8Array): boolean {
    this.relay.room.webSocketMessage(this.socket as never, data.slice().buffer);
    return true;
  }

  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  private emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }
}

async function settle(relay: Room): Promise<void> {
  for (let seen = -1; seen !== relay.pending.length;) {
    seen = relay.pending.length;
    await Promise.all(relay.pending);
    await Bun.sleep(0);
  }
}

function replicaOf(session: YrsSession, mode: 'sync' | 'async'): CollaborationReplica {
  if (mode === 'sync') return session;
  return createAsyncCollaborationReplica({
    clientId: session.clientId,
    encodeStateVector: async () => session.encodeStateVector(),
    encodeStateAsUpdate: async (vector) => session.encodeStateAsUpdate(vector),
    encodeStateAsUpdateChunks: async (vector, maxChunkBytes) => session.encodeStateAsUpdateChunks(vector, maxChunkBytes),
    applyUpdate: async (update) => session.applyUpdate(update),
    applyUpdateChunks: async (updates) => session.applyUpdateChunks(updates),
    onUpdate: (listener) => session.onUpdate(listener),
  });
}

async function join(relay: Room, session: YrsSession, mode: 'sync' | 'async') {
  const provider = new CollaborationProvider(replicaOf(session, mode), new RelayTransport(relay));
  const errors: unknown[] = [];
  provider.onError(error => errors.push(error));
  provider.connect();
  await settle(relay);
  return { provider, errors };
}

const text = (session: YrsSession) => session.paragraphs('body').map(paragraph => paragraph.text).join('\n');

test.each(['sync', 'async'] as const)('the first joiner of an empty room exchanges edits with a later joiner, %s replica', async mode => {
  const relay = createRoom();
  await relay.ready();
  const first = await createYrsSession({ clientId: 1 });
  const second = await createYrsSession({ clientId: 2 });
  const { paraId } = first.createStory('body', 'Seed');
  const a = await join(relay, first, mode);
  const b = await join(relay, second, mode);
  try {
    expect(text(second)).toBe('Seed');
    second.insertText({ story: 'body', paraId, offset: 4 }, ' from B');
    await settle(relay);
    expect(text(first)).toBe('Seed from B');
    first.insertText({ story: 'body', paraId, offset: 0 }, 'A: ');
    await settle(relay);
    expect(text(second)).toBe('A: Seed from B');
    expect([a.provider.synced, b.provider.synced]).toEqual([true, true]);
    expect([...a.errors, ...b.errors]).toEqual([]);
  } finally {
    a.provider.destroy();
    b.provider.destroy();
    first.destroy();
    second.destroy();
  }
});
