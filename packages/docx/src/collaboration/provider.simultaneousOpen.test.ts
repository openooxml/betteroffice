import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from '../yrs';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { decodeMessages } from './protocol';
import { CollaborationProvider } from './provider';
import type { CollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

const options = { maxFrameBytes: 64 * 1024, syncChunkBytes: 16 * 1024 };

beforeAll(() => {
  preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
});

async function settle(): Promise<void> {
  for (let tick = 0; tick < 32; tick += 1) await new Promise((done) => setTimeout(done, 0));
}

function later<T>(operation: () => T): Promise<T> {
  return new Promise((done, fail) => setTimeout(() => {
    try { done(operation()); } catch (cause) { fail(cause); }
  }, 0));
}

/** A point-to-point link that queues each direction until delivered and loses both queues when a side disconnects. */
function link() {
  const listeners = [0, 1].map(() => new Set<(event: CollaborationTransportEvent) => void>());
  const queues: Uint8Array[][] = [[], []];
  const sent: Uint8Array[] = [];
  let open = false;
  const emit = (side: number, event: CollaborationTransportEvent) => {
    for (const listener of [...listeners[side]]) listener(event);
  };
  const close = () => {
    if (!open) return;
    open = false;
    for (const queue of queues) queue.length = 0;
    for (const side of [0, 1]) emit(side, { type: 'close' });
  };
  const transports = [0, 1].map((side): CollaborationTransport => ({
    connect() {},
    disconnect: close,
    send(data) {
      sent.push(data.slice());
      if (open) queues[1 - side].push(data.slice());
      return true;
    },
    onEvent(listener) {
      listeners[side].add(listener);
      return () => { listeners[side].delete(listener); };
    },
  }));
  const next = (to: number) => (queues[to].length ? decodeMessages(queues[to][0])[0]?.type : undefined);
  const deliver = async (to: number, until?: string, limit = Number.POSITIVE_INFINITY) => {
    let delivered = 0;
    while (delivered < limit && queues[to].length && next(to) !== until) {
      emit(to, { type: 'message', data: queues[to].shift()! });
      delivered += 1;
      await settle();
    }
    return delivered;
  };
  return {
    transports, sent, next, deliver,
    open() {
      open = true;
      for (const side of [0, 1]) emit(side, { type: 'open' });
    },
    async drain() {
      for (let round = 0; round < 64 && (queues[0].length || queues[1].length); round += 1) {
        await deliver(0);
        await deliver(1);
      }
      return queues[0].length + queues[1].length;
    },
  };
}

async function peer(kind: 'sync' | 'async', clientId: number, paragraphs = 0) {
  const session: YrsSession = await createYrsSession({ clientId });
  if (paragraphs) {
    session.loadStories([{
      storyId: 'body',
      paragraphs: Array.from({ length: paragraphs }, (_, index) => ({
        text: `${index} `.padEnd(256, 'lorem ipsum dolor sit amet '),
      })),
    }]);
  }
  const applied = { updates: 0, syncs: [] as number[] };
  const gate = { answer: Promise.resolve() };
  const apply = (update: Uint8Array) => {
    applied.updates += 1;
    return session.applyUpdate(update);
  };
  const applyChunks = (updates: readonly Uint8Array[]) => {
    applied.syncs.push(updates.length);
    session.applyUpdateChunks(updates);
  };
  const replica: CollaborationReplica = kind === 'sync' ? {
    clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (vector) => session.encodeStateAsUpdate(vector),
    encodeStateAsUpdateChunks: (vector, maxChunkBytes) => session.encodeStateAsUpdateChunks(vector, maxChunkBytes),
    applyUpdate: apply,
    applyUpdateChunks: applyChunks,
    onUpdate: (listener) => session.onUpdate(listener),
  } : createAsyncCollaborationReplica({
    clientId,
    encodeStateVector: () => later(() => session.encodeStateVector()),
    encodeStateAsUpdate: (vector) => gate.answer.then(() => later(() => session.encodeStateAsUpdate(vector))),
    encodeStateAsUpdateChunks: (vector, maxChunkBytes) =>
      later(() => session.encodeStateAsUpdateChunks(vector, maxChunkBytes)),
    applyUpdate: (update) => later(() => apply(update)),
    applyUpdateChunks: (updates) => later(() => applyChunks(updates)),
    onUpdate: (listener) => session.onUpdate(listener),
  });
  return { session, replica, applied, gate };
}

const bodyText = (session: YrsSession) => session.paragraphs('body').map((paragraph) => paragraph.text).join('\n');
const requests = (frames: Uint8Array[]) =>
  frames.flatMap((frame) => decodeMessages(frame)).filter((message) => message.type === 'sync-step-1').length;

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas opening a point-to-point link at once`, () => {
    it('join a large replica through its held chunked answer while the answers to each open are delayed', async () => {
      const source = await peer(kind, 401, 1024);
      const joiner = await peer(kind, 402);
      expect(source.session.encodeState().byteLength).toBeGreaterThan(2 * options.maxFrameBytes);
      const net = link();
      const providers = [source, joiner].map((side, index) =>
        new CollaborationProvider(side.replica, net.transports[index], options));
      const errors: string[] = [];
      for (const provider of providers) {
        provider.onError((error) => { errors.push(error.message); });
        provider.connect();
      }
      const empty = joiner.session.encodeStateVector();
      try {
        net.open();
        await settle();
        await net.deliver(1);
        expect(requests(net.sent)).toBe(2);
        await net.deliver(0);
        expect(errors).toEqual([]);

        const held = await net.deliver(1, 'sync-step-2');
        expect(net.next(1)).toBe('sync-step-2');
        expect(held).toBeGreaterThan(1);
        expect(joiner.applied).toEqual({ updates: 0, syncs: [] });
        expect(joiner.session.encodeStateVector()).toEqual(empty);

        await net.deliver(1, undefined, 1);
        expect(joiner.applied.syncs).toEqual([held + 1]);
        expect(await net.drain()).toBe(0);
        expect(requests(net.sent)).toBe(4);
        expect(errors).toEqual([]);
        expect(net.sent.map((frame) => frame.byteLength).filter((bytes) => bytes > options.maxFrameBytes)).toEqual([]);
        expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
        expect([...joiner.session.encodeStateVector()]).toEqual([...source.session.encodeStateVector()]);
        expect(bodyText(joiner.session) === bodyText(source.session)).toBe(true);
      } finally {
        for (const provider of providers) provider.destroy();
        source.session.destroy();
        joiner.session.destroy();
      }
    });

    it('join a large replica in chunks when only the answerer chunks, every frame within the limit', async () => {
      const source = await peer(kind, 401, 1024);
      const joiner = await peer(kind, 402);
      const net = link();
      const providers = [
        new CollaborationProvider(source.replica, net.transports[0], options),
        new CollaborationProvider(joiner.replica, net.transports[1], { maxFrameBytes: options.maxFrameBytes }),
      ];
      const errors: string[] = [];
      for (const provider of providers) {
        provider.onError((error) => { errors.push(error.message); });
        provider.connect();
      }
      try {
        net.open();
        await settle();
        expect(await net.drain()).toBe(0);
        expect(errors).toEqual([]);
        expect(net.sent.map((frame) => frame.byteLength).filter((bytes) => bytes > options.maxFrameBytes)).toEqual([]);
        expect(joiner.applied.syncs).toHaveLength(1);
        expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
        expect([...joiner.session.encodeStateVector()]).toEqual([...source.session.encodeStateVector()]);
        expect(bodyText(joiner.session) === bodyText(source.session)).toBe(true);
      } finally {
        for (const provider of providers) provider.destroy();
        source.session.destroy();
        joiner.session.destroy();
      }
    });

    if (kind === 'async') {
      it('join a large replica when only the answerer chunks and its answer lands while the joiner still encodes its own', async () => {
        const source = await peer(kind, 401, 1024);
        const joiner = await peer(kind, 402);
        const net = link();
        const providers = [
          new CollaborationProvider(source.replica, net.transports[0], options),
          new CollaborationProvider(joiner.replica, net.transports[1], { maxFrameBytes: options.maxFrameBytes }),
        ];
        const errors: string[] = [];
        for (const provider of providers) {
          provider.onError((error) => { errors.push(error.message); });
          provider.connect();
        }
        let answer!: () => void;
        joiner.gate.answer = new Promise((done) => { answer = done; });
        try {
          net.open();
          await settle();
          await net.deliver(1);
          await net.deliver(0);
          const chunks = await net.deliver(1);
          expect(chunks).toBeGreaterThan(1);
          expect(joiner.applied.syncs).toEqual([]);

          answer();
          await settle();
          expect(await net.drain()).toBe(0);
          expect(errors).toEqual([]);
          expect(net.sent.map((frame) => frame.byteLength).filter((bytes) => bytes > options.maxFrameBytes)).toEqual([]);
          expect(joiner.applied.syncs).toEqual([chunks]);
          expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
          expect([...joiner.session.encodeStateVector()]).toEqual([...source.session.encodeStateVector()]);
          expect(bodyText(joiner.session) === bodyText(source.session)).toBe(true);
        } finally {
          for (const provider of providers) provider.destroy();
          source.session.destroy();
          joiner.session.destroy();
        }
      });
    }
  });
}
