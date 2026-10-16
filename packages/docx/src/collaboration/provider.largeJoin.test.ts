import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from '../yrs';
import { decodeMessages, DEFAULT_MAX_FRAME_BYTES } from './protocol';
import { CollaborationProvider } from './provider';
import type {
  CollaborationError,
  CollaborationProviderOptions,
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
} from './types';

const MiB = 1024 * 1024;
const HEAVY = { timeout: 600_000 };

class LoopbackTransport implements CollaborationTransport {
  peer: LoopbackTransport | null = null;
  sent: Uint8Array[] = [];
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();

  connect(): void {}

  disconnect(): void {}

  send(data: Uint8Array): boolean {
    this.sent.push(data);
    this.peer?.emit({ type: 'message', data: data.slice() });
    return true;
  }

  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }
}

interface Applied {
  updates: number;
  syncs: number[];
}

/** The session as a replica, counting merges; `atomic: false` hides `applyUpdateChunks`, as a peer from before it. */
function replicaOf(session: YrsSession, atomic = true): { replica: CollaborationReplica; applied: Applied } {
  const applied: Applied = { updates: 0, syncs: [] };
  const replica: CollaborationReplica = {
    clientId: session.clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (vector) => session.encodeStateAsUpdate(vector),
    encodeStateAsUpdateChunks: (vector, maxChunkBytes) =>
      session.encodeStateAsUpdateChunks(vector, maxChunkBytes),
    applyUpdate: (update) => {
      applied.updates += 1;
      return session.applyUpdate(update);
    },
    onUpdate: (listener) => session.onUpdate(listener),
  };
  if (atomic) {
    replica.applyUpdateChunks = (updates) => {
      applied.syncs.push(updates.length);
      session.applyUpdateChunks(updates);
    };
  }
  return { replica, applied };
}

/**
 * Opens the source first, so only the joiner's SyncStep1 is answered: the joiner receives the source's state. The pair
 * is point-to-point, so the source chunks its answer at 1 MiB.
 */
function join(
  source: CollaborationReplica,
  joiner: CollaborationReplica,
  options: CollaborationProviderOptions = {}
) {
  const sourceTransport = new LoopbackTransport();
  const joinerTransport = new LoopbackTransport();
  sourceTransport.peer = joinerTransport;
  joinerTransport.peer = sourceTransport;
  const sourceProvider = new CollaborationProvider(source, sourceTransport, { syncChunkBytes: MiB });
  const joinerProvider = new CollaborationProvider(joiner, joinerTransport, options);
  const errors: CollaborationError[] = [];
  sourceProvider.onError((error) => errors.push(error));
  joinerProvider.onError((error) => errors.push(error));
  sourceProvider.connect();
  joinerProvider.connect();
  sourceTransport.emit({ type: 'open' });
  sourceTransport.sent = [];
  joinerTransport.emit({ type: 'open' });
  const answer = sourceTransport.sent.flatMap((frame) => decodeMessages(frame));
  const { synced, status } = joinerProvider;
  sourceProvider.destroy();
  joinerProvider.destroy();
  return { synced, status, errors, frames: sourceTransport.sent, answer };
}

async function sessionWithState(clientId: number, bytes: number, paragraphBytes: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  const paragraphs = Array.from({ length: Math.ceil(bytes / paragraphBytes) }, (_, index) => ({
    text: `${index} `.padEnd(paragraphBytes, 'lorem ipsum dolor sit amet '),
  }));
  session.loadStories([{ storyId: 'body', paragraphs }]);
  return session;
}

function bodyText(session: YrsSession): string {
  return session.paragraphs('body').map((paragraph) => paragraph.text).join('\n');
}

function expectSameDocument(joiner: YrsSession, source: YrsSession): void {
  expect([...joiner.encodeStateVector()]).toEqual([...source.encodeStateVector()]);
  const text = bodyText(joiner);
  expect(text.length).toBe(bodyText(source).length);
  expect(text === bodyText(source)).toBe(true);
}

describe('chunked initial sync between yrs replicas', () => {
  beforeAll(() =>
    preloadEditWasm(
      new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
    )
  );

  let large: YrsSession | undefined;
  const largeSource = async (): Promise<YrsSession> => {
    large ??= await sessionWithState(1001, 20 * MiB, 1024);
    expect(large.encodeState().byteLength).toBeGreaterThan(DEFAULT_MAX_FRAME_BYTES);
    return large;
  };

  it('joins a room whose state exceeds one frame', async () => {
    const source = await largeSource();
    const joiner = await createYrsSession({ clientId: 1002 });
    const { replica, applied } = replicaOf(joiner);
    const room = join(replicaOf(source).replica, replica);

    expect(room.errors).toEqual([]);
    expect(room.synced).toBe(true);
    expect(applied).toEqual({ updates: 0, syncs: [room.answer.length] });
    expectSameDocument(joiner, source);
    joiner.destroy();
  }, HEAVY);

  it('streams the answer as bounded Update frames closed by a delete-set SyncStep2', async () => {
    const source = await sessionWithState(1003, 6 * MiB, 512);
    const joiner = await createYrsSession({ clientId: 1004 });
    const room = join(replicaOf(source).replica, replicaOf(joiner).replica);

    expect(room.errors).toEqual([]);
    expect(room.synced).toBe(true);
    const types = room.answer.map((message) => message.type);
    expect(types.length).toBeGreaterThan(6);
    expect(types.slice(0, -1).every((type) => type === 'update')).toBe(true);
    expect(types.at(-1)).toBe('sync-step-2');
    for (const frame of room.frames.slice(0, -1)) expect(frame.byteLength).toBeLessThanOrEqual(MiB + 6);
    const last = room.answer.at(-1)!;
    if (last.type !== 'sync-step-2') throw new Error('expected SyncStep2');
    expect(last.update[0]).toBe(0);
    expectSameDocument(joiner, source);
    source.destroy();
    joiner.destroy();
  }, HEAVY);

  it('answers a receiver without applyUpdateChunks with one SyncStep2, as before', async () => {
    const source = await sessionWithState(1009, 6 * MiB, 64 * 1024);
    const joiner = await createYrsSession({ clientId: 1005 });
    const { replica, applied } = replicaOf(joiner, false);
    const room = join(replicaOf(source).replica, replica);

    expect(room.errors).toEqual([]);
    expect(room.synced).toBe(true);
    expect(room.answer.map((message) => message.type)).toEqual(['sync-step-2']);
    expect(applied).toEqual({ updates: 1, syncs: [] });
    expectSameDocument(joiner, source);
    source.destroy();

    const large = await largeSource();
    const old = replicaOf(joiner, false);
    const refused = join(replicaOf(large).replica, old.replica);
    expect(refused.errors.map((error) => error.code)).toEqual(['protocol']);
    expect(refused.synced).toBe(false);
    expect(old.applied).toEqual({ updates: 0, syncs: [] });
    joiner.destroy();
  }, HEAVY);

  describe('a state over the 64 MiB per-update bound', () => {
    let huge: YrsSession | undefined;
    const hugeSource = async (): Promise<YrsSession> => {
      huge ??= await sessionWithState(1006, 70 * MiB, 256 * 1024);
      expect(huge.encodeState().byteLength).toBeGreaterThan(64 * MiB);
      return huge;
    };

    it('joins when the receiver raises maxSyncBytes', async () => {
      const source = await hugeSource();
      const joiner = await createYrsSession({ clientId: 1007 });
      const room = join(replicaOf(source).replica, replicaOf(joiner).replica, {
        maxSyncBytes: 128 * MiB,
      });

      expect(room.errors).toEqual([]);
      expect(room.synced).toBe(true);
      expectSameDocument(joiner, source);
      joiner.destroy();
    }, HEAVY);

    it('fails the connection cleanly at the default maxSyncBytes', async () => {
      const source = await hugeSource();
      const joiner = await createYrsSession({ clientId: 1008 });
      const before = joiner.encodeState();
      const room = join(replicaOf(source).replica, replicaOf(joiner).replica);

      expect(room.errors.map((error) => [error.code, error.message])).toEqual([
        ['backpressure', `Initial sync exceeds ${64 * MiB} bytes`],
      ]);
      expect(room.synced).toBe(false);
      expect(room.status).toBe('disconnected');
      expect(joiner.encodeState()).toEqual(before);
      joiner.destroy();
    }, HEAVY);
  });
});
