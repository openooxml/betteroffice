import { describe, expect, it } from 'bun:test';
import { createYrsSession, type YrsSession } from '../yrs';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { decodeMessages } from './protocol';
import { CollaborationProvider } from './provider';
import type {
  CollaborationProviderOptions,
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
} from './types';

class Hub {
  readonly members: HubTransport[] = [];
}

class HubTransport implements CollaborationTransport {
  readonly sent: Uint8Array[] = [];
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();
  constructor(private readonly hub: Hub) {}
  connect(): void {
    this.hub.members.push(this);
    this.emit({ type: 'open' });
  }
  disconnect(): void {
    const at = this.hub.members.indexOf(this);
    if (at >= 0) this.hub.members.splice(at, 1);
  }
  send(data: Uint8Array): boolean {
    this.sent.push(data.slice());
    for (const member of this.hub.members) {
      if (member === this) continue;
      const frame = data.slice();
      queueMicrotask(() => member.emit({ type: 'message', data: frame }));
    }
    return true;
  }
  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
  emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }
}

const settle = async () => { for (let turn = 0; turn < 32; turn += 1) await Bun.sleep(0); };
const text = (session: YrsSession) => session.paragraphs('body').map((paragraph) => paragraph.text).join('\n');
const sent = (transport: HubTransport) => transport.sent.flatMap((frame) => decodeMessages(frame));
const maxFrameBytes = 64 * 1024;

function replica(kind: 'sync' | 'async', session: YrsSession, holds: number[]): CollaborationReplica {
  const applyUpdateChunks = (updates: readonly Uint8Array[]) => {
    holds.push(updates.length);
    session.applyUpdateChunks(updates);
  };
  if (kind === 'sync') {
    return {
      clientId: session.clientId,
      encodeStateVector: () => session.encodeStateVector(),
      encodeStateAsUpdate: (vector) => session.encodeStateAsUpdate(vector),
      encodeStateAsUpdateChunks: (vector, maxChunkBytes) => session.encodeStateAsUpdateChunks(vector, maxChunkBytes),
      applyUpdate: (update) => session.applyUpdate(update),
      applyUpdateChunks,
      onUpdate: (listener) => session.onUpdate(listener),
    };
  }
  return createAsyncCollaborationReplica({
    clientId: session.clientId,
    encodeStateVector: async () => session.encodeStateVector(),
    encodeStateAsUpdate: async (vector) => session.encodeStateAsUpdate(vector),
    encodeStateAsUpdateChunks: async (vector, maxChunkBytes) => session.encodeStateAsUpdateChunks(vector, maxChunkBytes),
    applyUpdate: async (update) => session.applyUpdate(update),
    applyUpdateChunks: async (updates) => applyUpdateChunks(updates),
    onUpdate: (listener) => session.onUpdate(listener),
  });
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas on a stateless broadcast transport`, () => {
    for (const [name, answerer] of [
      ['let the first joiner ask back once and take a later joiner\'s edits', {}],
      ['answer the first joiner\'s ask-back in chunks when only the later joiner chunks', { syncChunkBytes: 4096 }],
    ] as [string, CollaborationProviderOptions][]) {
      it(name, async () => {
        const hub = new Hub();
        const first = await createYrsSession({ clientId: 1 });
        const second = await createYrsSession({ clientId: 2 });
        const { paraId } = first.createStory('body', 'Seed');
        first.insertText({ story: 'body', paraId, offset: 4 }, ' '.padEnd(24 * 1024, 'first '));
        const own = first.encodeStateAsUpdate().byteLength;
        second.applyUpdate(first.encodeStateAsUpdate());
        for (let insert = 0; insert < 40; insert += 1) {
          second.insertText({ story: 'body', paraId, offset: 0 }, ''.padEnd(256, 'second '));
        }
        const holds: number[] = [];
        const transports = [new HubTransport(hub), new HubTransport(hub)];
        const providers = [
          new CollaborationProvider(replica(kind, first, holds), transports[0], { maxFrameBytes }),
          new CollaborationProvider(replica(kind, second, []), transports[1], { maxFrameBytes, ...answerer }),
        ];
        const errors: unknown[] = [];
        for (const provider of providers) provider.onError((error) => { errors.push(error); });
        try {
          providers[0].connect();
          await settle();
          providers[1].connect();
          await settle();
          expect(text(first)).toBe(text(second));

          second.insertText({ story: 'body', paraId, offset: 0 }, 'from B ');
          await settle();
          expect(text(first)).toBe(text(second));
          expect(text(first).startsWith('from B second')).toBe(true);
          const requests = (transport: HubTransport) =>
            sent(transport).filter((message) => message.type === 'sync-step-1').length;
          expect(transports.map(requests)).toEqual([2, 1]);
          const answer = sent(transports[1])
            .flatMap((message) => (message.type === 'sync-step-2' || message.type === 'update' ? [message] : []))
            .slice(0, -1);
          expect(answer.length > 1).toBe(answerer.syncChunkBytes !== undefined);
          expect(holds).toEqual(answerer.syncChunkBytes ? [answer.length] : []);
          expect(answer.reduce((bytes, message) => bytes + message.update.byteLength, 0)).toBeLessThan(own);
          expect(transports.flatMap((transport) => transport.sent).every((frame) => frame.byteLength <= maxFrameBytes))
            .toBe(true);
          expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
          expect(errors).toEqual([]);
        } finally {
          for (const provider of providers) provider.destroy();
          first.destroy();
          second.destroy();
        }
      });
    }
  });
}
