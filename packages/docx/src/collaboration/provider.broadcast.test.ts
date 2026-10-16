import { describe, expect, it } from 'bun:test';
import { createYrsSession, type YrsSession } from '../yrs';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { CollaborationProvider } from './provider';
import type { CollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

class Hub {
  readonly members: HubTransport[] = [];
}

class HubTransport implements CollaborationTransport {
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

function replica(kind: 'sync' | 'async', session: YrsSession): CollaborationReplica {
  if (kind === 'sync') return session;
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

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas on a stateless broadcast transport`, () => {
    it('let the first joiner take a later joiner\'s edits', async () => {
      const hub = new Hub();
      const first = await createYrsSession({ clientId: 1 });
      const second = await createYrsSession({ clientId: 2 });
      const { paraId } = first.createStory('body', 'Seed');
      const providers = [first, second].map((session) =>
        new CollaborationProvider(replica(kind, session), new HubTransport(hub)));
      const errors: unknown[] = [];
      for (const provider of providers) provider.onError((error) => { errors.push(error); });
      try {
        providers[0].connect();
        await settle();
        providers[1].connect();
        await settle();
        expect(text(second)).toBe('Seed');

        second.insertText({ story: 'body', paraId, offset: 4 }, ' from B');
        await settle();
        expect(text(first)).toBe('Seed from B');
        expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
        expect(errors).toEqual([]);
      } finally {
        for (const provider of providers) provider.destroy();
        first.destroy();
        second.destroy();
      }
    });
  });
}
