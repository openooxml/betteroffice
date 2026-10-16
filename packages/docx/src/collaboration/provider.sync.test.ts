import { afterEach, describe, expect, it } from 'bun:test';
import {
  decodeMessages,
  encodeAwarenessMessage,
  encodeSyncStep2,
  encodeUpdate,
} from './protocol';
import { CollaborationProvider } from './provider';
import type {
  CollaborationError,
  CollaborationPeer,
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
  CollaborationUpdateOrigin,
} from './types';

class SyncReplica implements CollaborationReplica {
  readonly clientId = 1;
  stateVector = Uint8Array.of(7);
  vectorError: unknown;
  applyResult: unknown;
  readonly applied: Uint8Array[] = [];

  encodeStateVector(): Uint8Array {
    if (this.vectorError) throw this.vectorError;
    return this.stateVector;
  }

  encodeStateAsUpdate(): Uint8Array {
    return Uint8Array.of(8);
  }

  applyUpdate(update: Uint8Array): unknown {
    this.applied.push(update);
    return this.applyResult;
  }

  onUpdate(_listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void {
    return () => {};
  }
}

class Transport implements CollaborationTransport {
  readonly sent: Uint8Array[] = [];
  onSend: (() => void) | undefined;
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();

  connect(): void {}
  disconnect(): void {}

  send(frame: Uint8Array): boolean {
    this.sent.push(frame.slice());
    this.onSend?.();
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

const providers = new Set<CollaborationProvider>();

afterEach(() => {
  for (const provider of providers) provider.destroy();
  providers.clear();
});

function setup() {
  const replica = new SyncReplica();
  const transport = new Transport();
  const provider = new CollaborationProvider(replica, transport, { user: { name: 'Calm Otter' } });
  const errors: CollaborationError[] = [];
  const peers: CollaborationPeer[][] = [];
  provider.onError((error) => errors.push(error));
  provider.onPeers((snapshot) => peers.push(snapshot.slice()));
  providers.add(provider);
  provider.connect();
  return { replica, transport, provider, errors, peers };
}

function kinds(frames: readonly Uint8Array[]): string[] {
  return frames.flatMap((frame) => decodeMessages(frame).map((message) => message.type));
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((bytes, part) => bytes + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.byteLength;
  }
  return output;
}

function peerAwareness(clock: number, offset: number): Uint8Array {
  return encodeAwarenessMessage([{
    clientId: 8,
    clock,
    state: {
      user: { name: 'Bright Fox', color: '#137333' },
      cursor: { story: 'body', anchor: Uint8Array.of(offset), head: Uint8Array.of(offset) },
    },
  }]);
}

const typed = { clientId: 8, story: 'body', paraId: 'p1', endOffset: 6 };

describe('CollaborationProvider sync replicas on paths shared with async replicas', () => {
  it('sends Step1, the awareness query and local awareness on every open', () => {
    const { replica, transport, errors } = setup();
    transport.emit({ type: 'open' });
    expect(kinds(transport.sent)).toEqual(['sync-step-1', 'query-awareness', 'awareness']);
    expect(decodeMessages(transport.sent[0])).toEqual([
      { type: 'sync-step-1', stateVector: Uint8Array.of(7) },
    ]);

    transport.emit({ type: 'close' });
    transport.sent.length = 0;
    replica.stateVector = Uint8Array.of(9);
    transport.emit({ type: 'open' });
    expect(kinds(transport.sent)).toEqual(['sync-step-1', 'query-awareness', 'awareness']);
    expect(decodeMessages(transport.sent[0])).toEqual([
      { type: 'sync-step-1', stateVector: Uint8Array.of(9) },
    ]);
    expect(errors).toEqual([]);
  });

  it('stops a handshake superseded by a re-entrant close and reopen', () => {
    const { transport, errors } = setup();
    transport.onSend = () => {
      transport.onSend = undefined;
      transport.emit({ type: 'close' });
      transport.emit({ type: 'open' });
    };
    transport.emit({ type: 'open' });
    expect(kinds(transport.sent)).toEqual([
      'sync-step-1',
      'sync-step-1',
      'query-awareness',
      'awareness',
    ]);
    expect(errors).toEqual([]);
  });

  it('fails with a replica error and sends nothing when the state vector cannot be encoded', () => {
    const { replica, transport, provider, errors } = setup();
    replica.vectorError = new Error('closed');
    transport.emit({ type: 'open' });
    expect(transport.sent).toEqual([]);
    expect(errors.map((error) => error.code)).toEqual(['replica']);
    expect(provider.status).toBe('disconnected');
  });

  it('infers typing from updates but not from SyncStep2', () => {
    const { replica, transport, provider, peers } = setup();
    transport.emit({ type: 'open' });
    transport.emit({ type: 'message', data: peerAwareness(3, 1) });
    transport.sent.length = 0;
    replica.applyResult = typed;

    transport.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of(2)) });
    expect(provider.synced).toBe(true);
    expect(kinds(transport.sent)).toEqual(['query-awareness']);
    expect(peers.at(-1)?.[0]).toMatchObject({ clientId: 8, inferredCursor: null });

    transport.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(3)) });
    expect(peers.at(-1)?.[0]).toMatchObject({ clientId: 8, inferredCursor: typed });
    expect(replica.applied).toEqual([Uint8Array.of(2), Uint8Array.of(3)]);
  });

  it('applies awareness and typing inference in frame order', () => {
    const { replica, transport, peers } = setup();
    transport.emit({ type: 'open' });

    replica.applyResult = typed;
    transport.emit({
      type: 'message',
      data: concat(peerAwareness(3, 1), encodeUpdate(Uint8Array.of(4))),
    });
    expect(peers.at(-1)?.[0]).toMatchObject({ clientId: 8, inferredCursor: typed });

    replica.applyResult = { ...typed, endOffset: 9 };
    transport.emit({
      type: 'message',
      data: concat(encodeUpdate(Uint8Array.of(5)), peerAwareness(4, 2)),
    });
    expect(peers.at(-1)?.[0]).toMatchObject({
      clientId: 8,
      inferredCursor: null,
      cursor: { story: 'body', anchor: Uint8Array.of(2), head: Uint8Array.of(2) },
    });
  });

  it('clears peers on close', () => {
    const { transport, peers } = setup();
    transport.emit({ type: 'open' });
    transport.emit({ type: 'message', data: peerAwareness(3, 1) });
    expect(peers.at(-1)).toHaveLength(1);
    transport.emit({ type: 'close' });
    expect(peers.at(-1)).toEqual([]);
  });
});
