import { describe, expect, it } from 'bun:test';
import {
  decodeAwarenessUpdate,
  decodeMessages,
  encodeAwarenessMessage,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
} from './protocol';
import { CollaborationProvider } from './provider';
import type {
  CollaborationError,
  CollaborationPeer,
  CollaborationReplica,
  CollaborationStatusChange,
  CollaborationTransport,
  CollaborationTransportEvent,
  CollaborationUpdateOrigin,
} from './types';

function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((length, part) => length + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.byteLength;
  }
  return output;
}

class FakeReplica implements CollaborationReplica {
  readonly clientId = 1;
  stateVector = Uint8Array.of(7);
  stateUpdate = Uint8Array.of(8);
  applied: Uint8Array[] = [];
  remoteStateVectors: Uint8Array[] = [];
  unsubscribeCount = 0;
  disposeCount = 0;
  applyError: unknown;
  applyResult: unknown;
  emitRemoteOnApply = false;
  private readonly listeners = new Set<
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();

  encodeStateVector(): Uint8Array {
    return this.stateVector;
  }

  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array {
    if (remoteStateVector) this.remoteStateVectors.push(remoteStateVector);
    return this.stateUpdate;
  }

  applyUpdate(update: Uint8Array): unknown {
    if (this.applyError) throw this.applyError;
    this.applied.push(update);
    if (this.emitRemoteOnApply) this.emit(update, 'remote');
    return this.applyResult;
  }

  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void {
    this.listeners.add(listener);
    let subscribed = true;
    return () => {
      if (!subscribed) return;
      subscribed = false;
      this.listeners.delete(listener);
      this.unsubscribeCount += 1;
    };
  }

  emit(update: Uint8Array, origin: CollaborationUpdateOrigin): void {
    for (const listener of [...this.listeners]) listener(update, origin);
  }

  dispose(): void {
    this.disposeCount += 1;
  }
}

class StatefulReplica implements CollaborationReplica {
  readonly clientId = 2;
  readonly state = new Set<number>();
  private readonly listeners = new Set<
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();

  encodeStateVector(): Uint8Array {
    return Uint8Array.from([...this.state].sort((left, right) => left - right));
  }

  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array {
    const remote = new Set(remoteStateVector ?? []);
    return Uint8Array.from([...this.state].filter((value) => !remote.has(value)));
  }

  applyUpdate(update: Uint8Array): void {
    for (const value of update) this.state.add(value);
    this.emit(update, 'remote');
  }

  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  edit(value: number): void {
    this.state.add(value);
    this.emit(Uint8Array.of(value), 'local');
  }

  private emit(update: Uint8Array, origin: CollaborationUpdateOrigin): void {
    for (const listener of [...this.listeners]) listener(update, origin);
  }
}

class FakeTransport implements CollaborationTransport {
  connectCount = 0;
  disconnectCount = 0;
  accept: unknown = true;
  mutateRejectedFrame = false;
  connectError: unknown;
  disconnectError: unknown;
  unsubscribeError: unknown;
  sendError: unknown;
  sent: Uint8Array[] = [];
  attempted: Uint8Array[] = [];
  readonly listenerHistory: Array<(event: CollaborationTransportEvent) => void> = [];
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();

  connect(): void {
    this.connectCount += 1;
    if (this.connectError) throw this.connectError;
  }

  disconnect(): void {
    this.disconnectCount += 1;
    if (this.disconnectError) throw this.disconnectError;
  }

  send(data: Uint8Array): boolean {
    if (this.sendError) throw this.sendError;
    this.attempted.push(data);
    if (this.accept === false) {
      if (this.mutateRejectedFrame) data.fill(0xff);
      return false;
    }
    if (typeof this.accept !== 'boolean') return this.accept as boolean;
    this.sent.push(data);
    return true;
  }

  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    this.listenerHistory.push(listener);
    return () => {
      this.listeners.delete(listener);
      if (this.unsubscribeError) throw this.unsubscribeError;
    };
  }

  emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }
}

function open(replica = new FakeReplica(), transport = new FakeTransport()) {
  const provider = new CollaborationProvider(replica, transport);
  provider.connect();
  transport.emit({ type: 'open' });
  return { provider, replica, transport };
}

describe('CollaborationProvider sync', () => {
  it('starts unsynced and sends an exact Step1 frame on open', () => {
    const replica = new FakeReplica();
    replica.stateVector = Uint8Array.of(1, 2);
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);

    expect(provider.status).toBe('disconnected');
    expect(provider.synced).toBe(false);
    provider.connect();
    expect(provider.status).toBe('connecting');
    transport.emit({ type: 'open' });

    expect(provider.status).toBe('connected');
    expect(provider.synced).toBe(false);
    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 0, 2, 1, 2]]);
  });

  it('answers Step1 with a diff encoded as Step2', () => {
    const { provider, replica, transport } = open();
    transport.sent = [];
    replica.stateUpdate = Uint8Array.of(5, 6);

    transport.emit({ type: 'message', data: encodeSyncStep1(Uint8Array.of(9, 10)) });

    expect(replica.remoteStateVectors).toEqual([Uint8Array.of(9, 10)]);
    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 1, 2, 5, 6]]);
    expect(provider.synced).toBe(false);
  });

  it('applies Step2 and updates sync state', () => {
    const { provider, replica, transport } = open();
    transport.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of(4)) });
    expect(replica.applied).toEqual([Uint8Array.of(4)]);
    expect(provider.synced).toBe(true);
  });

  it('forwards only local observer updates and never echoes remote application', () => {
    const { replica, transport } = open();
    transport.sent = [];
    replica.emitRemoteOnApply = true;

    replica.emit(Uint8Array.of(1), 'local');
    replica.emit(Uint8Array.of(2), 'remote');
    transport.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(3)) });

    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 2, 1, 1]]);
    expect(replica.applied).toEqual([Uint8Array.of(3)]);
  });

  it('allows an explicit reconnect after a physical close', () => {
    const { provider, transport } = open();
    transport.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of()) });
    expect(provider.synced).toBe(true);

    transport.emit({ type: 'close' });
    expect(provider.status).toBe('disconnected');
    expect(provider.synced).toBe(false);
    provider.connect();
    expect(provider.status).toBe('connecting');
    expect(transport.connectCount).toBe(2);
    transport.emit({ type: 'open' });

    expect(provider.status).toBe('connected');
    expect(provider.synced).toBe(false);
    expect(transport.sent.map((frame) => [...frame])).toEqual([
      [0, 0, 1, 7],
      [0, 0, 1, 7],
    ]);
  });

  it('also accepts a transport-managed reopen on the existing subscription', () => {
    const { provider, transport } = open();
    transport.emit({ type: 'close' });
    transport.emit({ type: 'open' });

    expect(provider.status).toBe('connected');
    expect(transport.connectCount).toBe(1);
    expect(transport.sent.map((frame) => [...frame])).toEqual([
      [0, 0, 1, 7],
      [0, 0, 1, 7],
    ]);
  });

  it('delegates duplicate and out-of-order updates unchanged', () => {
    const { replica, transport } = open();
    for (const value of [2, 1, 2]) {
      transport.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(value)) });
    }
    expect(replica.applied).toEqual([
      Uint8Array.of(2),
      Uint8Array.of(1),
      Uint8Array.of(2),
    ]);
  });

  it('handles concatenated sync, awareness, and query-awareness messages', () => {
    const { provider, replica, transport } = open();
    transport.sent = [];
    const frame = concat(
      encodeUpdate(Uint8Array.of(3)),
      Uint8Array.of(1, 1, 0),
      Uint8Array.of(3),
      encodeSyncStep2(Uint8Array.of(4))
    );

    transport.emit({ type: 'message', data: frame });

    expect(replica.applied).toEqual([Uint8Array.of(3), Uint8Array.of(4)]);
    expect(transport.sent.map((message) => [...message])).toEqual([[1, 1, 0]]);
    expect(provider.synced).toBe(true);
  });
});

describe('CollaborationProvider awareness', () => {
  it('announces identity, queries peers, and sends an explicit leave', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, {
      user: { name: 'Calm Otter' },
    });

    provider.connect();
    transport.emit({ type: 'open' });

    expect(decodeMessages(transport.sent[0])).toEqual([
      { type: 'sync-step-1', stateVector: Uint8Array.of(7) },
    ]);
    expect(decodeMessages(transport.sent[1])).toEqual([{ type: 'query-awareness' }]);
    const announced = decodeMessages(transport.sent[2])[0];
    expect(announced.type).toBe('awareness');
    if (announced.type !== 'awareness') return;
    expect(decodeAwarenessUpdate(announced.update)).toMatchObject([
      {
        clientId: 1,
        clock: 1,
        state: {
          user: { name: 'Calm Otter', color: '#B3261E' },
          cursor: null,
        },
      },
    ]);

    transport.emit({ type: 'message', data: encodeQueryAwareness() });
    const queryResponse = decodeMessages(transport.sent.at(-1) ?? Uint8Array.of())[0];
    expect(queryResponse.type).toBe('awareness');
    if (queryResponse.type !== 'awareness') return;
    expect(decodeAwarenessUpdate(queryResponse.update)).toMatchObject([
      {
        clientId: 1,
        clock: 2,
        state: { user: { name: 'Calm Otter', color: '#B3261E' } },
      },
    ]);

    provider.destroy();
    const left = decodeMessages(transport.sent.at(-1) ?? Uint8Array.of())[0];
    expect(left.type).toBe('awareness');
    if (left.type !== 'awareness') return;
    expect(decodeAwarenessUpdate(left.update)).toEqual([
      { clientId: 1, clock: 3, state: null },
    ]);
  });

  it('publishes remote peers and applies typing inference immediately', () => {
    const { provider, replica, transport } = open();
    const snapshots: CollaborationPeer[][] = [];
    provider.onPeers((peers) => snapshots.push(peers.slice()));
    transport.emit({
      type: 'message',
      data: encodeAwarenessMessage([
        {
          clientId: 8,
          clock: 3,
          state: {
            user: { name: 'Bright Fox', color: '#137333' },
            cursor: {
              story: 'body',
              anchor: Uint8Array.of(1),
              head: Uint8Array.of(1),
            },
          },
        },
      ]),
    });
    replica.applyResult = {
      clientId: 8,
      story: 'body',
      paraId: 'p1',
      endOffset: 6,
    };
    transport.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(9)) });

    expect(snapshots.at(-1)?.[0]).toMatchObject({
      clientId: 8,
      inferredCursor: {
        clientId: 8,
        story: 'body',
        paraId: 'p1',
        endOffset: 6,
      },
    });
  });

  it('discards malformed awareness without interrupting document sync', () => {
    const { provider, replica, transport } = open();
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    transport.emit({
      type: 'message',
      data: encodeSyncStep2(Uint8Array.of(4)),
    });

    transport.emit({
      type: 'message',
      data: concat(Uint8Array.of(1, 5, 1, 8, 1, 1, 123), encodeUpdate(Uint8Array.of(9))),
    });

    expect(provider.status).toBe('connected');
    expect(provider.synced).toBe(true);
    expect(transport.disconnectCount).toBe(0);
    expect(replica.applied).toEqual([Uint8Array.of(4), Uint8Array.of(9)]);
    expect(errors.map((error) => error.code)).toEqual(['protocol']);
  });

  it('coalesces cursor movement and keeps typing updates local', async () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, {
      user: { name: 'Calm Otter' },
    });
    provider.connect();
    transport.emit({ type: 'open' });
    const initialFrames = transport.sent.length;

    provider.setCursor({
      story: 'body',
      anchor: Uint8Array.of(1),
      head: Uint8Array.of(1),
    });
    provider.setCursor({
      story: 'body',
      anchor: Uint8Array.of(2),
      head: Uint8Array.of(2),
    });
    await Bun.sleep(100);

    expect(transport.sent).toHaveLength(initialFrames + 1);
    const movement = decodeMessages(transport.sent.at(-1) ?? Uint8Array.of())[0];
    expect(movement.type).toBe('awareness');
    if (movement.type !== 'awareness') return;
    expect(decodeAwarenessUpdate(movement.update)[0]?.state?.cursor).toEqual({
      story: 'body',
      anchor: Uint8Array.of(2),
      head: Uint8Array.of(2),
    });

    provider.setCursor({
      story: 'body',
      anchor: Uint8Array.of(3),
      head: Uint8Array.of(3),
    });
    provider.setCursor(
      {
        story: 'body',
        anchor: Uint8Array.of(4),
        head: Uint8Array.of(4),
      },
      false
    );
    await Bun.sleep(100);
    expect(transport.sent).toHaveLength(initialFrames + 1);
    provider.destroy();
  });
});

describe('CollaborationProvider transport flow control', () => {
  it('queues on backpressure and drains frames in order', () => {
    const { provider, replica, transport } = open();
    transport.sent = [];
    transport.accept = false;
    replica.emit(Uint8Array.of(1), 'local');
    replica.emit(Uint8Array.of(2), 'local');

    expect(provider.pendingBytes).toBe(8);
    expect(transport.sent).toEqual([]);
    transport.accept = true;
    transport.emit({ type: 'drain' });

    expect(provider.pendingBytes).toBe(0);
    expect(transport.sent.map((frame) => [...frame])).toEqual([
      [0, 2, 1, 1],
      [0, 2, 1, 2],
    ]);
  });

  it('keeps owned queue bytes when a rejecting transport mutates its argument', () => {
    const { provider, replica, transport } = open();
    transport.sent = [];
    transport.accept = false;
    transport.mutateRejectedFrame = true;
    replica.emit(Uint8Array.of(5), 'local');

    transport.accept = true;
    transport.emit({ type: 'drain' });
    expect(provider.pendingBytes).toBe(0);
    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 2, 1, 5]]);
  });

  it('attempts a leave directly while document frames are backpressured', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, {
      user: { name: 'Calm Otter' },
    });
    provider.connect();
    transport.emit({ type: 'open' });
    transport.attempted = [];
    transport.sent = [];
    transport.accept = false;

    replica.emit(Uint8Array.of(5), 'local');
    expect(provider.pendingBytes).toBe(4);
    provider.destroy();

    expect(provider.pendingBytes).toBe(0);
    expect(transport.attempted).toHaveLength(2);
    const [leave] = decodeMessages(transport.attempted[1]);
    expect(leave.type).toBe('awareness');
    if (leave.type !== 'awareness') return;
    expect(decodeAwarenessUpdate(leave.update)).toEqual([{ clientId: 1, clock: 2, state: null }]);
  });

  it('fails on overflow and converges through a fresh state-vector handshake', () => {
    const replica = new StatefulReplica();
    const server = new StatefulReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, { maxPendingBytes: 4 });
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    provider.connect();
    transport.emit({ type: 'open' });
    transport.sent = [];
    transport.accept = false;

    replica.edit(1);
    replica.edit(2);
    expect(provider.pendingBytes).toBe(0);
    expect(provider.status).toBe('disconnected');
    expect(provider.synced).toBe(false);
    expect(transport.disconnectCount).toBe(1);
    expect(errors.map((error) => error.code)).toEqual(['backpressure']);

    transport.accept = true;
    provider.connect();
    transport.emit({ type: 'open' });
    transport.sent = [];
    transport.emit({ type: 'message', data: encodeSyncStep1(server.encodeStateVector()) });

    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 1, 2, 1, 2]]);
    const [message] = decodeMessages(transport.sent[0]);
    expect(message.type).toBe('sync-step-2');
    if (message.type === 'sync-step-2') server.applyUpdate(message.update);
    expect([...server.state]).toEqual([...replica.state]);
  });

  it('fails instead of dropping a frame when send returns a non-boolean', () => {
    const { provider, replica, transport } = open();
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    transport.accept = undefined;

    replica.emit(Uint8Array.of(1), 'local');

    expect(provider.status).toBe('disconnected');
    expect(provider.synced).toBe(false);
    expect(transport.disconnectCount).toBe(1);
    expect(errors.map((error) => error.code)).toEqual(['transport']);
  });

  it('drops pending frames on close and resynchronizes on reopen', () => {
    const { provider, replica, transport } = open();
    transport.accept = false;
    replica.emit(Uint8Array.of(1), 'local');
    expect(provider.pendingBytes).toBe(4);

    transport.emit({ type: 'close' });
    expect(provider.pendingBytes).toBe(0);
    transport.accept = true;
    transport.sent = [];
    transport.emit({ type: 'open' });
    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 0, 1, 7]]);
  });
});

describe('CollaborationProvider errors and ownership', () => {
  it('fails closed on malformed, unknown, auth, and oversized frames', () => {
    const cases = [
      Uint8Array.of(0, 2, 2, 1),
      Uint8Array.of(9),
      Uint8Array.of(2, 0, 3, 110, 111, 112),
      new Uint8Array(7),
    ];

    for (const data of cases) {
      const replica = new FakeReplica();
      const transport = new FakeTransport();
      const provider = new CollaborationProvider(replica, transport, { maxFrameBytes: 6 });
      const errors: CollaborationError[] = [];
      provider.onError((error) => errors.push(error));
      provider.connect();
      transport.emit({ type: 'open' });
      transport.emit({ type: 'message', data });

      expect(provider.status).toBe('disconnected');
      expect(provider.synced).toBe(false);
      expect(transport.disconnectCount).toBe(1);
      expect(errors.map((error) => error.code)).toEqual(['protocol']);
      if (data[0] === 2) expect(errors[0].message).toContain('Authentication denied: nop');
    }
  });

  it('fails before processing a concatenated message flood', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, { maxMessagesPerFrame: 4 });
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    provider.connect();
    transport.emit({ type: 'open' });
    transport.sent = [];

    transport.emit({ type: 'message', data: new Uint8Array(5).fill(3) });

    expect(transport.sent).toEqual([]);
    expect(provider.status).toBe('disconnected');
    expect(errors[0].message).toContain('Frame exceeds 4 messages');
  });

  it('normalizes transport and replica failures and fails closed', () => {
    const send = open();
    const sendErrors: CollaborationError[] = [];
    send.provider.onError((error) => sendErrors.push(error));
    send.transport.sendError = new Error('blocked');
    send.replica.emit(Uint8Array.of(1), 'local');
    expect(sendErrors.map((error) => error.code)).toEqual(['transport']);
    expect(send.provider.status).toBe('disconnected');

    const transportEvent = open();
    const transportErrors: CollaborationError[] = [];
    transportEvent.provider.onError((error) => transportErrors.push(error));
    transportEvent.transport.emit({ type: 'error', error: 'socket failed' });
    expect(transportErrors.map((error) => error.code)).toEqual(['transport']);
    expect(transportEvent.provider.status).toBe('disconnected');

    const apply = open();
    const applyErrors: CollaborationError[] = [];
    apply.provider.onError((error) => applyErrors.push(error));
    apply.replica.applyError = new Error('bad update');
    apply.transport.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of(2)) });
    expect(applyErrors.map((error) => error.code)).toEqual(['replica']);
    expect(apply.provider.status).toBe('disconnected');
    expect(apply.provider.synced).toBe(false);
  });

  it('normalizes synchronous connect failures and returns to disconnected', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    transport.connectError = new Error('offline');
    const provider = new CollaborationProvider(replica, transport);
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));

    provider.connect();
    expect(provider.status).toBe('disconnected');
    expect(errors).toHaveLength(1);
    expect(errors[0].code).toBe('transport');
    expect(errors[0].cause).toBe(transport.connectError);
    expect(transport.disconnectCount).toBe(1);
  });

  it('copies incoming frame bytes before applying them', () => {
    const { replica, transport } = open();
    const frame = encodeUpdate(Uint8Array.of(4, 5));
    transport.emit({ type: 'message', data: frame });
    frame.fill(9);
    expect(replica.applied).toEqual([Uint8Array.of(4, 5)]);
  });

  it('enforces outbound frame limits', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, { maxFrameBytes: 4 });
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    provider.connect();
    transport.emit({ type: 'open' });
    replica.emit(Uint8Array.of(1, 2), 'local');
    expect(errors.map((error) => error.code)).toEqual(['protocol']);
  });
});

describe('CollaborationProvider lifecycle', () => {
  it('emits connection and initial-sync status changes', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);
    const changes: CollaborationStatusChange[] = [];
    provider.onStatus((change) => changes.push({ ...change }));

    provider.connect();
    transport.emit({ type: 'open' });
    transport.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of()) });
    transport.emit({ type: 'close', reason: 'network' });

    expect(changes).toEqual([
      { status: 'connecting', synced: false },
      { status: 'connected', synced: false },
      { status: 'connected', synced: true },
      { status: 'disconnected', synced: false },
    ]);
  });

  it('unsubscribes status and error listeners independently', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);
    const changes: CollaborationStatusChange[] = [];
    const errors: CollaborationError[] = [];
    const offStatus = provider.onStatus((change) => changes.push(change));
    const offError = provider.onError((error) => errors.push(error));
    offStatus();
    offError();

    provider.connect();
    transport.emit({ type: 'error', error: new Error('ignored') });
    expect(changes).toEqual([]);
    expect(errors).toEqual([]);
  });

  it('keeps duplicate status and error subscriptions independent', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);
    let statusCalls = 0;
    let errorCalls = 0;
    const statusListener = () => {
      statusCalls += 1;
    };
    const errorListener = () => {
      errorCalls += 1;
    };
    const offStatusFirst = provider.onStatus(statusListener);
    provider.onStatus(statusListener);
    const offErrorFirst = provider.onError(errorListener);
    provider.onError(errorListener);
    offStatusFirst();
    offErrorFirst();

    provider.connect();
    expect(statusCalls).toBe(1);
    transport.emit({ type: 'open' });
    transport.emit({ type: 'message', data: Uint8Array.of(9) });

    expect(errorCalls).toBe(1);
  });

  it('cannot regress from destroyed during a reentrant cleanup error', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);
    transport.unsubscribeError = new Error('unsubscribe failed');
    provider.onError(() => provider.destroy());
    provider.connect();

    provider.disconnect();

    expect(provider.status).toBe('destroyed');
    provider.connect();
    expect(transport.connectCount).toBe(1);
  });

  it('ignores stale callbacks from an earlier connection attempt', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);
    provider.connect();
    const staleListener = transport.listenerHistory[0];
    provider.disconnect();
    provider.connect();

    staleListener({ type: 'open' });
    staleListener({ type: 'message', data: encodeUpdate(Uint8Array.of(9)) });
    expect(transport.sent).toEqual([]);
    expect(replica.applied).toEqual([]);

    transport.emit({ type: 'open' });
    expect(transport.sent.map((frame) => [...frame])).toEqual([[0, 0, 1, 7]]);
  });

  it('makes connect, disconnect, and destroy idempotent without disposing the replica', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);

    provider.connect();
    provider.connect();
    expect(transport.connectCount).toBe(1);
    provider.disconnect();
    provider.disconnect();
    expect(transport.disconnectCount).toBe(1);
    provider.destroy();
    provider.destroy();
    provider.connect();

    expect(provider.status).toBe('destroyed');
    expect(provider.synced).toBe(false);
    expect(replica.unsubscribeCount).toBe(1);
    expect(replica.disposeCount).toBe(0);
    expect(transport.connectCount).toBe(1);
    expect(transport.disconnectCount).toBe(1);
  });
});
