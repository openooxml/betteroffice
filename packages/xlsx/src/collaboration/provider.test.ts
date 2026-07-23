import { describe, expect, it } from 'bun:test';
import {
  decodeAwarenessUpdate,
  decodeMessages,
  encodeAwarenessUpdate,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
} from './protocol';
import { CollaborationProvider } from './provider';
import type {
  CollaborationError,
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

function messageTypes(frames: readonly Uint8Array[]): string[] {
  return frames.flatMap((frame) => decodeMessages(frame).map((message) => message.type));
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
    return undefined;
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
  disconnectResult: Promise<void> | undefined;
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

  disconnect(): void | Promise<void> {
    this.disconnectCount += 1;
    if (this.disconnectError) throw this.disconnectError;
    return this.disconnectResult;
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
    expect([...transport.sent[0]]).toEqual([0, 0, 2, 1, 2]);
    expect(messageTypes(transport.sent)).toEqual([
      'sync-step-1',
      'query-awareness',
      'awareness',
    ]);
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
    expect(messageTypes(transport.sent)).toEqual([
      'sync-step-1',
      'query-awareness',
      'awareness',
      'sync-step-1',
      'query-awareness',
      'awareness',
    ]);
  });

  it('also accepts a transport-managed reopen on the existing subscription', () => {
    const { provider, transport } = open();
    transport.emit({ type: 'close' });
    transport.emit({ type: 'open' });

    expect(provider.status).toBe('connected');
    expect(transport.connectCount).toBe(1);
    expect(messageTypes(transport.sent)).toEqual([
      'sync-step-1',
      'query-awareness',
      'awareness',
      'sync-step-1',
      'query-awareness',
      'awareness',
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
    expect(messageTypes(transport.sent)).toEqual(['awareness']);
    const [response] = decodeMessages(transport.sent[0]);
    if (response.type !== 'awareness') throw new Error('expected awareness');
    expect(decodeAwarenessUpdate(response.update)).toEqual([
      {
        clientId: 1,
        clock: 2,
        state: {
          user: { name: 'Anonymous', color: '#B3261E' },
          cursor: null,
        },
      },
    ]);
    expect(provider.synced).toBe(true);
  });

  it('isolates malformed awareness while document updates keep flowing', () => {
    const { provider, replica, transport } = open();
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    transport.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of()) });
    replica.applied = [];

    transport.emit({
      type: 'message',
      data: concat(
        Uint8Array.of(1, 5, 1, 22, 1, 1, 123),
        encodeUpdate(Uint8Array.of(4))
      ),
    });
    transport.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(5)) });

    expect(provider.status).toBe('connected');
    expect(provider.synced).toBe(true);
    expect(transport.disconnectCount).toBe(0);
    expect(replica.applied).toEqual([Uint8Array.of(4), Uint8Array.of(5)]);
    expect(errors.map((error) => error.code)).toEqual(['protocol']);
    expect(errors[0].message).toContain('Invalid awareness update');
    provider.destroy();
  });

  it('publishes host identity and a cursor in the initial awareness state', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, {
      user: { name: 'Swift Fox', color: '#123abc' },
    });
    provider.setCursor({
      sheet: 'sheet:0',
      anchor: { row: 2, col: 3 },
      head: { row: 5, col: 7 },
    });
    provider.connect();
    transport.emit({ type: 'open' });

    const frame = transport.sent.find(
      (candidate) => decodeMessages(candidate)[0]?.type === 'awareness'
    );
    expect(frame).toBeDefined();
    const [message] = decodeMessages(frame!);
    if (message.type !== 'awareness') throw new Error('expected awareness');
    expect(decodeAwarenessUpdate(message.update)).toEqual([
      {
        clientId: 1,
        clock: 1,
        state: {
          user: { name: 'Swift Fox', color: '#123ABC' },
          cursor: {
            sheet: 'sheet:0',
            anchor: { row: 2, col: 3 },
            head: { row: 5, col: 7 },
          },
        },
      },
    ]);
    provider.destroy();
  });

  it('applies remote awareness clocks and clears peers on explicit leave', () => {
    const { provider, transport } = open();
    const snapshots: number[][] = [];
    provider.onAwareness((peers) => snapshots.push(peers.map((peer) => peer.clock)));
    const state = {
      user: { name: 'Quiet Otter', color: '#0B57D0' },
      cursor: {
        sheet: 'sheet:0',
        anchor: { row: 1, col: 2 },
        head: { row: 1, col: 2 },
      },
    };
    transport.emit({
      type: 'message',
      data: encodeAwarenessUpdate([{ clientId: 22, clock: 4, state }]),
    });
    transport.emit({
      type: 'message',
      data: encodeAwarenessUpdate([{ clientId: 22, clock: 3, state: null }]),
    });
    transport.emit({
      type: 'message',
      data: encodeAwarenessUpdate([{ clientId: 22, clock: 5, state: null }]),
    });

    expect(snapshots).toEqual([[], [4], []]);
    expect(provider.peers).toEqual([]);
    provider.destroy();
  });

  it('sends an explicit null awareness state before disconnecting', () => {
    const { provider, transport } = open();
    transport.sent = [];
    provider.disconnect();

    expect(messageTypes(transport.sent)).toEqual(['awareness']);
    const [message] = decodeMessages(transport.sent[0]);
    if (message.type !== 'awareness') throw new Error('expected awareness');
    expect(decodeAwarenessUpdate(message.update)).toEqual([
      { clientId: 1, clock: 2, state: null },
    ]);
  });

  it('coalesces rapid cursor changes and publishes the latest selection', async () => {
    const { provider, transport } = open();
    transport.sent = [];
    provider.setCursor({
      sheet: 'sheet:0',
      anchor: { row: 1, col: 1 },
      head: { row: 1, col: 1 },
    });
    provider.setCursor({
      sheet: 'sheet:0',
      anchor: { row: 4, col: 5 },
      head: { row: 8, col: 9 },
    });
    await Bun.sleep(100);

    expect(messageTypes(transport.sent)).toEqual(['awareness']);
    const [message] = decodeMessages(transport.sent[0]);
    if (message.type !== 'awareness') throw new Error('expected awareness');
    expect(decodeAwarenessUpdate(message.update)[0].state?.cursor).toEqual({
      sheet: 'sheet:0',
      anchor: { row: 4, col: 5 },
      head: { row: 8, col: 9 },
    });
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

  it('attempts an explicit leave directly under backpressure during teardown', () => {
    for (const method of ['disconnect', 'destroy'] as const) {
      const { provider, replica, transport } = open();
      transport.sent = [];
      transport.accept = false;
      replica.emit(Uint8Array.of(5), 'local');
      expect(provider.pendingBytes).toBeGreaterThan(0);
      transport.attempted = [];

      provider[method]();

      expect(provider.pendingBytes).toBe(0);
      expect(messageTypes(transport.attempted)).toEqual(['awareness']);
      const [message] = decodeMessages(transport.attempted[0]);
      if (message.type !== 'awareness') throw new Error('expected awareness');
      expect(decodeAwarenessUpdate(message.update)).toEqual([
        { clientId: 1, clock: 2, state: null },
      ]);
    }
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
    expect(messageTypes(transport.sent)).toEqual([
      'sync-step-1',
      'query-awareness',
      'awareness',
    ]);
  });
});

describe('CollaborationProvider errors and ownership', () => {
  it('fails closed on malformed, unknown, auth, and oversized frames', () => {
    const cases = [
      Uint8Array.of(0, 2, 2, 1),
      Uint8Array.of(9),
      Uint8Array.of(2, 0, 3, 110, 111, 112),
      new Uint8Array(129),
    ];

    for (const data of cases) {
      const replica = new FakeReplica();
      const transport = new FakeTransport();
      const provider = new CollaborationProvider(replica, transport, { maxFrameBytes: 128 });
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
    expect(messageTypes(transport.sent)).toEqual([
      'sync-step-1',
      'query-awareness',
      'awareness',
    ]);
  });

  it('waits for asynchronous teardown before reconnecting', async () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    let finishDisconnect = () => {};
    transport.disconnectResult = new Promise<void>((resolve) => {
      finishDisconnect = resolve;
    });
    const provider = new CollaborationProvider(replica, transport);
    provider.connect();
    transport.emit({ type: 'open' });

    provider.disconnect();
    provider.connect();
    expect(provider.status).toBe('connecting');
    expect(transport.connectCount).toBe(1);

    finishDisconnect();
    await transport.disconnectResult;
    await Promise.resolve();
    expect(transport.connectCount).toBe(2);
  });

  it('does not reconnect after asynchronous teardown fails', async () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    transport.disconnectResult = Promise.reject(new Error('still connected'));
    const provider = new CollaborationProvider(replica, transport);
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    provider.connect();
    transport.emit({ type: 'open' });

    provider.disconnect();
    provider.connect();
    await transport.disconnectResult.catch(() => {});
    await Promise.resolve();

    expect(provider.status).toBe('disconnected');
    expect(transport.connectCount).toBe(1);
    expect(errors[0].message).toContain('still connected');
  });

  it('does not deliver stale status after a reentrant transition', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport);
    const statuses: string[] = [];
    provider.onStatus(({ status }) => {
      if (status === 'connected') provider.disconnect();
    });
    provider.onStatus(({ status }) => statuses.push(status));

    provider.connect();
    transport.emit({ type: 'open' });

    expect(statuses).toEqual(['connecting', 'disconnected']);
  });

  it('rejects oversized inbound frames before copying them', () => {
    const replica = new FakeReplica();
    const transport = new FakeTransport();
    const provider = new CollaborationProvider(replica, transport, { maxFrameBytes: 4 });
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    provider.connect();
    transport.emit({ type: 'open' });
    const frame = new Uint8Array(5);
    Object.defineProperty(frame, 'slice', {
      value: () => {
        throw new Error('frame was copied');
      },
    });

    transport.emit({ type: 'message', data: frame });

    expect(errors).toHaveLength(1);
    expect(errors[0].message).toContain('Frame exceeds 4 bytes');
    expect(errors[0].message).not.toContain('frame was copied');
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
