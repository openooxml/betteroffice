import { expect, test } from 'bun:test';
import { CollaborationProvider } from './provider';
import { decodeMessages, encodeSyncStep1, encodeSyncStep2, encodeUpdate } from './protocol';
import type { AsyncCollaborationReplica, CollaborationError, CollaborationTransport, CollaborationTransportEvent } from './types';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function setup(maxPendingBytes?: number) {
  const listeners = new Set<(event: CollaborationTransportEvent) => void>();
  const sent: Uint8Array[] = [];
  const errors: CollaborationError[] = [];
  const applied: number[] = [];
  const gate = deferred<void>();
  const replica: AsyncCollaborationReplica = {
    clientId: 1,
    encodeStateVector: async () => Uint8Array.of(1),
    encodeStateAsUpdate: async () => Uint8Array.of(2),
    async applyUpdate(update) { applied.push(update[0]!); await gate.promise; },
    onUpdate: () => () => {},
  };
  const transport: CollaborationTransport = { connect() {}, disconnect() {},
    send(bytes) { sent.push(bytes); return true; },
    onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); } } };
  const provider = new CollaborationProvider(replica, transport, { maxPendingBytes });
  provider.onError((error) => { errors.push(error); });
  const emit = (event: CollaborationTransportEvent) => { for (const listener of [...listeners]) listener(event); };
  provider.connect();
  emit({ type: 'open' });
  return { provider, applied, sent, emit, errors, gate };
}

async function settle() { for (let index = 0; index < 15; index++) await Promise.resolve(); }

test('async_handshake_and_ordered_remote_updates', async () => {
  const value = setup();
  try {
    await settle();
    expect(value.sent.flatMap((bytes) => decodeMessages(bytes)).some((message) => message.type === 'sync-step-1')).toBe(true);
    value.emit({ type: 'message', data: encodeSyncStep1(Uint8Array.of(1)) });
    await settle();
    expect(value.sent.flatMap((bytes) => decodeMessages(bytes)).some((message) => message.type === 'sync-step-2')).toBe(true);
    value.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of(3)) });
    value.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(4)) });
    await settle();
    expect(value.applied).toEqual([3]);
    expect(value.provider.synced).toBe(false);
    value.gate.resolve();
    await settle();
    expect(value.applied).toEqual([3, 4]);
    expect(value.provider.synced).toBe(true);
    expect(value.errors).toEqual([]);
  } finally { value.gate.resolve(); value.provider.destroy(); }
});

test('async_retired_completion_cannot_sync_reopened_connection', async () => {
  const value = setup();
  try {
    await settle();
    value.emit({ type: 'message', data: encodeSyncStep2(Uint8Array.of(3)) });
    value.emit({ type: 'close' });
    value.emit({ type: 'open' });
    value.gate.resolve();
    await settle();
    expect(value.provider.synced).toBe(false);
  } finally { value.provider.destroy(); }
});

test('async_inbound_backpressure_is_bounded', async () => {
  const value = setup(6);
  try {
    await settle();
    value.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(3)) });
    value.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(4)) });
    value.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(5)) });
    await settle();
    expect(value.errors.some((error) => error.code === 'backpressure')).toBe(true);
    expect(value.provider.status).toBe('disconnected');
    expect(value.applied).toEqual([3]);
  } finally { value.gate.resolve(); value.provider.destroy(); }
});
