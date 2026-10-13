import type { createWorkerEditSession } from '../commands/session';
import type { CommandEnvelope } from '../commands/types';

/** @internal A buffer reached by freezeDeep with its length at freeze time. */
export type FrozenBuffer = { buffer: ArrayBufferLike; byteLength: number };

/** @internal An envelope an adapter received, frozen on arrival, with the buffers it holds. */
export type FrozenEnvelope = { envelope: CommandEnvelope<unknown>; buffers: FrozenBuffer[] };

/** @internal Freezes every object and array under value and returns the buffers it reached. */
export function freezeDeep(value: unknown): FrozenBuffer[] {
  const buffers: FrozenBuffer[] = [];
  const seen = new Set<object>();
  const visit = (node: unknown): void => {
    if (node === null || typeof node !== 'object' || seen.has(node)) return;
    seen.add(node);
    if (node instanceof ArrayBuffer || ArrayBuffer.isView(node)) {
      const buffer = node instanceof ArrayBuffer ? node : node.buffer;
      buffers.push({ buffer, byteLength: buffer.byteLength });
      return;
    }
    for (const key of Reflect.ownKeys(node)) visit((node as Record<PropertyKey, unknown>)[key]);
    Object.freeze(node);
  };
  visit(value);
  return buffers;
}

/** @internal The buffers whose length changed since freezeDeep, which a transfer detaches to zero. */
export function detachedBuffers(buffers: readonly FrozenBuffer[]): FrozenBuffer[] {
  return buffers.filter((entry) => entry.buffer.byteLength !== entry.byteLength);
}

/** @internal Wraps a session factory so its connection's edit receives every envelope deep-frozen. */
export function freezingEdits(
  create: typeof createWorkerEditSession, received: FrozenEnvelope[]
): typeof createWorkerEditSession {
  return (options) => {
    const call = options.connection.client.call;
    const edit = call.edit;
    call.edit = (envelope) => {
      received.push({ envelope, buffers: freezeDeep(envelope) });
      return edit.call(call, envelope);
    };
    return create(options);
  };
}
