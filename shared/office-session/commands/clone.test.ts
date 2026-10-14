import { expect, test } from 'bun:test';
import { cloneData } from './clone';
import { createCommandHost } from './host';
import { createWorkerEditSession } from './session';
import type { CommandAck } from './types';

test('passes Blob and File through by reference without structuredClone', () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'structuredClone');
  Object.defineProperty(globalThis, 'structuredClone', { value: undefined, configurable: true });
  try {
    const blob = new Blob(['open'], { type: 'text/plain' });
    const file = new File([blob], 'open.txt', { lastModified: 123 });
    const cloned = cloneData({ blob, file, aliases: [blob, file] });
    expect(cloned.blob).toBe(blob);
    expect(cloned.file).toBe(file);
    expect(cloned.aliases).toEqual([blob, file]);
    expect(cloned.aliases[0]).toBe(blob);
    expect(cloned.aliases[1]).toBe(file);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'structuredClone', descriptor);
    else Reflect.deleteProperty(globalThis, 'structuredClone');
  }
});

test('copies class instance own enumerable data to plain objects without structuredClone', () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'structuredClone');
  Object.defineProperty(globalThis, 'structuredClone', { value: undefined, configurable: true });
  try {
    class Payload {
      data = { text: 'insert' };
      self: unknown = this;
      method() { return 'prototype'; }
    }
    const payload = new Payload();
    Object.defineProperty(payload, 'hidden', { value: 'private' });
    const cloned = cloneData({ payload, alias: payload });
    expect(Object.getPrototypeOf(cloned.payload)).toBe(Object.prototype);
    expect(cloned.payload).not.toBe(payload);
    expect(cloned.payload.data).toEqual(payload.data);
    expect(cloned.payload.data).not.toBe(payload.data);
    expect(cloned.payload.self).toBe(cloned.payload);
    expect(cloned.alias).toBe(cloned.payload);
    expect('hidden' in cloned.payload).toBe(false);
    expect('method' in cloned.payload).toBe(false);
    expect(() => cloneData({ callback() {} })).toThrow(TypeError);
    expect(() => cloneData({ value: Symbol('data') })).toThrow(TypeError);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'structuredClone', descriptor);
    else Reflect.deleteProperty(globalThis, 'structuredClone');
  }
});

test('clones plain command data and initializes both layers without structuredClone', async () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'structuredClone');
  Object.defineProperty(globalThis, 'structuredClone', { value: undefined, configurable: true });
  try {
    const bytes = new Uint8Array([1, 2]);
    const source = {
      bytes, buffer: bytes.buffer, date: new Date(123),
      map: new Map<unknown, unknown>([[bytes, new Set([bytes])]]),
      array: [undefined, null, true, 'text', 4, BigInt(5)],
      sparse: new Array<unknown>(3),
    };
    const cyclic: { source: typeof source; self?: unknown } = { source };
    cyclic.self = cyclic;
    const cloned = cloneData(cyclic);
    expect(cloned).not.toBe(cyclic);
    expect(cloned.self).toBe(cloned);
    expect(cloned.source.bytes).not.toBe(bytes);
    expect(cloned.source.bytes.buffer).toBe(cloned.source.buffer);
    expect(cloned.source.date.getTime()).toBe(123);
    expect(cloned.source.map.get(cloned.source.bytes)).toEqual(new Set([cloned.source.bytes]));
    expect(cloned.source.array).toEqual(source.array);
    expect(cloned.source.sparse.length).toBe(3);
    expect(0 in cloned.source.sparse).toBe(false);
    bytes[0] = 9;
    expect(cloned.source.bytes[0]).toBe(1);
    let version = 0;
    const host = createCommandHost<string, typeof source, typeof source, string>({
      session: 'clone', incarnation: 0, epoch: 1, getVersion: () => version,
      apply() { version += 1; return { status: 'applied', selection: source, editState: source }; },
      frame: ({ coversThroughSeq }) => ({ coversThroughSeq, revision: version, payload: 'frame' }),
      snapshot: () => new Uint8Array(),
    });
    const initial: CommandAck<typeof source, typeof source, string> = {
      session: 'clone', incarnation: 0, epoch: 1, version: 0, throughSeq: 0, outcomes: [],
    };
    const queued: Array<() => void> = [];
    const session = createWorkerEditSession<typeof source, string, typeof source, typeof source, string, never>({
      source, connection: { initial, client: {
        call: {
          async edit(envelope) {
            const reply = await host.handler(null, envelope);
            return 'value' in reply ? reply.value : reply;
          },
          checkpoint: () => Promise.resolve({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
        },
        failure: undefined, onFailure: () => () => {}, dispose: () => Promise.resolve(),
      } },
      recovery: { reopen: () => Promise.reject(new Error('unexpected recovery')) },
      requestFullFrame: () => Promise.resolve(initial), enqueue: (callback) => { queued.push(callback); },
      now: () => 0,
    });
    const ticket = session.submit('insert', 'a');
    while (queued.length) queued.shift()!();
    expect(await ticket.ack).toMatchObject({ status: 'applied' });
    expect(await ticket.presented).toBe('presented');
    expect(session.state.selection?.bytes).toEqual(bytes);
    expect(session.state.selection?.bytes).not.toBe(bytes);
    await session.dispose();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'structuredClone', descriptor);
    else Reflect.deleteProperty(globalThis, 'structuredClone');
  }
});

test('clones all buffer views with shared identity, offsets, and lengths without structuredClone', () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'structuredClone');
  Object.defineProperty(globalThis, 'structuredClone', { value: undefined, configurable: true });
  try {
    const buffer = new ArrayBuffer(64);
    const views = [
      new Int8Array(buffer, 8, 4), new Uint8Array(buffer, 8, 4),
      new Uint8ClampedArray(buffer, 8, 4), new Int16Array(buffer, 8, 4),
      new Uint16Array(buffer, 8, 4), new Int32Array(buffer, 8, 4),
      new Uint32Array(buffer, 8, 4), new Float32Array(buffer, 8, 4),
      new Float64Array(buffer, 8, 4), new BigInt64Array(buffer, 8, 4),
      new BigUint64Array(buffer, 8, 4), new DataView(buffer, 7, 13),
      new Int32Array(buffer, 12, 0), new DataView(buffer, 5, 0),
    ];
    const bytes = new Uint8Array(buffer);
    bytes.forEach((_value, index) => { bytes[index] = index; });
    const cloned = cloneData({ views, buffer, alias: views[3], map: new Map([[views[3], buffer]]) });
    expect(cloned.buffer).not.toBe(buffer);
    expect(cloned.alias).toBe(cloned.views[3]);
    expect(cloned.map.get(cloned.alias)).toBe(cloned.buffer);
    for (const [index, view] of cloned.views.entries()) {
      const original = views[index]!;
      expect(view).not.toBe(original);
      expect(view.constructor).toBe(original.constructor);
      expect(view.buffer).toBe(cloned.buffer);
      expect(view.byteOffset).toBe(original.byteOffset);
      expect(view.byteLength).toBe(original.byteLength);
      if (!(view instanceof DataView) && !(original instanceof DataView)) {
        expect(view.length).toBe(original.length);
      }
    }
    expect([...new Uint8Array(cloned.buffer)]).toEqual([...bytes]);
    const integers = cloned.views[5] as Int32Array;
    const data = cloned.views[11] as DataView;
    integers[0] = 0;
    expect(data.getInt32(1)).toBe(0);
    expect(new DataView(buffer).getInt32(8)).not.toBe(0);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'structuredClone', descriptor);
    else Reflect.deleteProperty(globalThis, 'structuredClone');
  }
});

test('preserves shared-buffer view identity without structuredClone', () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'structuredClone');
  Object.defineProperty(globalThis, 'structuredClone', { value: undefined, configurable: true });
  try {
    const buffer = new SharedArrayBuffer(32);
    const integers = new Int32Array(buffer, 8, 2);
    const data = new DataView(buffer, 7, 10);
    integers[0] = 123;
    const cloned = cloneData({ integers, data, buffer });
    expect(cloned.integers).not.toBe(integers);
    expect(cloned.integers.buffer).toBe(cloned.buffer);
    expect(cloned.data.buffer).toBe(cloned.buffer);
    expect(cloned.integers.byteOffset).toBe(8);
    expect(cloned.integers.length).toBe(2);
    expect(cloned.data.byteOffset).toBe(7);
    expect(cloned.data.byteLength).toBe(10);
    expect(cloned.integers[0]).toBe(123);
    integers[0] = 456;
    expect(cloned.integers[0]).toBe(456);
    cloned.integers[1] = 789;
    expect(integers[1]).toBe(789);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'structuredClone', descriptor);
    else Reflect.deleteProperty(globalThis, 'structuredClone');
  }
});

test('clones regexps and boxed primitives without structuredClone', () => {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'structuredClone');
  Object.defineProperty(globalThis, 'structuredClone', { value: undefined, configurable: true });
  try {
    const regexp = /command/gi;
    regexp.lastIndex = 3;
    const values = [new Number(NaN), new String('text'), new Boolean(false), Object(BigInt(7))];
    const cloned = cloneData({ regexp, alias: regexp, values, valueAlias: values[3] });
    expect(cloned.regexp).not.toBe(regexp);
    expect(cloned.regexp).toBe(cloned.alias);
    expect(cloned.regexp.source).toBe('command');
    expect(cloned.regexp.flags).toBe('gi');
    expect(cloned.regexp.lastIndex).toBe(0);
    for (let index = 0; index < values.length; index += 1) {
      expect(cloned.values[index]).not.toBe(values[index]);
      expect(cloned.values[index]!.valueOf()).toEqual(values[index]!.valueOf());
    }
    expect(cloned.valueAlias).toBe(cloned.values[3]);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'structuredClone', descriptor);
    else Reflect.deleteProperty(globalThis, 'structuredClone');
  }
});
