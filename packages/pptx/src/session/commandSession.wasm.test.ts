import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { resolve } from 'node:path';
import { sessionWasmFactory, type WasmTestWorker } from '../../../../shared/office-session/testing/wasmWorker';
import type { ClientMessage } from '../../../../shared/office-session/protocol';
import type { PresentationCommandOpenOptions } from './commandProtocol';

type Client = typeof import('./commandSession');
let factory: Awaited<ReturnType<typeof sessionWasmFactory<Client>>>;

beforeAll(async () => {
  factory = await sessionWasmFactory<Client>(resolve(import.meta.dir, 'commandSession.ts'),
    resolve(import.meta.dir, 'commandSessionWasmTestWorker.ts'));
});

function input(worker: WasmTestWorker) {
  const open = worker.requests.find((message) => message.kind === 'call' && message.method === 'openCommandEditor');
  if (!open || open.kind !== 'call') throw new Error('Missing open request');
  return open.args[1] as PresentationCommandOpenOptions;
}

function compileRequests(worker: WasmTestWorker) {
  return worker.requests.filter((message) => message.kind === 'wasm-compile');
}

describe('peerless session wasm reuse', () => {
  test('compiles before payload copying and open, then reuses the advertised module after disposal', async () => {
    const h = factory();
    const bytes = new Uint8Array([1, 2, 3]);
    const first = h.client.createPeerlessSession(bytes, { fonts: [{ family: 'test',
      get bytes() { h.trace.push('font-access'); return new Uint8Array([4, 5, 6]); } }] });
    try {
      await first.ready;
      expect(compileRequests(h.workers[0]!)).toEqual([{ protocol: 1, kind: 'wasm-compile' }]);
      expect(h.trace.indexOf('post:wasm-compile')).toBeLessThan(h.trace.indexOf('font-access'));
      expect(h.trace.indexOf('post:wasm-compile')).toBeLessThan(h.trace.indexOf('post:openCommandEditor'));
      expect(h.trace.indexOf('fetch')).toBeLessThan(h.trace.indexOf('receive:openCommandEditor'));
      expect(h.compiles).toBe(1);
      expect(input(h.workers[0]!).wasm).toBeUndefined();
    } finally { await first.dispose(); }
    const fetches = h.trace.filter((event) => event === 'fetch').length;
    const second = h.client.createPeerlessSession(bytes);
    try {
      const client = await second.ready;
      expect(h.workers[0]!.terminated).toBe(true);
      expect(input(h.workers[1]!).wasm).toBe(h.workers[0]!.modules[0]);
      expect(compileRequests(h.workers[1]!)).toEqual([]);
      expect(h.workers[1]!.sources[0]).toBeInstanceOf(WebAssembly.Module);
      expect(h.workers[1]!.modules).toEqual([]);
      expect(h.compiles).toBe(1);
      expect(h.trace.filter((event) => event === 'fetch')).toHaveLength(fetches);
      expect(await client.save()).toEqual(bytes);
    } finally { await second.dispose(); }
  });

  for (const kind of ['module', 'bytes'] as const) test(`explicit wasm ${kind} overrides an empty and populated cache`, async () => {
    const h = factory();
    const wasm = kind === 'module' ? new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]))
      : new Uint8Array([7, 8, 9]).buffer;
    for (const warm of [false, true]) {
      if (warm) {
        const session = h.client.createPeerlessSession(new Uint8Array());
        await session.ready;
        await session.dispose();
      }
      const session = h.client.createPeerlessSession(new Uint8Array(), { wasm });
      try {
        await session.ready;
        const worker = h.workers[h.workers.length - 1]!;
        expect(compileRequests(worker)).toEqual([]);
        if (kind === 'module') expect(input(worker).wasm).toBe(wasm);
        else {
          expect(new Uint8Array(worker.sources[0] as ArrayBuffer)).toEqual(new Uint8Array(wasm as ArrayBuffer));
          expect((wasm as ArrayBuffer).byteLength).toBe(3);
        }
        expect(h.compiles).toBe(warm ? 1 : 0);
      } finally { await session.dispose(); }
    }
    const cached = h.client.createPeerlessSession(new Uint8Array());
    try { await cached.ready; expect(input(h.workers[h.workers.length - 1]!).wasm).toBe(h.workers[1]!.modules[0]); expect(h.compiles).toBe(1); }
    finally { await cached.dispose(); }
  });

  test('isolates custom workers from the default cache in both directions', async () => {
    const h = factory();
    for (const custom of [true, true, false, true, false]) {
      const session = h.client.createPeerlessSession(new Uint8Array(), custom ? { worker: h.worker } : {});
      try {
        await session.ready;
        const worker = h.workers[h.workers.length - 1]!;
        if (custom) { expect(compileRequests(worker)).toEqual([]); expect(input(worker).wasm).toBeUndefined(); }
      } finally { await session.dispose(); }
    }
    expect(input(h.workers[4]!).wasm).toBe(h.workers[2]!.modules[0]);
    expect(h.compiles).toBe(4);
  });

  for (const failure of ['compile', 'initialize'] as const) test(`retries after ${failure} fails`, async () => {
    const h = factory();
    if (failure === 'compile') h.rejectCompile(new Error('Compile failed'));
    else h.rejectInitialize(new Error('Initialize failed'));
    const failed = h.client.createPeerlessSession(new Uint8Array());
    await expect(failed.ready).rejects.toThrow();
    await failed.dispose();
    expect(h.workers[0]!.modules).toEqual([]);
    const retry = h.client.createPeerlessSession(new Uint8Array());
    try {
      await retry.ready;
      expect(input(h.workers[1]!).wasm).toBeUndefined();
      expect(compileRequests(h.workers[1]!)).toHaveLength(1);
      expect(h.compiles).toBe(2);
    } finally { await retry.dispose(); }
    const cached = h.client.createPeerlessSession(new Uint8Array());
    try { await cached.ready; expect(input(h.workers[2]!).wasm).toBe(h.workers[1]!.modules[0]); expect(h.compiles).toBe(2); }
    finally { await cached.dispose(); }
  });

  test('keys cached modules by asset URL', async () => {
    const h = factory();
    for (const url of ['https://example.test/first.wasm', 'https://example.test/second.wasm', 'https://example.test/first.wasm']) {
      h.assetUrl = url;
      const session = h.client.createPeerlessSession(new Uint8Array());
      await session.ready;
      await session.dispose();
    }
    expect(input(h.workers[2]!).wasm).toBe(h.workers[0]!.modules[0]);
    expect(h.compiles).toBe(2);
  });

  test('transfers copies of document, font views, fallback fonts, updates and explicit wasm', async () => {
    const h = factory();
    const document = new Uint8Array([0, 1, 2, 3, 4]);
    const fonts = new Uint8Array([5, 6, 7, 8, 9]);
    const fallback = new Uint8Array([10, 11, 12]).buffer;
    const update = new Uint8Array([13, 14, 15]);
    const wasm = new Uint8Array([16, 17, 18]).buffer;
    const worker = h.worker();
    const messages: ClientMessage[] = [];
    const post = worker.postMessage.bind(worker);
    const capture = spyOn(worker, 'postMessage').mockImplementation((message, transfer) => {
      messages.push(structuredClone(message));
      if (Array.isArray(transfer)) post(message, transfer);
      else post(message, transfer);
    });
    const session = h.client.createPeerlessSession(document.subarray(1, 4), { worker: () => worker,
      fonts: [{ family: 'test', bytes: fonts.subarray(1, 4) }],
      fallbackFonts: [{ family: 'fallback', bytes: new Uint8Array(fallback) }], initialUpdate: update, wasm });
    try {
      const client = await session.ready;
      const open = messages.find((message) => message.kind === 'call' && message.method === 'openCommandEditor');
      if (!open || open.kind !== 'call') throw new Error('Missing open request');
      const sent = open.args[1] as PresentationCommandOpenOptions;
      expect(new Uint8Array(sent.fonts![0]!.bytes)).toEqual(new Uint8Array([6, 7, 8]));
      expect(new Uint8Array(sent.fallbackFonts![0]!.bytes)).toEqual(new Uint8Array([10, 11, 12]));
      expect(sent.initialUpdate).toEqual(update);
      expect(await client.save()).toEqual(new Uint8Array([1, 2, 3]));
      expect(document).toEqual(new Uint8Array([0, 1, 2, 3, 4]));
      expect(fonts).toEqual(new Uint8Array([5, 6, 7, 8, 9]));
      expect(new Uint8Array(fallback)).toEqual(new Uint8Array([10, 11, 12]));
      expect(update).toEqual(new Uint8Array([13, 14, 15]));
      expect(new Uint8Array(wasm)).toEqual(new Uint8Array([16, 17, 18]));
    } finally { capture.mockRestore(); await session.dispose(); }
  });
});
