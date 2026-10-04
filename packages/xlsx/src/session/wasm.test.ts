import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  createSourceFile, forEachChild, isIdentifier, isNewExpression, isObjectLiteralExpression,
  isPropertyAssignment, isStringLiteral, ScriptTarget, type Node,
} from 'typescript';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import {
  mainThreadWasmSpies, sessionWasmFactory, type WasmTestWorker,
} from '../../../../shared/office-session/testing/wasmWorker';

type Client = typeof import('./client');
let factory: Awaited<ReturnType<typeof sessionWasmFactory<Client>>>;

beforeAll(async () => {
  factory = await sessionWasmFactory<Client>(
    resolve(import.meta.dir, 'client.ts'), resolve(import.meta.dir, 'worker.ts')
  );
});

function input(worker: WasmTestWorker): { wasm?: ArrayBuffer | WebAssembly.Module } {
  const open = worker.requests.find((message) =>
    isClientMessage(message) && message.kind === 'call' && message.method === 'open'
  );
  if (!open || open.kind !== 'call') throw new Error('Missing open request');
  return open.args[1] as { wasm?: ArrayBuffer | WebAssembly.Module };
}

function compileRequests(worker: WasmTestWorker) {
  return worker.requests.filter((message) => message.kind === 'wasm-compile');
}

describe('workbook session wasm reuse', () => {
  test('keeps worker URLs and options statically evaluable', () => {
    const source = createSourceFile('client.ts',
      readFileSync(resolve(import.meta.dir, 'client.ts'), 'utf8'), ScriptTarget.Latest);
    const workers: Array<Record<string, string>> = [];
    function visit(node: Node): void {
      if (isNewExpression(node) && isIdentifier(node.expression) && node.expression.text === 'Worker') {
        expect(node.getText(source))
          .toBe("new Worker(new URL('./xlsxSessionWorker.mjs', import.meta.url), { type: 'module' })");
        expect(node.arguments?.[0]?.getText(source))
          .toBe("new URL('./xlsxSessionWorker.mjs', import.meta.url)");
        const options = node.arguments?.[1];
        if (!options || !isObjectLiteralExpression(options)) throw new Error('Non-literal worker options');
        workers.push(Object.fromEntries(options.properties.map((property) => {
          if (!isPropertyAssignment(property) || !isIdentifier(property.name) ||
            !isStringLiteral(property.initializer)) throw new Error('Non-literal worker option');
          return [property.name.text, property.initializer.text];
        })));
      }
      forEachChild(node, visit);
    }
    visit(source);
    expect(workers).toEqual([{ type: 'module' }]);
  });

  test('starts before open delivery and reuses the first worker module after disposal', async () => {
    const h = factory();
    const spies = mainThreadWasmSpies();
    const bytes = new Uint8Array([1, 2, 3]);
    try {
      const first = await h.client.openWorkbookSession(bytes);
      expect(compileRequests(h.workers[0])).toEqual([{ protocol: 1, kind: 'wasm-compile' }]);
      expect(h.trace.indexOf('post:wasm-compile')).toBeLessThan(h.trace.indexOf('post:open'));
      expect(h.trace.indexOf('fetch')).toBeLessThan(h.trace.indexOf('receive:open'));
      expect(h.compiles).toBe(1);
      expect(input(h.workers[0]).wasm).toBeUndefined();
      expect(await first.save()).toEqual(bytes);
      await first.dispose();

      const second = await h.client.openWorkbookSession(bytes);
      try {
        expect(h.workers[0].terminated).toBe(true);
        expect(input(h.workers[1]).wasm).toBe(h.workers[0].modules[0]);
        expect(compileRequests(h.workers[1])).toEqual([]);
        expect(h.workers[1].sources[0]).toBeInstanceOf(WebAssembly.Module);
        expect(h.workers[1].modules).toEqual([]);
        expect(h.compiles).toBe(1);
        expect(await second.save()).toEqual(bytes);
        expect(bytes).toEqual(new Uint8Array([1, 2, 3]));
        for (const spy of spies) expect(spy).not.toHaveBeenCalled();
      } finally { await second.dispose(); }
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('honours an explicit module before and after the default cache is populated', async () => {
    const h = factory();
    const explicit = new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
    const spies = mainThreadWasmSpies();
    try {
      const first = await h.client.openWorkbookSession(new Uint8Array(), { wasm: explicit });
      expect(input(h.workers[0]).wasm).toBe(explicit);
      expect(compileRequests(h.workers[0])).toEqual([]);
      expect(h.compiles).toBe(0);
      await first.dispose();

      const warm = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[1]).wasm).toBeUndefined();
      expect(h.compiles).toBe(1);
      await warm.dispose();

      const overridden = await h.client.openWorkbookSession(new Uint8Array(), { wasm: explicit });
      expect(input(h.workers[2]).wasm).toBe(explicit);
      expect(compileRequests(h.workers[2])).toEqual([]);
      expect(h.compiles).toBe(1);
      await overridden.dispose();

      const cached = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[3]).wasm).toBe(h.workers[1].modules[0]);
      expect(h.compiles).toBe(1);
      await cached.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('copies explicit bytes and bypasses the default cache', async () => {
    const h = factory();
    const wasm = new Uint8Array([1, 2, 3]).buffer;
    const spies = mainThreadWasmSpies();
    try {
      const first = await h.client.openWorkbookSession(new Uint8Array(), { wasm });
      expect(compileRequests(h.workers[0])).toEqual([]);
      expect(h.compiles).toBe(0);
      expect(new Uint8Array(h.workers[0].sources[0] as ArrayBuffer)).toEqual(new Uint8Array(wasm));
      expect(wasm.byteLength).toBe(3);
      await first.dispose();

      const second = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[1]).wasm).toBeUndefined();
      expect(h.compiles).toBe(1);
      await second.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('isolates custom worker factories from the default cache in both directions', async () => {
    const h = factory();
    const spies = mainThreadWasmSpies();
    try {
      const first = await h.client.openWorkbookSession(new Uint8Array(), { worker: h.worker });
      expect(compileRequests(h.workers[0])).toEqual([]);
      expect(h.trace.indexOf('post:open')).toBeLessThan(h.trace.indexOf('fetch'));
      expect(input(h.workers[0]).wasm).toBeUndefined();
      expect(h.compiles).toBe(1);
      await first.dispose();
      const second = await h.client.openWorkbookSession(new Uint8Array(), { worker: h.worker });
      expect(compileRequests(h.workers[1])).toEqual([]);
      expect(input(h.workers[1]).wasm).toBeUndefined();
      expect(h.compiles).toBe(2);
      await second.dispose();

      const warm = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[2]).wasm).toBeUndefined();
      expect(h.compiles).toBe(3);
      await warm.dispose();

      const custom = await h.client.openWorkbookSession(new Uint8Array(), { worker: h.worker });
      expect(compileRequests(h.workers[3])).toEqual([]);
      expect(input(h.workers[3]).wasm).toBeUndefined();
      expect(h.compiles).toBe(4);
      await custom.dispose();

      const cached = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[4]).wasm).toBe(h.workers[2].modules[0]);
      expect(h.compiles).toBe(4);
      await cached.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test.each(['module', 'bytes'])('honours explicit %s with a custom worker factory', async (kind) => {
    const h = factory();
    const bytes = new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]);
    const wasm = kind === 'module' ? new WebAssembly.Module(bytes) : bytes.buffer;
    const spies = mainThreadWasmSpies();
    try {
      const warm = await h.client.openWorkbookSession(new Uint8Array());
      await warm.dispose();

      const custom = await h.client.openWorkbookSession(new Uint8Array(), { worker: h.worker, wasm });
      expect(compileRequests(h.workers[1])).toEqual([]);
      if (kind === 'bytes') {
        expect(input(h.workers[1]).wasm).not.toBe(wasm);
        expect(new Uint8Array(h.workers[1].sources[0] as ArrayBuffer)).toEqual(bytes);
      } else expect(input(h.workers[1]).wasm).toBe(wasm);
      expect(bytes.byteLength).toBe(8);
      expect(h.compiles).toBe(1);
      await custom.dispose();

      const cached = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[2]).wasm).toBe(h.workers[0].modules[0]);
      expect(h.compiles).toBe(1);
      await cached.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('retries a rejected compile in the next worker', async () => {
    const h = factory();
    const error = new Error('Compile failed');
    h.rejectCompile(error);
    const spies = mainThreadWasmSpies();
    try {
      await expect(h.client.openWorkbookSession(new Uint8Array())).rejects.toThrow('Compile failed');
      expect(h.workers[0].terminated).toBe(true);
      expect(h.workers[0].modules).toEqual([]);

      const second = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[1]).wasm).toBeUndefined();
      expect(h.compiles).toBe(2);
      await second.dispose();

      const third = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[2]).wasm).toBe(h.workers[1].modules[0]);
      expect(h.compiles).toBe(2);
      await third.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('recompiles after initialization fails without caching the failed module', async () => {
    const h = factory();
    h.rejectInitialize(new Error('Initialize failed'));
    const spies = mainThreadWasmSpies();
    try {
      await expect(h.client.openWorkbookSession(new Uint8Array())).rejects.toThrow('Initialize failed');
      expect(h.workers[0].terminated).toBe(true);
      expect(h.workers[0].sources[0]).toBeInstanceOf(WebAssembly.Module);
      expect(h.workers[0].modules).toEqual([]);
      expect(h.compiles).toBe(1);

      const second = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[1]).wasm).toBeUndefined();
      expect(compileRequests(h.workers[1])).toEqual([{ protocol: 1, kind: 'wasm-compile' }]);
      expect(h.workers[1].modules).toHaveLength(1);
      expect(h.compiles).toBe(2);
      await second.dispose();

      const third = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[2]).wasm).toBe(h.workers[1].modules[0]);
      expect(compileRequests(h.workers[2])).toEqual([]);
      expect(h.workers[2].modules).toEqual([]);
      expect(h.compiles).toBe(2);
      await third.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('opens independently while the first compile is pending and survives its worker crashing', async () => {
    const h = factory();
    const release = h.holdCompile();
    const spies = mainThreadWasmSpies();
    try {
      const pending = h.client.openWorkbookSession(new Uint8Array()).catch((error) => error);
      expect(input(h.workers[0]).wasm).toBeUndefined();
      const second = await h.client.openWorkbookSession(new Uint8Array());
      expect(h.compiles).toBe(2);
      expect(input(h.workers[1]).wasm).toBeUndefined();
      h.workers[0].crash();
      expect(await pending).toMatchObject({ code: 'crash' });
      release();
      await second.dispose();

      const third = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[2]).wasm).toBe(h.workers[1].modules[0]);
      expect(h.compiles).toBe(2);
      await third.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      release();
      for (const spy of spies) spy.mockRestore();
    }
  });

  test('keys compiled modules by asset URL', async () => {
    const h = factory();
    const spies = mainThreadWasmSpies();
    try {
      const first = await h.client.openWorkbookSession(new Uint8Array());
      await first.dispose();
      h.assetUrl = 'https://example.test/other.wasm';
      const second = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[1]).wasm).toBeUndefined();
      expect(h.compiles).toBe(2);
      await second.dispose();

      h.assetUrl = 'https://example.test/session.wasm';
      const third = await h.client.openWorkbookSession(new Uint8Array());
      expect(input(h.workers[2]).wasm).toBe(h.workers[0].modules[0]);
      expect(h.compiles).toBe(2);
      await third.dispose();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });
});
