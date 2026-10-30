import { expect, test } from 'bun:test';
import { resolve } from 'node:path';
import { engineBackend } from './engine';

test.skipIf(engineBackend() !== 'native')('native initialization ignores wasm inputs and opens a working handle', async () => {
  const script = `
    import assert from 'node:assert/strict';
    const fixture = new Uint8Array(await Bun.file(${JSON.stringify(resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.pptx'))}).arrayBuffer());
    const module = new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
    const blocked = () => { throw new Error('Unexpected wasm fetch, compilation or instantiation'); };
    globalThis.fetch = blocked;
    WebAssembly.Module = new Proxy(WebAssembly.Module, { construct: blocked });
    WebAssembly.Instance = new Proxy(WebAssembly.Instance, { construct: blocked });
    WebAssembly.compile = blocked;
    WebAssembly.compileStreaming = blocked;
    WebAssembly.instantiate = blocked;
    WebAssembly.instantiateStreaming = blocked;
    const { default: initEngine, engineBackend } = await import(${JSON.stringify(resolve(import.meta.dir, 'engine.ts'))});
    const { initWasm, openPresentation, wasmVersion } = await import(${JSON.stringify(resolve(import.meta.dir, 'loader.ts'))});
    assert.equal(engineBackend(), 'native');
    const inputs = [undefined, new URL('https://example.invalid/pptx.wasm'),
      new Uint8Array([1, 2, 3]), module, Promise.resolve(module), new Response('not wasm')];
    for (const input of inputs) {
      assert.deepEqual(await initEngine({ module_or_path: input }), { memory: undefined });
    }
    await initWasm(module);
    for (const input of inputs) await initWasm(input);
    assert.equal(typeof wasmVersion(), 'string');
    const handle = openPresentation(fixture);
    try {
      assert.ok(handle.snapshot().slides.length > 0);
      const slideId = handle.snapshot().slides[0].id;
      handle.setSlideNotes(slideId, 'Native engine selection');
      assert.equal(handle.snapshot().slides[0].notes, 'Native engine selection');
      const reopened = openPresentation(handle.save());
      try { assert.equal(reopened.snapshot().slides[0].notes, 'Native engine selection'); }
      finally { reopened.dispose(); }
    } finally { handle.dispose(); }
  `;
  const child = Bun.spawn([process.execPath, '--eval', script], {
    cwd: resolve(import.meta.dir, '../../../..'), stdout: 'ignore', stderr: 'pipe',
  });
  const error = await new Response(child.stderr).text();
  expect({ exit: await child.exited, error }).toEqual({ exit: 0, error: '' });
});
