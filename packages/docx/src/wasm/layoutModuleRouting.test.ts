import { expect, test } from 'bun:test';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { editWasmModule, editWasmState, preloadEditWasm } from './edit';
import { wasmModuleMemories } from './loadWasmAsset';
import { PARAGRAPH_PATCH_WRITER_VERSION, paragraphPatchWriterVersionWire } from './parse';

test('main-thread hit testing leaves the worker engine compiled without instantiating it', async () => {
  expect(editWasmState.status()).toBe('idle');
  const compiling = editWasmModule();
  await compiling;
  expect(editWasmState.status()).toBe('loading');
  const instances: WebAssembly.Instance[] = [];
  const counted = (result: WebAssembly.Instance | WebAssembly.WebAssemblyInstantiatedSource) => {
    instances.push(result instanceof WebAssembly.Instance ? result : result.instance);
    return result;
  };
  const instantiate = WebAssembly.instantiate;
  const streaming = WebAssembly.instantiateStreaming;
  const Instance = WebAssembly.Instance;
  WebAssembly.instantiate = function (...args: Parameters<typeof instantiate>) {
    return Reflect.apply(instantiate, WebAssembly, args).then(counted);
  } as typeof instantiate;
  WebAssembly.instantiateStreaming = function (...args: Parameters<typeof streaming>) {
    return Reflect.apply(streaming, WebAssembly, args).then(counted);
  } as typeof streaming;
  WebAssembly.Instance = new Proxy(Instance, {
    construct(target, args, newTarget) {
      return counted(Reflect.construct(target, args, newTarget));
    },
  });
  try {
    const engine = await loadRustDisplayListQueryEngine();
    const handle = engine.openDisplayList!(JSON.stringify({ pages: [{
      pageIndex: 0, width: 100, height: 100, primitives: [{
        kind: 'text', text: 'abcd', x: 10, baselineY: 20, width: 40,
        font: '10px serif', color: '#000', docStart: 1, docEnd: 5,
      }],
    }] }));
    try {
      const hit = engine.hitTestRegionsByHandle!(handle, 0, 15, 15);
      const rects = engine.rangeRectsByHandle!(handle, 1, 3);
      expect(JSON.parse(hit)).toMatchObject({ region: 'body', target: 'text' });
      expect(JSON.parse(rects)).toHaveLength(1);
      expect(instances.length).toBeGreaterThan(0);
      expect(instances.filter(instance => 'editsession_new' in instance.exports)).toEqual([]);
      expect(editWasmState.status()).toBe('loading');
      expect(editWasmState.sharedModule()).toBe(compiling);
      expect(wasmModuleMemories().map(({ label }) => label)).toEqual(['docx-parse']);
      expect(paragraphPatchWriterVersionWire()).toBe(PARAGRAPH_PATCH_WRITER_VERSION);

      await preloadEditWasm();
      expect(editWasmState.status()).toBe('ready');
      expect(engine.hitTestRegionsByHandle!(handle, 0, 15, 15)).toBe(hit);
      expect(engine.rangeRectsByHandle!(handle, 1, 3)).toBe(rects);
    } finally {
      engine.closeDisplayList!(handle);
    }
  } finally {
    WebAssembly.instantiate = instantiate;
    WebAssembly.instantiateStreaming = streaming;
    WebAssembly.Instance = Instance;
  }
});
