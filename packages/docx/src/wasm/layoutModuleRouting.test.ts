import { expect, test } from 'bun:test';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { editWasmModule, editWasmState, preloadEditWasm } from './edit';
import * as editGlue from './generated/edit/docx_edit.js';
import { wasmModuleMemories } from './loadWasmAsset';
import { paragraphPatchWriterVersionWire, parseWasmState } from './parse';
import { queryWasmState } from './query';

test('main-thread queries use only query wasm with an idle or compiled worker engine', async () => {
  expect(editWasmState.status()).toBe('idle');
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
    expect(editWasmState.status()).toBe('idle');
    expect(parseWasmState.status()).toBe('idle');
    expect(wasmModuleMemories().map(({ label }) => label)).toEqual(['docx-query']);
    const compiling = editWasmModule();
    await compiling;
    expect(editWasmState.status()).toBe('loading');
    const displayList = JSON.stringify({ pages: [{
      pageIndex: 0, width: 100, height: 100, primitives: [{
        kind: 'text', text: 'abcd', x: 10, baselineY: 20, width: 40,
        font: '10px serif', color: '#000', docStart: 1, docEnd: 5,
      }, {
        kind: 'text', text: 'efgh', x: 10, baselineY: 40, width: 40,
        font: '10px serif', color: '#000', docStart: 6, docEnd: 10,
      }],
    }] });
    const handle = engine.openDisplayList!(displayList);
    try {
      const hit = engine.hitTestRegionsByHandle!(handle, 0, 15, 15);
      const rects = engine.rangeRectsByHandle!(handle, 1, 3);
      expect(JSON.parse(hit)).toMatchObject({ region: 'body', target: 'text' });
      expect(JSON.parse(rects)).toHaveLength(1);
      expect(instances.length).toBeGreaterThan(0);
      expect(instances.filter(instance => 'editsession_new' in instance.exports)).toEqual([]);
      for (const instance of instances) {
        expect(Object.keys(instance.exports).filter(name =>
          /parse_docx|layout_document|build_display_list|measure_|outline_glyph/.test(name)
        )).toEqual([]);
      }
      expect(editWasmState.status()).toBe('loading');
      expect(editWasmState.sharedModule()).toBe(compiling);
      expect(wasmModuleMemories().map(({ label }) => label)).toEqual(['docx-query']);
      expect(queryWasmState.status()).toBe('ready');
      expect(parseWasmState.status()).toBe('idle');
      expect(paragraphPatchWriterVersionWire()).toBeNull();

      await preloadEditWasm();
      expect(editWasmState.status()).toBe('ready');
      expect(engine.hitTestRegionsByHandle!(handle, 0, 15, 15)).toBe(hit);
      expect(engine.rangeRectsByHandle!(handle, 1, 3)).toBe(rects);
      expect(engine.hitTestRegionsJson(displayList, 0, 15, 15))
        .toBe(editGlue.hit_test_regions_json(displayList, 0, 15, 15));
      expect(engine.rangeRectsJson(displayList, 1, 3)).toBe(editGlue.range_rects_json(displayList, 1, 3));
      expect(engine.verticalMoveJson!(displayList, 2, 'down', 20))
        .toBe(editGlue.vertical_move_json(displayList, 2, 'down', 20));
      expect(engine.rangeRectsOnPagesJson!(displayList, 1, 3, 0, 0))
        .toBe(editGlue.range_rects_on_pages_json(displayList, 1, 3, 0, 0));
      expect(engine.rangeRectsRegionJson!(displayList, 'body', '', 1, 3))
        .toBe(editGlue.range_rects_region_json(displayList, 'body', '', 1, 3));
      const residentHandle = editGlue.open_display_list(displayList);
      try {
        expect(engine.hitTestRegionsByHandle!(handle, 0, 15, 15))
          .toBe(editGlue.hit_test_regions_by_handle(residentHandle, 0, 15, 15));
        expect(engine.verticalMoveByHandle!(handle, 2, 'down', 20))
          .toBe(editGlue.vertical_move_by_handle(residentHandle, 2, 'down', 20));
        expect(engine.rangeRectsByHandle!(handle, 1, 3))
          .toBe(editGlue.range_rects_by_handle(residentHandle, 1, 3));
        expect(engine.rangeRectsOnPagesByHandle!(handle, 1, 3, 0, 0))
          .toBe(editGlue.range_rects_on_pages_by_handle(residentHandle, 1, 3, 0, 0));
        expect(engine.rangeRectsRegionByHandle!(handle, 'body', '', 1, 3))
          .toBe(editGlue.range_rects_region_by_handle(residentHandle, 'body', '', 1, 3));
        const update = JSON.stringify({ total: 1, keep: true, replace: [[0, {
          pageIndex: 0, width: 100, height: 100, primitives: [],
        }]] });
        engine.updateDisplayList!(handle, update);
        editGlue.update_display_list(residentHandle, update);
        expect(engine.hitTestRegionsByHandle!(handle, 0, 15, 15))
          .toBe(editGlue.hit_test_regions_by_handle(residentHandle, 0, 15, 15));
        expect(engine.rangeRectsByHandle!(handle, 1, 3)).toBe('[]');
        expect(parseWasmState.status()).toBe('idle');
      } finally {
        editGlue.close_display_list(residentHandle);
      }
    } finally {
      engine.closeDisplayList!(handle);
    }
  } finally {
    WebAssembly.instantiate = instantiate;
    WebAssembly.instantiateStreaming = streaming;
    WebAssembly.Instance = Instance;
  }
});
