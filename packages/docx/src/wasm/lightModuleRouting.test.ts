import { expect, test } from 'bun:test';
import { wasmModuleMemories } from './loadWasmAsset';
import { preloadOpcWasm, rezipContainer, unzipContainer } from './opc';
import {
  PARAGRAPH_PATCH_WRITER_VERSION,
  paragraphPatchWriterVersionWire,
  parseDocxS9Wire,
  parseRelationshipsXmlWire,
  preloadParseWasm,
} from './parse';
import { residentEngine } from './residentEngine';
import { routingDocxEntries } from './__fixtures__/routingDocx';

test('parse-only calls use light modules with no engine or an idle engine', async () => {
  expect(residentEngine()).toBeUndefined();
  expect(wasmModuleMemories()).toEqual([]);
  expect(paragraphPatchWriterVersionWire()).toBeNull();
  await Promise.all([preloadOpcWasm(), preloadParseWasm()]);
  const bytes = rezipContainer(routingDocxEntries);
  expect(unzipContainer(bytes)).toEqual(routingDocxEntries);
  expect(JSON.parse(parseDocxS9Wire(bytes, '{}')).wireVersion).toBe(1);
  expect(parseRelationshipsXmlWire(routingDocxEntries['_rels/.rels'], '_rels/.rels'))
    .toContain('word/document.xml');
  expect(paragraphPatchWriterVersionWire()).toBe(PARAGRAPH_PATCH_WRITER_VERSION);
  expect(wasmModuleMemories().map(({ label }) => label).sort()).toEqual(['docx-parse', 'ooxml-opc']);
  expect(residentEngine()).toBeUndefined();

  const { editWasmState } = await import('./edit');
  expect(editWasmState.status()).toBe('idle');
  await Promise.all([preloadOpcWasm(), preloadParseWasm()]);
  expect(unzipContainer(bytes)).toEqual(routingDocxEntries);
  expect(JSON.parse(parseDocxS9Wire(bytes, '{}')).wireVersion).toBe(1);
  expect(editWasmState.status()).toBe('idle');
  expect(wasmModuleMemories().map(({ label }) => label).sort()).toEqual(['docx-parse', 'ooxml-opc']);

  const { loadRustDisplayListQueryEngine } = await import('../layout/render/rustDisplayList');
  const queries = await loadRustDisplayListQueryEngine();
  const handle = queries.openDisplayList!('{"pages":[]}');
  try {
    expect(queries.hitTestRegionsByHandle!(handle, 0, 0, 0)).toBe('null');
    expect(editWasmState.status()).toBe('idle');
    expect(wasmModuleMemories().map(({ label }) => label).sort())
      .toEqual(['docx-parse', 'docx-query', 'ooxml-opc']);
  } finally {
    queries.closeDisplayList!(handle);
  }
});
