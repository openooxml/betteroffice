import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { editWasmModule, editWasmState } from './edit';
import { wasmModuleMemories } from './loadWasmAsset';
import { preloadOpcWasm, rezipContainer, unzipContainer } from './opc';
import {
  PARAGRAPH_PATCH_WRITER_VERSION,
  paragraphPatchWriterVersionWire,
  parseDocxS9Wire,
  parseRelationshipsXmlWire,
  preloadParseWasm,
} from './parse';
import { routingDocxEntries } from './__fixtures__/routingDocx';

const lightLabels = () => wasmModuleMemories().filter(
  ({ label }) => label === 'ooxml-opc' || label === 'docx-parse'
);

test('OPC and parser preloads reuse the resident engine compile and route calls to it', async () => {
  expect(editWasmState.status()).toBe('idle');
  const compiling = editWasmModule();
  expect(editWasmState.status()).toBe('loading');
  await compiling;
  expect(editWasmState.status()).toBe('loading');
  expect(lightLabels()).toEqual([]);
  await Promise.all([preloadOpcWasm(), preloadParseWasm()]);
  expect(editWasmState.status()).toBe('ready');
  expect(editWasmState.sharedModule()).toBe(compiling);

  const bytes = rezipContainer(routingDocxEntries);
  expect(unzipContainer(bytes)).toEqual(routingDocxEntries);
  expect(JSON.parse(parseDocxS9Wire(bytes, '{}')).wireVersion).toBe(1);
  expect(parseRelationshipsXmlWire(routingDocxEntries['_rels/.rels'], '_rels/.rels'))
    .toContain('word/document.xml');
  expect(paragraphPatchWriterVersionWire()).toBe(PARAGRAPH_PATCH_WRITER_VERSION);
  await Promise.all([preloadOpcWasm(), preloadParseWasm()]);
  expect(lightLabels()).toEqual([]);
  expect(wasmModuleMemories().map(({ label }) => label)).toEqual(['docx-edit']);

  await preloadOpcWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, 'generated/opc/ooxml_opc_bg.wasm')
  )));
  await preloadParseWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, 'generated/parse/docx_parse_bg.wasm')
  )));
  expect(lightLabels().map(({ label }) => label).sort()).toEqual(['docx-parse', 'ooxml-opc']);
  expect(unzipContainer(bytes)).toEqual(routingDocxEntries);
  expect(JSON.parse(parseDocxS9Wire(bytes, '{}')).wireVersion).toBe(1);
});
