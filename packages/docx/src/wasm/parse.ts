/**
 * Loader for the docx-parse wasm (the Rust DOCX parser/serializer wire).
 * External-asset pattern — see ./loadWasmAsset.ts for the init contract and
 * URL-geometry invariant.
 */

import * as parseModule from './generated/parse/docx_parse.js';
import wasmInit, { initSync } from './generated/parse/docx_parse.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';
import { residentEngine } from './residentEngine';

export const parseWasmState = createWasmModuleState({
  label: 'docx-parse',
  preloadName: 'preloadParseWasm',
  assetUrl: () => new URL('./generated/parse/docx_parse_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

type ParseExports = Pick<
  typeof parseModule,
  | 'decodeTiffPng'
  | 'parse_docx_s9'
  | 'parse_relationships_xml'
  | 'serialize_docx_s10'
  | 'serialize_docx_s11'
  | 'serialize_docx_s12'
  | 'write_docx_s13_wasm'
  | 'write_docx_s13_wasm_with_warnings'
>;

function parserGlue(): ParseExports {
  const engine = residentEngine();
  if (engine?.state.status() === 'ready') return engine.glue as ParseExports;
  parseWasmState.ensure();
  return parseModule;
}

/** Load + instantiate the parser wasm (browser path). Idempotent. */
export function preloadParseWasm(input?: WasmAsyncInput): Promise<void> {
  const engine = residentEngine();
  if (input === undefined && engine && engine.state.status() !== 'idle') {
    return engine.state.preload();
  }
  return parseWasmState.preload(input);
}

export const PARAGRAPH_PATCH_WRITER_VERSION = 1;

export function paragraphPatchWriterVersionWire(): number | null | undefined {
  const engine = residentEngine();
  const engineReady = engine?.state.status() === 'ready';
  const capability = ((engineReady ? engine.glue : parseModule) as {
    paragraphPatchWriterVersion?: () => number;
  }).paragraphPatchWriterVersion;
  if (typeof capability !== 'function') return undefined;
  if (!engineReady && parseWasmState.status() !== 'ready') return null;
  return capability();
}

export function parseDocxS9Wire(data: Uint8Array, optionsJson: string): string {
  return parserGlue().parse_docx_s9(data, optionsJson);
}

export function parseRelationshipsXmlWire(data: Uint8Array, partPath: string): string {
  return parserGlue().parse_relationships_xml(data, partPath);
}

export function serializeDocxS10Wire(requestJson: string): string {
  return parserGlue().serialize_docx_s10(requestJson);
}

export function serializeDocxS11Wire(requestJson: string): string {
  return parserGlue().serialize_docx_s11(requestJson);
}

export function serializeDocxS12Wire(requestJson: string): string {
  return parserGlue().serialize_docx_s12(requestJson);
}

export function writeDocxS13Wire(requestJson: string, originalDocx: Uint8Array): Uint8Array {
  return parserGlue().write_docx_s13_wasm(requestJson, originalDocx);
}

export function writeDocxS13WithWarningsWire(
  requestJson: string,
  originalDocx: Uint8Array
): { bytes: Uint8Array; warnings: string[] } {
  const framed = parserGlue().write_docx_s13_wasm_with_warnings(requestJson, originalDocx);
  if (framed.byteLength < 8) throw new Error('Rust save response is truncated');
  const warningsLength = Number(
    new DataView(framed.buffer, framed.byteOffset, framed.byteLength).getBigUint64(0, true)
  );
  if (!Number.isSafeInteger(warningsLength) || 8 + warningsLength > framed.byteLength) {
    throw new Error('Rust save response carries a corrupt warnings frame');
  }
  const warnings: unknown = JSON.parse(
    new TextDecoder().decode(framed.subarray(8, 8 + warningsLength))
  );
  if (!Array.isArray(warnings) || !warnings.every((entry): entry is string => typeof entry === 'string')) {
    throw new Error('Rust save response carries malformed warnings');
  }
  return { bytes: framed.subarray(8 + warningsLength), warnings };
}

export function decodeTiffImage(bytes: Uint8Array): Uint8Array {
  const glue = parserGlue();
  try {
    return glue.decodeTiffPng(bytes);
  } catch (error) {
    throw error instanceof Error
      ? error
      : new Error(typeof error === 'string' ? error : String(error));
  }
}
