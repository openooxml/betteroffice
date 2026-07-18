/**
 * Loader for the docx-parse wasm (the Rust DOCX parser/serializer wire).
 * External-asset pattern — see ./loadWasmAsset.ts for the init contract and
 * URL-geometry invariant.
 */

import wasmInit, {
  initSync,
  parse_docx_s9,
  parse_relationships_xml,
  serialize_docx_s10,
  serialize_docx_s11,
  serialize_docx_s12,
  write_docx_s13_wasm,
} from './generated/parse/docx_parse.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

const state = createWasmModuleState({
  label: 'docx-parse',
  preloadName: 'preloadParseWasm',
  assetUrl: () => new URL('./generated/parse/docx_parse_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

/** Load + instantiate the parser wasm (browser path). Idempotent. */
export function preloadParseWasm(input?: WasmAsyncInput): Promise<void> {
  return state.preload(input);
}

export function parseDocxS9Wire(data: Uint8Array, optionsJson: string): string {
  state.ensure();
  return parse_docx_s9(data, optionsJson);
}

export function parseRelationshipsXmlWire(data: Uint8Array, partPath: string): string {
  state.ensure();
  return parse_relationships_xml(data, partPath);
}

export function serializeDocxS10Wire(requestJson: string): string {
  state.ensure();
  return serialize_docx_s10(requestJson);
}

export function serializeDocxS11Wire(requestJson: string): string {
  state.ensure();
  return serialize_docx_s11(requestJson);
}

export function serializeDocxS12Wire(requestJson: string): string {
  state.ensure();
  return serialize_docx_s12(requestJson);
}

export function writeDocxS13Wire(requestJson: string, originalDocx: Uint8Array): Uint8Array {
  state.ensure();
  return write_docx_s13_wasm(requestJson, originalDocx);
}
