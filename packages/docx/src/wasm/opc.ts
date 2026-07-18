/**
 * Loader for the ooxml-opc wasm (the OPC zip container: unzip/rezip with
 * zip-bomb and path-traversal guards). External-asset pattern — see
 * ./loadWasmAsset.ts for the init contract and URL-geometry invariant.
 */

import wasmInit, { initSync, unzip_docx, rezip_docx } from './generated/opc/ooxml_opc.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

const state = createWasmModuleState({
  label: 'ooxml-opc',
  preloadName: 'preloadOpcWasm',
  assetUrl: () => new URL('./generated/opc/ooxml_opc_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

/** Load + instantiate the container wasm (browser path). Idempotent. */
export function preloadOpcWasm(input?: WasmAsyncInput): Promise<void> {
  return state.preload(input);
}

/** Inflate a DOCX container into `{ [path]: Uint8Array }`. */
export function unzipContainer(data: Uint8Array): Record<string, Uint8Array> {
  state.ensure();
  return unzip_docx(data) as Record<string, Uint8Array>;
}

/** Deflate `{ [path]: Uint8Array }` into DOCX bytes. */
export function rezipContainer(entries: Record<string, Uint8Array>): Uint8Array {
  state.ensure();
  return rezip_docx(entries) as Uint8Array;
}
