import wasmInit, { initSync } from './generated/query/docx_query.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

export const queryWasmState = createWasmModuleState({
  label: 'docx-query',
  preloadName: 'preloadQueryWasm',
  assetUrl: () => new URL('./generated/query/docx_query_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

/** Load the display-list query wasm. */
export function preloadQueryWasm(input?: WasmAsyncInput): Promise<void> {
  return queryWasmState.preload(input);
}
