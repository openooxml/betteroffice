import wasmInit, { initSync, engineKind } from './engineQuery';
import { wasmAssetUrl } from './assetQuery';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

export const queryWasmState = createWasmModuleState({
  label: 'docx-query',
  preloadName: 'preloadQueryWasm',
  assetUrl: wasmAssetUrl,
  engineKind,
  initAsync: wasmInit,
  initSync,
});

/** Load the display-list query wasm. */
export function preloadQueryWasm(input?: WasmAsyncInput): Promise<void> {
  return queryWasmState.preload(input);
}
