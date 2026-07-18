/**
 * Loader for the docx-edit wasm (the Rust yrs editing core + resident layout
 * engine). External-asset pattern — see ./loadWasmAsset.ts for the init
 * contract and URL-geometry invariant.
 *
 * Reached only via the `src/yrs/` facade's dynamic `import()`, so the glue
 * (and the wasm fetch) stay out of every non-editor bundle path. The facade's
 * async `createYrsSession` awaits {@link preloadEditWasm} before constructing
 * a session, which also covers the resident engine worker (it bootstraps
 * through the same facade inside the worker context).
 */

import wasmInit, { initSync, EditSession } from './generated/edit/docx_edit.js';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

const state = createWasmModuleState({
  label: 'docx-edit',
  preloadName: 'preloadEditWasm',
  assetUrl: () => new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url),
  initAsync: wasmInit,
  initSync,
});

/** Load + instantiate the editing-core wasm (browser path). Idempotent. */
export function preloadEditWasm(input?: WasmAsyncInput): Promise<void> {
  return state.preload(input);
}

/**
 * Constructs a raw wasm `EditSession` replica after ensuring the module is
 * initialized. The `src/yrs/` facade wraps this in the typed `YrsSession`
 * surface — nothing else should call it.
 */
export function createEditSession(clientId: number): EditSession {
  state.ensure();
  return new EditSession(clientId);
}

export type { EditSession };
