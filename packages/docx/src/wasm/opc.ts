/**
 * Loader for the ooxml-opc wasm (the OPC zip container: unzip/rezip with
 * zip-bomb and path-traversal guards). External-asset pattern: see
 * ./loadWasmAsset.ts for the init contract and URL-geometry invariant.
 */

import wasmInit, { initSync, engineKind, unzip_docx, rezip_docx } from './engineOpc';
import { wasmAssetUrl } from './assetOpc';
import { createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';
import { residentEngine } from './residentEngine';

const state = createWasmModuleState({
  label: 'ooxml-opc',
  preloadName: 'preloadOpcWasm',
  assetUrl: wasmAssetUrl,
  engineKind,
  initAsync: wasmInit,
  initSync,
});

/** Load + instantiate the container wasm (browser path). Idempotent. */
export function preloadOpcWasm(input?: WasmAsyncInput): Promise<void> {
  const engine = residentEngine();
  if (input === undefined && engine && engine.state.status() !== 'idle') {
    return engine.state.preload();
  }
  return state.preload(input);
}

/** Inflate a DOCX container into `{ [path]: Uint8Array }`. */
export function unzipContainer(data: Uint8Array): Record<string, Uint8Array> {
  const engine = residentEngine();
  if (engine?.state.status() === 'ready') {
    const glue = engine.glue as { unzip_docx: typeof unzip_docx };
    return glue.unzip_docx(data) as Record<string, Uint8Array>;
  }
  state.ensure();
  return unzip_docx(data) as Record<string, Uint8Array>;
}

/** Deflate `{ [path]: Uint8Array }` into DOCX bytes. */
export function rezipContainer(entries: Record<string, Uint8Array>): Uint8Array {
  const engine = residentEngine();
  if (engine?.state.status() === 'ready') {
    return (engine.glue as { rezip_docx: typeof rezip_docx }).rezip_docx(entries);
  }
  state.ensure();
  return rezip_docx(entries) as Uint8Array;
}
