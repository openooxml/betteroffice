export function wasmAssetUrl(): URL {
  return new URL('./generated/pptx_wasm_bg.wasm', import.meta.url);
}

/** The session worker's core: the full build without the one-shot `compileSlideJson` and `renderPptxMarkdownJson`. */
export function sessionWasmAssetUrl(): URL {
  return new URL('./generated/pptx_wasm_lean_bg.wasm', import.meta.url);
}

/** Whether this build ships the wasm asset; the native package builds these sources with its own asset module. */
export const shipsWasmAsset = true;
