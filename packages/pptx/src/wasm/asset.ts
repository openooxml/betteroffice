export function wasmAssetUrl(): URL {
  return new URL('./generated/pptx_wasm_bg.wasm', import.meta.url);
}

/** The session worker's core: the full build without one-shot compilation and metafile SVG conversion. */
export function sessionWasmAssetUrl(): URL {
  return new URL('./generated/pptx_wasm_lean_bg.wasm', import.meta.url);
}

/** The metafile SVG converter the session worker loads on demand. */
export function metafileWasmAssetUrl(): URL {
  return new URL('./generated/metafile_svg_wasm_bg.wasm', import.meta.url);
}

/** The converter's bindings, imported only when a session needs them. */
export function metafileSvgBindings(): Promise<typeof import('./generated/metafile_svg_wasm')> {
  return import('./generated/metafile_svg_wasm');
}

/** Whether this build ships the wasm asset; the native package builds these sources with its own asset module. */
export const shipsWasmAsset = true;
