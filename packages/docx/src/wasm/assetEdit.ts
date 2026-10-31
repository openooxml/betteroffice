export const shipsWasmAsset: boolean = true;

export function wasmAssetUrl(): URL {
  return new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url);
}
