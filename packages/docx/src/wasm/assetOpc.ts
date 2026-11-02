export const shipsWasmAsset: boolean = true;

export function wasmAssetUrl(): URL {
  return new URL('./generated/opc/ooxml_opc_bg.wasm', import.meta.url);
}
