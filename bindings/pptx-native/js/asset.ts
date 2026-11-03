export function wasmAssetUrl(): URL {
  return new URL('betteroffice-native:pptx');
}

export function sessionWasmAssetUrl(): URL {
  return wasmAssetUrl();
}

export function metafileWasmAssetUrl(): URL {
  return wasmAssetUrl();
}

export const shipsWasmAsset = false;
