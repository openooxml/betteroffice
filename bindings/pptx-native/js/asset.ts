export function wasmAssetUrl(): URL {
  return new URL('betteroffice-native:pptx');
}

export function sessionWasmAssetUrl(): URL {
  return wasmAssetUrl();
}

export function metafileWasmAssetUrl(): URL {
  return wasmAssetUrl();
}

export function metafileSvgBindings(): Promise<never> {
  return Promise.reject(new Error('The native engine converts metafiles itself'));
}

export const shipsWasmAsset = false;
