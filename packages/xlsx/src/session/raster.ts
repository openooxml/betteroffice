import { rasterWasmAssetUrl } from '../wasm/asset';

type Raster = typeof import('../wasm/generated/raster/xlsx_raster.js');
let initialization: Promise<Raster> | undefined;

export async function renderPng(displayListJson: string): Promise<Uint8Array> {
  initialization ??= import('../wasm/generated/raster/xlsx_raster.js').then(async (raster) => {
    await raster.default({ module_or_path: rasterWasmAssetUrl() });
    return raster;
  });
  const raster = await initialization;
  return raster.render_png(displayListJson);
}
