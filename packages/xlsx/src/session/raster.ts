import { rasterWasmAssetUrl } from '../wasm/asset';

type Raster = typeof import('../wasm/engineRaster');
let initialization: Promise<Raster> | undefined;

export async function renderPng(displayListJson: string): Promise<Uint8Array> {
  const pending = initialization ??= import('../wasm/engineRaster').then(async (raster) => {
    await raster.default({ module_or_path: rasterWasmAssetUrl() });
    return raster;
  });
  let raster: Raster;
  try { raster = await pending; }
  catch (error) {
    if (initialization === pending) initialization = undefined;
    throw error;
  }
  return raster.render_png(displayListJson);
}
