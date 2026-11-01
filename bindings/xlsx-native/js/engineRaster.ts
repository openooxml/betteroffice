import type * as Wasm from '../../../packages/xlsx/src/wasm/generated/raster/xlsx_raster.js';

export type { InitInput, InitOutput, SyncInitInput } from '../../../packages/xlsx/src/wasm/generated/raster/xlsx_raster.js';

function unavailable(): never {
  throw new Error('@betteroffice/xlsx-native renders PNG in the native engine and ships no raster wasm');
}

export const engineKind: 'wasm' | 'native' = 'native';

export const render_png = unavailable as unknown as typeof Wasm.render_png;

export const initSync = unavailable as unknown as typeof Wasm.initSync;

export default (async () => unavailable()) as unknown as typeof Wasm.default;
