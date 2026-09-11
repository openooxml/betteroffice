import { buildWasmModules } from './wasm.ts';

await buildWasmModules([
  {
    crate: 'xlsx-wasm',
    name: 'xlsx_wasm',
    generated: 'packages/xlsx/src/wasm/generated',
  },
]);
