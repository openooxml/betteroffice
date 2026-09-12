import { buildWasmModules } from './wasm.ts';

await buildWasmModules([
  {
    crate: 'vsdx-wasm',
    name: 'vsdx_wasm',
    generated: 'packages/vsdx/src/wasm/generated',
  },
]);
