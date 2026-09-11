import { buildWasmModules } from './wasm.ts';

await buildWasmModules([
  {
    crate: 'pptx-wasm',
    name: 'pptx_wasm',
    generated: 'packages/pptx/src/wasm/generated',
  },
]);
