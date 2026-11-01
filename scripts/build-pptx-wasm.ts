import { copyFile, readFile } from 'node:fs/promises';
import { buildWasmModules } from './wasm.ts';

const generated = new URL('../packages/pptx/src/wasm/generated/', import.meta.url);
const lean = new URL('../target/wasm-pack/pptx-wasm-lean/generated/', import.meta.url);

await buildWasmModules([
  {
    crate: 'pptx-wasm',
    name: 'pptx_wasm',
    generated: 'packages/pptx/src/wasm/generated',
  },
  {
    crate: 'pptx-wasm',
    name: 'pptx_wasm',
    generated: 'target/wasm-pack/pptx-wasm-lean/generated',
    cargoArgs: ['--no-default-features', '--locked'],
    variant: 'lean',
  },
]);

for (const file of ['pptx_wasm.js', 'pptx_wasm.d.ts', 'pptx_wasm_bg.wasm.d.ts']) {
  const fullGlue = await readFile(new URL(file, generated));
  const leanGlue = await readFile(new URL(file, lean));
  if (!fullGlue.equals(leanGlue)) throw new Error(`pptx full and editor-core wasm glue differ: ${file}`);
}

const module = await WebAssembly.compile(await readFile(new URL('pptx_wasm_bg.wasm', lean)));
if (WebAssembly.Module.customSections(module, 'betteroffice.pptx.editor-core').length !== 1) {
  throw new Error('pptx editor-core wasm is missing its marker section');
}

await copyFile(new URL('pptx_wasm_bg.wasm', lean), new URL('pptx_wasm_lean_bg.wasm', generated));
