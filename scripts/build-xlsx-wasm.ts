import { copyFile, readFile } from 'node:fs/promises';
import { buildWasmModules } from './wasm.ts';

const generated = new URL('../packages/xlsx/src/wasm/generated/', import.meta.url);
const lean = new URL('../target/wasm-pack/xlsx-wasm-lean/generated/', import.meta.url);

await buildWasmModules([
  {
    crate: 'xlsx-wasm',
    name: 'xlsx_wasm',
    generated: 'packages/xlsx/src/wasm/generated',
  },
  {
    crate: 'xlsx-wasm',
    name: 'xlsx_wasm',
    generated: 'target/wasm-pack/xlsx-wasm-lean/generated',
    cargoArgs: ['--no-default-features', '--locked'],
    variant: 'lean',
  },
]);

for (const file of ['xlsx_wasm.js', 'xlsx_wasm.d.ts', 'xlsx_wasm_bg.wasm.d.ts']) {
  const fullGlue = await readFile(new URL(file, generated));
  const leanGlue = await readFile(new URL(file, lean));
  if (!fullGlue.equals(leanGlue)) throw new Error(`xlsx full and lean wasm glue differ: ${file}`);
}

const module = await WebAssembly.compile(await readFile(new URL('xlsx_wasm_bg.wasm', lean)));
if (WebAssembly.Module.customSections(module, 'betteroffice.xlsx.no-png').length !== 1) {
  throw new Error('xlsx lean wasm is missing its PNG capability marker');
}

await copyFile(new URL('xlsx_wasm_bg.wasm', lean), new URL('xlsx_wasm_lean_bg.wasm', generated));
