// Builds the four docx wasm cores (container / layout / edit / parse) and
// vendors the wasm-pack output into packages/docx/src/wasm/generated/. The glue
// .js and the *_bg.wasm binaries are gitignored and rebuilt on demand
// (predev/prebuild hooks) so the repo never carries multi-MB binaries.
import { buildWasmModules, type WasmModule } from './wasm.ts';

// docx-edit, docx-parse and ooxml-opc keep their wasm-bindgen boundary behind a
// `wasm` cargo feature so native builds never pull that stack.
const MODULES: WasmModule[] = [
  {
    crate: 'ooxml-opc',
    name: 'ooxml_opc',
    generated: 'packages/docx/src/wasm/generated/opc',
    cargoArgs: ['--locked', '--features', 'wasm'],
  },
  {
    crate: 'docx-layout',
    name: 'docx_layout',
    generated: 'packages/docx/src/wasm/generated/layout',
  },
  {
    crate: 'docx-edit',
    name: 'docx_edit',
    generated: 'packages/docx/src/wasm/generated/edit',
    cargoArgs: ['--locked', '--features', 'wasm'],
  },
  {
    crate: 'docx-parse',
    name: 'docx_parse',
    generated: 'packages/docx/src/wasm/generated/parse',
    cargoArgs: ['--locked', '--features', 'wasm'],
  },
];

await buildWasmModules(MODULES);
