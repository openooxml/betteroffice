import { defineConfig } from 'tsup';
import { copyFile, mkdir, rename } from 'node:fs/promises';
import { join } from 'node:path';

export default defineConfig((options) => ({
  entry: {
    index: 'src/index.ts',
    'wasm-preflight': 'src/wasm/preflight.ts',
    headless: 'src/headless.ts',
    worker: 'src/worker.ts',
    collaboration: 'src/collaboration/index.ts',
    xlsxSessionWorker: 'src/session/worker.ts',
    'xlsxSessionWorker.node': 'src/session/worker.node.ts',
    'sessionWorker.node': 'src/session/sessionWorker.node.ts',
    'sessionWorker.missing': 'src/session/sessionWorker.missing.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
  external: ['#session-worker'],
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    const out = options.outDir ?? 'dist';
    await rename(join(out, 'xlsxSessionWorker.js'), join(out, 'xlsxSessionWorker.mjs'));
    await rename(join(out, 'xlsxSessionWorker.node.js'), join(out, 'xlsxSessionWorker.node.mjs'));
    await mkdir(join(out, 'generated'), { recursive: true });
    await mkdir(join(out, 'generated/preflight'), { recursive: true });
    await copyFile('src/wasm/generated/preflight/xlsx_preflight_bg.wasm', join(out, 'generated/preflight/xlsx_preflight_bg.wasm'));
    await copyFile(
      'src/wasm/generated/xlsx_wasm_bg.wasm',
      join(out, 'generated/xlsx_wasm_bg.wasm')
    );
    await copyFile(
      'src/wasm/generated/xlsx_wasm_lean_bg.wasm',
      join(out, 'generated/xlsx_wasm_lean_bg.wasm')
    );
    await copyFile(
      'src/wasm/generated/xlsx_wasm_lean_bg.wasm',
      'dist/generated/xlsx_wasm_lean_bg.wasm'
    );
  },
}));
