import { defineConfig } from 'tsup';
import { copyFile, mkdir, rename } from 'node:fs/promises';
import { join } from 'node:path';

export default defineConfig((options) => ({
  entry: {
    index: 'src/index.ts',
    headless: 'src/headless.ts',
    collaboration: 'src/collaboration/index.ts',
    xlsxSessionWorker: 'src/session/worker.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
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
    await mkdir('dist/generated', { recursive: true });
    await copyFile(
      'src/wasm/generated/xlsx_wasm_bg.wasm',
      'dist/generated/xlsx_wasm_bg.wasm'
    );
  },
}));
