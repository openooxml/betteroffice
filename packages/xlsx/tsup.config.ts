import { defineConfig } from 'tsup';
import { copyFile, mkdir, rename } from 'node:fs/promises';
import { join } from 'node:path';

export default defineConfig((options) => ({
  entry: {
    index: 'src/index.ts',
    headless: 'src/headless.ts',
    worker: 'src/worker.ts',
    collaboration: 'src/collaboration/index.ts',
    xlsxSessionWorker: 'src/session/worker.ts',
  },
  format: ['esm'],
  external: ['@betteroffice/xlsx-native'],
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
    await mkdir(join(out, 'generated/raster'), { recursive: true });
    await copyFile(
      'src/wasm/generated/xlsx_wasm_bg.wasm',
      join(out, 'generated/xlsx_wasm_bg.wasm')
    );
    await copyFile(
      'src/wasm/generated/xlsx_wasm_lean_bg.wasm',
      join(out, 'generated/xlsx_wasm_lean_bg.wasm')
    );
    await copyFile(
      'src/wasm/generated/raster/xlsx_raster_bg.wasm',
      join(out, 'generated/raster/xlsx_raster_bg.wasm')
    );
  },
}));
