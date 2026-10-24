import { copyFile, mkdir, rename } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'tsup';
import { packageBoundary } from '../../scripts/package-boundary';

export default defineConfig((options) => ({
  entry: {
    index: 'src/index.ts',
    editor: 'src/editor.ts',
    worker: 'src/worker.ts',
    collaboration: 'src/collaboration/index.ts',
    pptxSessionWorker: 'src/session/worker.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  esbuildPlugins: [packageBoundary(fileURLToPath(new URL('.', import.meta.url)))],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    const out = options.outDir ?? 'dist';
    await rename(join(out, 'pptxSessionWorker.js'), join(out, 'pptxSessionWorker.mjs'));
    await mkdir('dist/generated', { recursive: true });
    await copyFile('src/wasm/generated/pptx_wasm_bg.wasm', 'dist/generated/pptx_wasm_bg.wasm');
  },
}));
