import { copyFile, mkdir, rename } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'tsup';
import { packageBoundary } from '../../scripts/package-boundary';

export default defineConfig((options) => ({
  entry: {
    index: 'src/index.ts',
    worker: 'src/worker.ts',
    collaboration: 'src/collaboration/index.ts',
    pptxSessionWorker: 'src/session/worker.ts',
    'sessionWorker.browser': 'src/session/sessionWorker.browser.ts',
    'sessionScope.browser': 'src/session/sessionScope.browser.ts',
    'sessionWorker.node': 'src/session/sessionWorker.node.ts',
    'sessionScope.node': 'src/session/sessionScope.node.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
  external: ['#session-worker', '#session-scope'],
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
    await copyFile('src/wasm/generated/metafile_svg_wasm_bg.wasm', 'dist/generated/metafile_svg_wasm_bg.wasm');
    await copyFile('src/wasm/generated/pptx_wasm_bg.wasm', 'dist/generated/pptx_wasm_bg.wasm');
    await copyFile('src/wasm/generated/pptx_wasm_lean_bg.wasm', 'dist/generated/pptx_wasm_lean_bg.wasm');
  },
}));
