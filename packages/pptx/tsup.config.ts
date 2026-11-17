import { copyFile, mkdir, rename } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'tsup';
import { packageBoundary } from '../../scripts/package-boundary';

export default defineConfig((options) => ({
  entry: {
    index: 'src/index.ts',
    'wasm-preflight': 'src/wasm/preflight.ts',
    worker: 'src/worker.ts',
    collaboration: 'src/collaboration/index.ts',
    pptxSessionWorker: 'src/session/worker.ts',
    'pptxSessionWorker.node': 'src/session/worker.node.ts',
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
  esbuildPlugins: [packageBoundary(fileURLToPath(new URL('.', import.meta.url)))],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    const out = options.outDir ?? 'dist';
    await rename(join(out, 'pptxSessionWorker.js'), join(out, 'pptxSessionWorker.mjs'));
    await rename(join(out, 'pptxSessionWorker.node.js'), join(out, 'pptxSessionWorker.node.mjs'));
    await mkdir('dist/generated', { recursive: true });
    await mkdir(join(out, 'generated/preflight'), { recursive: true });
    await copyFile('src/wasm/generated/preflight/pptx_preflight_bg.wasm', join(out, 'generated/preflight/pptx_preflight_bg.wasm'));
    await copyFile('src/wasm/generated/metafile_svg_wasm_bg.wasm', 'dist/generated/metafile_svg_wasm_bg.wasm');
    await copyFile('src/wasm/generated/pptx_wasm_bg.wasm', 'dist/generated/pptx_wasm_bg.wasm');
    await copyFile('src/wasm/generated/pptx_wasm_lean_bg.wasm', 'dist/generated/pptx_wasm_lean_bg.wasm');
  },
}));
