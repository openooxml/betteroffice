import { fileURLToPath } from 'node:url';
import { defineConfig } from 'tsup';
import { rename, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { nativeEnginePlugin } from '../../../scripts/native-engine';

const source = '../../../packages/xlsx';
const dist = '../dist';
const browser = 'throw new Error("@betteroffice/xlsx-native runs on Node.js, Bun and Deno; use @betteroffice/xlsx in browsers");\n';

export default defineConfig({
  entry: {
    index: `${source}/src/index.ts`,
    'wasm-preflight': `${source}/src/wasm/preflight.ts`,
    headless: `${source}/src/headless.ts`,
    worker: `${source}/src/worker.ts`,
    collaboration: `${source}/src/collaboration/index.ts`,
    xlsxSessionWorker: `${source}/src/session/worker.ts`,
    'xlsxSessionWorker.node': `${source}/src/session/worker.node.ts`,
    'sessionWorker.node': `${source}/src/session/sessionWorker.node.ts`,
    'sessionWorker.missing': `${source}/src/session/sessionWorker.missing.ts`,
  },
  outDir: dist,
  format: ['esm'],
  dts: {
    compilerOptions: {
      paths: { '#session-worker': [fileURLToPath(new URL('../../../packages/xlsx/src/session/sessionWorker.node.ts', import.meta.url))] },
    },
  },
  tsconfig: `${source}/tsconfig.json`,
  splitting: true,
  external: ['@betteroffice/fonts-cjk', '@betteroffice/fonts-intl', '@betteroffice/fonts-latin', '#session-worker'],
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  esbuildPlugins: [nativeEnginePlugin('xlsx')],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    await rename(join(dist, 'xlsxSessionWorker.js'), join(dist, 'xlsxSessionWorker.mjs'));
    await rename(join(dist, 'xlsxSessionWorker.node.js'), join(dist, 'xlsxSessionWorker.node.mjs'));
    await writeFile(join(dist, 'package.json'), JSON.stringify({
      type: 'module',
      imports: {
        '#session-worker': { node: './sessionWorker.node.js', default: './sessionWorker.missing.js' },
      },
    }, null, 2) + '\n');
    await writeFile(join(dist, 'browser.js'), browser);
  },
});
