import { defineConfig } from 'tsup';
import { rename, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { nativeEnginePlugin } from '../../../scripts/native-engine';
import { packageBoundary } from '../../../scripts/package-boundary';

const source = '../../../packages/pptx';
const dist = '../dist';
const browser = 'throw new Error("@betteroffice/pptx-native runs on Node.js, Bun and Deno; use @betteroffice/pptx in browsers");\n';

export default defineConfig({
  entry: {
    index: `${source}/src/index.ts`,
    worker: `${source}/src/worker.ts`,
    collaboration: `${source}/src/collaboration/index.ts`,
    pptxSessionWorker: `${source}/src/session/worker.ts`,
    'sessionWorker.browser': `${source}/src/session/sessionWorker.browser.ts`,
    'sessionWorker.node': `${source}/src/session/sessionWorker.node.ts`,
  },
  outDir: dist,
  format: ['esm'],
  dts: {
    compilerOptions: {
      paths: { '#session-worker': [fileURLToPath(new URL('../../../packages/pptx/src/session/sessionWorker.node.ts', import.meta.url))] },
    },
  },
  tsconfig: `${source}/tsconfig.json`,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  external: ['@betteroffice/fonts', '@betteroffice/fonts-cjk', '@betteroffice/fonts-intl', '@betteroffice/fonts-latin', '#session-worker'],
  esbuildPlugins: [
    packageBoundary(fileURLToPath(new URL(`${source}/`, import.meta.url))),
    nativeEnginePlugin('pptx'),
  ],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    await rename(join(dist, 'pptxSessionWorker.js'), join(dist, 'pptxSessionWorker.mjs'));
    await writeFile(join(dist, 'package.json'), JSON.stringify({
      type: 'module',
      imports: { '#session-worker': {
        node: './sessionWorker.node.js', default: './sessionWorker.browser.js',
      } },
    }, null, 2) + '\n');
    await writeFile(join(dist, 'browser.js'), browser);
  },
});
