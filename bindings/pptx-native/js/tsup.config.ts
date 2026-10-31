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
  },
  outDir: dist,
  format: ['esm'],
  dts: true,
  tsconfig: `${source}/tsconfig.json`,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  external: ['@betteroffice/fonts'],
  esbuildPlugins: [
    packageBoundary(fileURLToPath(new URL(`${source}/`, import.meta.url))),
    nativeEnginePlugin('pptx'),
  ],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    await rename(join(dist, 'pptxSessionWorker.js'), join(dist, 'pptxSessionWorker.mjs'));
    await writeFile(join(dist, 'package.json'), '{\n  "type": "module"\n}\n');
    await writeFile(join(dist, 'browser.js'), browser);
  },
});
