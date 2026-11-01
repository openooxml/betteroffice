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
    headless: `${source}/src/headless.ts`,
    worker: `${source}/src/worker.ts`,
    collaboration: `${source}/src/collaboration/index.ts`,
    xlsxSessionWorker: `${source}/src/session/worker.ts`,
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
  esbuildPlugins: [nativeEnginePlugin('xlsx')],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
  onSuccess: async () => {
    await rename(join(dist, 'xlsxSessionWorker.js'), join(dist, 'xlsxSessionWorker.mjs'));
    await writeFile(join(dist, 'package.json'), '{\n  "type": "module"\n}\n');
    await writeFile(join(dist, 'browser.js'), browser);
  },
});
