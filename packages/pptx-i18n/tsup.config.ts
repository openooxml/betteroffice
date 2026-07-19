import { defineConfig } from 'tsup';
import { readLocaleCodes } from './locale-files.ts';

const localeCodes = readLocaleCodes(import.meta.dirname);

export default defineConfig({
  entry: ['src/index.ts', ...localeCodes.map((code) => `src/${code}.ts`)],
  format: ['esm'],
  dts: { resolve: true },
  splitting: false,
  sourcemap: false,
  clean: true,
  minify: false,
  esbuildOptions(options) {
    options.charset = 'utf8';
  },
});
