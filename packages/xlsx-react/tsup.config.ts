import { defineConfig, type Options } from 'tsup';
import { fileURLToPath } from 'node:url';
import { packageBoundary } from '../../scripts/package-boundary';

export const library: Options = {
  entry: {
    index: 'src/index.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  external: ['react', 'react-dom', '@betteroffice/xlsx', '@betteroffice/fonts'],
  esbuildPlugins: [packageBoundary(fileURLToPath(new URL('.', import.meta.url)))],
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
};

export default defineConfig(library);
