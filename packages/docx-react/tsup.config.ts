import { defineConfig, type Options } from 'tsup';
import { fileURLToPath } from 'node:url';
import { packageBoundary } from '../../scripts/package-boundary';

const library: Options = {
  entry: {
    index: 'src/index.ts',
    'peerless-harness': 'src/components/DocxEditor/peerless/browserHarness.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  define: {
    'globalThis.__DOCX_PEERLESS_HARNESS__': JSON.stringify(process.env.DOCX_PEERLESS_HARNESS === '1'),
  },
  external: ['react', 'react-dom'],
  esbuildPlugins: [packageBoundary(fileURLToPath(new URL('.', import.meta.url)))],
  injectStyle: false,
  esbuildOptions(options) {
    options.chunkNames = 'chunk-[hash]';
  },
};

/**
 * The `./shortcuts` subpath builds on its own so each entry's declarations stand alone; a shared
 * declaration chunk would leave private imports in index.d.ts.
 */
export default defineConfig([
  library,
  { ...library, entry: { shortcuts: 'src/shortcuts.ts' }, clean: false },
]);
