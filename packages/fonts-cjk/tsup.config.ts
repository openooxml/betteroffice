import { defineConfig } from 'tsup';

// See the sibling config in `packages/fonts`: built ESM so plain Node can
// import it from under `node_modules`. `dist/` sits beside `assets/`, so the
// `new URL('../assets/…', import.meta.url)` literals resolve unchanged.
export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  dts: { resolve: true },
  splitting: false,
  sourcemap: false,
  clean: true,
  minify: false,
});
