import { defineConfig } from 'tsup';

// Plain Node cannot type-strip packages; dist remains beside assets so
// import.meta.url paths stay valid.
export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  dts: { resolve: true },
  splitting: false,
  sourcemap: false,
  clean: true,
  minify: false,
});
