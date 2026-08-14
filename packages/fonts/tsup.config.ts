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
  // The Word-family alias table is full of CJK; esbuild would otherwise
  // `\uHHHH`-escape every entry.
  esbuildOptions(options) {
    options.charset = 'utf8';
  },
});
