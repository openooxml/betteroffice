import { defineConfig } from 'tsup';

// Ships built ESM, not raw TypeScript: Node refuses to type-strip files under
// `node_modules`, so a source-published package is unimportable from plain
// Node (ERR_UNSUPPORTED_NODE_MODULES_TYPE_STRIPPING) — the exact server-side
// host this package exists to serve.
//
// `dist/` sits beside `assets/` at the package root, so the loader's
// `new URL('../assets/…', import.meta.url)` literals resolve unchanged.
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
