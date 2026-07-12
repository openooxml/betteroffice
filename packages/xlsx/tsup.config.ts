import { defineConfig } from 'tsup';
import { copyFile, mkdir } from 'node:fs/promises';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    headless: 'src/headless.ts',
  },
  format: ['esm'],
  dts: true,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  onSuccess: async () => {
    await mkdir('dist/generated', { recursive: true });
    await copyFile(
      'src/wasm/generated/xlsx_wasm_bg.wasm',
      'dist/generated/xlsx_wasm_bg.wasm'
    );
  },
});
