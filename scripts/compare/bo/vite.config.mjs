import { defineConfig } from 'vite';
import { resolve } from 'node:path';

export default defineConfig({
  root: import.meta.dirname,
  base: './',
  logLevel: 'warn',
  resolve: { dedupe: ['react', 'react-dom'] },
  build: { outDir: resolve(import.meta.dirname, '../dist'), emptyOutDir: false, target: 'es2022',
    ...(process.env.CMP_UNMIN === '1' ? { minify: false, sourcemap: true } : {}),
    rollupOptions: { input: { 'bo-docx': resolve(import.meta.dirname, 'bo-docx.html') } } },
  worker: { format: 'es' },
});
