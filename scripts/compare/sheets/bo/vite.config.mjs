import { defineConfig } from 'vite';
import { resolve } from 'node:path';

export default defineConfig({
  root: import.meta.dirname,
  base: './',
  logLevel: 'warn',
  resolve: { dedupe: ['react', 'react-dom'] },
  build: { outDir: resolve(import.meta.dirname, '../dist'), emptyOutDir: false, target: 'es2022',
    rollupOptions: { input: { 'bo-xlsx': resolve(import.meta.dirname, 'bo-xlsx.html') } } },
  worker: { format: 'es' },
});
