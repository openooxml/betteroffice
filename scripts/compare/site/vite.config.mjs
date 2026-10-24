import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

export default defineConfig({
  root: import.meta.dirname,
  base: './',
  logLevel: 'warn',
  plugins: [react()],
  resolve: { dedupe: ['react', 'react-dom'] },
  build: { outDir: resolve(import.meta.dirname, '../dist'), emptyOutDir: false, target: 'es2022',
    rollupOptions: { input: { 'dxe-docx': resolve(import.meta.dirname, 'dxe-docx.html') } } },
  worker: { format: 'es' },
});
