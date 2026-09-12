import { defineConfig } from 'tsup';
export default defineConfig({ entry: { index: 'src/index.ts' }, format: ['esm'], dts: true, splitting: true, sourcemap: false, clean: true, treeshake: true, minify: true, external: ['react', 'react-dom', '@betteroffice/vsdx', '@betteroffice/vsdx-i18n'] });
