import { defineConfig } from 'tsup';
import { fileURLToPath } from 'node:url';
import { packageBoundary } from '../../scripts/package-boundary';
export default defineConfig({ entry: { index: 'src/index.ts' }, format: ['esm'], dts: true, splitting: true, sourcemap: false, clean: true, treeshake: true, minify: true, external: ['react', 'react-dom', '@betteroffice/vsdx', '@betteroffice/vsdx-i18n'], esbuildPlugins: [packageBoundary(fileURLToPath(new URL('.', import.meta.url)))] });
