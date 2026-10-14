import { createServer } from 'vite';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { qualityRendererPlugin } from '../office-quality/renderer.mjs';

const packageRoot = process.env.QUALITY_PACKAGE_ROOT;
const reactRoot = process.env.QUALITY_REACT_ROOT;
const format = process.env.QUALITY_FORMAT ?? 'docx';
if (!['docx', 'pptx', 'xlsx', 'vsdx'].includes(format)) throw new Error('Invalid format');
const aliases = [];
for (const [name, override] of [
  [format, packageRoot],
  ['docx-react', reactRoot],
]) {
  if (!override) continue;
  const manifest = JSON.parse(readFileSync(resolve(override, 'package.json')));
  for (const [key, value] of Object.entries(manifest.exports)) {
    const target = typeof value === 'string' ? value : value.import ?? value.default;
    if (typeof target !== 'string') continue;
    aliases.push({
      find: new RegExp(`^@betteroffice/${name}${key === '.' ? '' : key.slice(1)}$`),
      replacement: resolve(override, target),
    });
  }
}
// Fonts this checkout has not published yet cannot come from the pinned CDN,
// so an explicit assets directory is served from the viewer's own origin.
const fontAssets = process.env.QUALITY_FONT_ASSETS?.trim();
const server = await createServer({
  configFile: false,
  define: {
    __QUALITY_FONT_BASE__: JSON.stringify(
      fontAssets ? `/@fs${resolve(fontAssets)}/` : null
    ),
    __QUALITY_FONT_BASE_CJK__: JSON.stringify(
      fontAssets
        ? `/@fs${resolve(process.env.QUALITY_FONT_ASSETS_CJK ?? 'packages/fonts-cjk/assets')}/`
        : null
    ),
  },
  plugins: format === 'docx' ? [] : [qualityRendererPlugin(format)],
  cacheDir: resolve(
    `.source/docx-quality/vite-cache-${process.env.QUALITY_PORT ?? 4178}`
  ),
  root: resolve(format === 'docx' ? 'scripts/docx-quality' : 'scripts/office-quality'),
  resolve: {
    alias: aliases,
    dedupe: [
      'react',
      'react-dom',
      'clsx',
      'sonner',
      '@radix-ui/react-select',
      '@betteroffice/docx-i18n',
    ],
  },
  server: {
    host: '127.0.0.1',
    port: Number(process.env.QUALITY_PORT ?? 4178),
    strictPort: true,
    watch: null,
    hmr: false,
    fs: {
      allow: [
        process.cwd(),
        ...(packageRoot ? [packageRoot] : []),
        ...(reactRoot ? [reactRoot] : []),
      ],
    },
  },
  optimizeDeps: {
    noDiscovery: true,
    include: [
      ...(format === 'docx' ? [] : ['jszip']),
      'react',
      'react-dom',
      'react-dom/client',
      'react/jsx-runtime',
      'react/jsx-dev-runtime',
      'clsx',
      'sonner',
      '@radix-ui/react-select',
    ],
    exclude: [
      '@betteroffice/docx',
      '@betteroffice/docx-react',
      '@betteroffice/pptx',
      '@betteroffice/xlsx',
      '@betteroffice/vsdx',
    ],
  },
  esbuild: { jsx: 'automatic' },
});
await server.listen();
console.log(server.resolvedUrls.local[0]);
