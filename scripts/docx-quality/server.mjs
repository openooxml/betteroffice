import { createServer } from 'vite';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const packageRoot = process.env.QUALITY_PACKAGE_ROOT;
const reactRoot = process.env.QUALITY_REACT_ROOT;
const aliases = [];
for (const [name, override] of [
  ['docx', packageRoot],
  ['docx-react', reactRoot],
]) {
  if (!override) continue;
  const manifest = JSON.parse(readFileSync(resolve(override, 'package.json')));
  for (const [key, value] of Object.entries(manifest.exports)) {
    const target = typeof value === 'string' ? value : (value.import ?? value.default);
    if (typeof target !== 'string') continue;
    aliases.push({
      find: new RegExp(`^@betteroffice/${name}${key === '.' ? '' : key.slice(1)}$`),
      replacement: resolve(override, target),
    });
  }
}
const server = await createServer({
  configFile: false,
  cacheDir: resolve(
    `.source/docx-quality/vite-cache-${process.env.QUALITY_PORT ?? 4178}`
  ),
  root: resolve('scripts/docx-quality'),
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
      'react',
      'react-dom',
      'react-dom/client',
      'react/jsx-runtime',
      'react/jsx-dev-runtime',
      'clsx',
      'sonner',
      '@radix-ui/react-select',
    ],
    exclude: ['@betteroffice/docx', '@betteroffice/docx-react'],
  },
  esbuild: { jsx: 'automatic' },
});
await server.listen();
console.log(server.resolvedUrls.local[0]);
