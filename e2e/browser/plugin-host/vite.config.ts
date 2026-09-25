import { createRequire } from 'node:module';
import { resolve } from 'node:path';
import { defineConfig } from 'vite';

const root = resolve(import.meta.dirname, '../../..');
const fromReact = createRequire(resolve(root, 'packages/docx-react/package.json'));
const tailwind = fromReact('tailwindcss');
const autoprefixer = fromReact('autoprefixer');

/** Serves the plugin-host harnesses against the package sources. */
export default defineConfig({
  root: import.meta.dirname,
  resolve: {
    alias: [
      {
        find: /^@betteroffice\/docx-react$/,
        replacement: resolve(root, 'packages/docx-react/src/index.ts'),
      },
      {
        find: /^@betteroffice\/docx-i18n$/,
        replacement: resolve(root, 'packages/docx-i18n/src/index.ts'),
      },
      { find: /^@betteroffice\/docx$/, replacement: resolve(root, 'packages/docx/src/core.ts') },
      { find: /^@betteroffice\/docx\/(.*)$/, replacement: resolve(root, 'packages/docx/src/$1') },
      {
        find: /^@betteroffice\/xlsx-react$/,
        replacement: resolve(root, 'packages/xlsx-react/src/index.ts'),
      },
      {
        find: /^@betteroffice\/xlsx-i18n$/,
        replacement: resolve(root, 'packages/xlsx-i18n/src/index.ts'),
      },
      { find: /^@betteroffice\/xlsx$/, replacement: resolve(root, 'packages/xlsx/src/index.ts') },
      { find: /^@betteroffice\/xlsx\/(.*)$/, replacement: resolve(root, 'packages/xlsx/src/$1') },
    ],
  },
  css: {
    postcss: {
      plugins: [
        tailwind({ config: resolve(root, 'packages/docx-react/tailwind.config.js') }),
        autoprefixer(),
      ],
    },
  },
  server: { fs: { allow: [root] } },
  worker: { format: 'es' },
});
