import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import { fontAssetResponse } from './test-assets';

test('esbuild emits Latin frames separately and fetches only the requested face on demand', async () => {
  const output = await mkdtemp(join(tmpdir(), 'latin-assets-'));
  const originalFetch = globalThis.fetch;
  try {
    const result = await build({
      stdin: {
        contents: `import { loadBundledFontBytes, resolveBundledFamilyFace } from './index';
export const load = () => loadBundledFontBytes(resolveBundledFamilyFace('Carlito', false, false));`,
        resolveDir: import.meta.dir, sourcefile: 'entry.ts', loader: 'ts',
      },
      outdir: output, bundle: true, format: 'esm', platform: 'browser',
      splitting: true, metafile: true, assetNames: '[name]-[hash]',
      outExtension: { '.js': '.mjs' }, loader: { '.zst': 'file' },
      external: ['@betteroffice/fonts-cjk', '@betteroffice/fonts-intl'],
      define: { 'import.meta.url': JSON.stringify('https://latin-bundle.example/entry.mjs') },
      plugins: [{
        name: 'literal-font-urls',
        setup(builder) {
          builder.onResolve({ filter: /^@betteroffice\/fonts-latin$/ }, () => ({
            path: new URL('../../fonts-latin/src/index.ts', import.meta.url).pathname,
          }));
          builder.onLoad({ filter: /fonts-latin\/src\/index\.ts$/ }, async ({ path }) => {
            const imports: string[] = [];
            const source = (await readFile(path, 'utf8')).replace(
              /new URL\('([^']+\.zst)', import\.meta\.url\)/g,
              (_, asset: string) => {
                const name = `asset${imports.length}`;
                imports.push(`import ${name} from ${JSON.stringify(asset)};`);
                return `new URL(${name}, import.meta.url)`;
              },
            );
            return { contents: `${imports.join('\n')}\n${source}`, loader: 'ts', resolveDir: dirname(path) };
          });
        },
      }],
    });
    const outputs = Object.entries(result.metafile.outputs);
    const frame = outputs.find(([file]) => basename(file).startsWith('Carlito-Regular.ttf-') && file.endsWith('.zst'));
    expect(frame).toBeDefined();
    expect(Object.keys(frame![1].inputs).some((file) => file.endsWith('Carlito-Regular.ttf.zst'))).toBe(true);
    const entry = outputs.find(([, info]) => info.entryPoint?.endsWith('entry.ts'))!;
    expect(entry[1].imports.some((entry) => entry.kind === 'dynamic-import')).toBe(true);
    const requested: string[] = [];
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      requested.push(url.href);
      if (url.pathname.endsWith('.wasm')) return fontAssetResponse(input);
      return new Response(await readFile(join(output, basename(url.pathname))));
    }) as typeof fetch;
    const bundled = await import(pathToFileURL(entry[0]).href);
    expect(requested).toEqual([]);
    const bytes = await bundled.load();
    expect(new DataView(bytes).getUint32(0)).toBe(0x00010000);
    expect(requested.filter((url) => url.endsWith('.zst'))).toEqual([
      `https://latin-bundle.example/${basename(frame![0])}`,
    ]);
  } finally {
    globalThis.fetch = originalFetch;
    await rm(output, { recursive: true, force: true });
  }
});
