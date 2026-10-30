import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';

test('esbuild deploys the published entry with separate Latin frames and decoder for baseUrl loading', async () => {
  const output = await mkdtemp(join(tmpdir(), 'latin-assets-'));
  const originalFetch = globalThis.fetch;
  try {
    const fontsPackage = new URL(import.meta.resolve('@betteroffice/fonts/package.json'));
    const latinPackage = new URL(import.meta.resolve('@betteroffice/fonts-latin/package.json'));
    const manifest: { faces: { file: string }[] } = JSON.parse(
      await readFile(new URL('./assets/manifest.json', latinPackage), 'utf8'),
    );
    const result = await build({
      stdin: {
        contents: `import { loadBundledFontBytes, resolveBundledFamilyFace } from '@betteroffice/fonts';
export const face = resolveBundledFamilyFace('Carlito', false, false);
export const load = (baseUrl) => loadBundledFontBytes(face, { baseUrl });`,
        resolveDir: import.meta.dir, sourcefile: 'consumer.ts', loader: 'ts',
      },
      entryPoints: [
        ...manifest.faces.map(({ file }) => ({
          in: fileURLToPath(new URL(`./assets/${file}`, latinPackage)),
          out: `fonts/${file.slice(0, -4)}`,
        })),
        {
          in: fileURLToPath(new URL('./dist/decoder/fontpack.wasm', fontsPackage)),
          out: 'fonts/decoder/fontpack',
        },
      ],
      outdir: output, bundle: true, format: 'esm', platform: 'browser',
      metafile: true, outExtension: { '.js': '.mjs' },
      loader: { '.zst': 'copy', '.wasm': 'copy' },
      external: ['@betteroffice/fonts-cjk', '@betteroffice/fonts-intl', '@betteroffice/fonts-latin'],
      define: { 'globalThis.process': 'undefined' },
    });
    expect(Object.keys(result.metafile.inputs).map((file) => resolve(file)))
      .toContain(fileURLToPath(import.meta.resolve('@betteroffice/fonts')));
    const outputs = new Map(Object.entries(result.metafile.outputs).map(([file, info]) => [resolve(file), info]));
    expect([...outputs.keys()].filter((file) => file.endsWith('.zst')).sort()).toEqual(
      manifest.faces.map(({ file }) => join(output, 'fonts', file)).sort(),
    );
    const wasm = join(output, 'fonts/decoder/fontpack.wasm');
    expect([...outputs.keys()].filter((file) => file.endsWith('.wasm'))).toEqual([wasm]);
    expect(outputs.get(wasm)!.bytes).toBeGreaterThan(0);
    const entry = [...outputs].find(([, info]) => info.entryPoint?.endsWith('consumer.ts'));
    expect(entry).toBeDefined();
    const deployment = new URL('https://latin-bundle.example/build/');
    const emitted = new Map([...outputs.keys()].map((file) => [new URL(relative(output, file), deployment).href, file]));
    const requested: string[] = [];
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      const url = new URL(input instanceof Request ? input.url : String(input));
      requested.push(url.href);
      const file = emitted.get(url.href);
      if (!file) throw new Error(`Fetch outside emitted bundle: ${url.href}`);
      return new Response(await readFile(file), file === wasm
        ? { headers: { 'Content-Type': 'application/wasm' } }
        : undefined);
    }) as typeof fetch;
    const bundled = await import(pathToFileURL(entry![0]).href);
    expect(requested).toEqual([]);
    const bytes = await bundled.load(new URL('fonts/', deployment));
    expect(bytes.byteLength).toBe(bundled.face.byteLength);
    expect(new DataView(bytes).getUint32(0)).toBe(0x00010000);
    expect(requested).toEqual([
      new URL('fonts/Carlito-Regular.ttf.zst', deployment).href,
      new URL('fonts/decoder/fontpack.wasm', deployment).href,
    ]);
  } finally {
    globalThis.fetch = originalFetch;
    await rm(output, { recursive: true, force: true });
  }
});
