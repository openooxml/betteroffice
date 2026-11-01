import { fontAssetResponse } from './test-assets';
import { version as fontsLatinVersion } from '../../fonts-latin/package.json';
import { afterEach, describe, expect, test } from 'bun:test';
import { version as fontsVersion } from '../package.json';
import { version as fontsCjkVersion } from '../../fonts-cjk/package.json';
import { version as fontsIntlVersion } from '../../fonts-intl/package.json';
import { createFontProvider } from './cdn';
import { createFontProvider as createBundledFontProvider } from './index';

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
});

function intercept(files: Record<string, string>, requested: string[]) {
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    if (!url.endsWith('.wasm')) requested.push(url);
    if (url.endsWith('.wasm')) return fontAssetResponse(input);
    const file = files[url];
    if (!file) return new Response('missing', { status: 404 });
    return new Response(await Bun.file(file).arrayBuffer());
  }) as typeof fetch;
}

describe('CDN font provider', () => {
  test('keeps existing bundled fetch behavior separate from CDN request policy', async () => {
    const requests: (RequestInit | undefined)[] = [];
    globalThis.fetch = (async (_input: RequestInfo | URL, options?: RequestInit) => {
      if (!String(_input).endsWith('.wasm')) requests.push(options);
      return fontAssetResponse(_input);
    }) as typeof fetch;
    const baseUrl = 'https://legacy.example/fonts/';
    await createBundledFontProvider({ baseUrl }).resolve('Calibri', false, false)!();
    await createFontProvider({ baseUrl }).resolve('Calibri', false, false)!();
    expect(requests).toHaveLength(2);
    expect(requests[0]).toBeUndefined();
    expect(requests[1]?.signal).toBeInstanceOf(AbortSignal);
    expect(requests[1]?.credentials).toBe('omit');
    expect(requests[1]?.referrerPolicy).toBe('no-referrer');
  });

  test('lazily loads Latin and Japanese faces from separately pinned packages', async () => {
    const requested: string[] = [];
    const latin = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-latin@${fontsLatinVersion}/assets/Carlito-Regular.ttf.zst`;
    const japanese = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@${fontsCjkVersion}/assets/NotoSansJP-Regular.otf.zst`;
    intercept(
      {
        [latin]: new URL('../../fonts-latin/assets/Carlito-Regular.ttf.zst', import.meta.url)
          .pathname,
        [japanese]: new URL(
          '../../fonts-cjk/assets/NotoSansJP-Regular.otf.zst',
          import.meta.url,
        ).pathname,
      },
      requested,
    );
    const provider = createFontProvider();
    expect(requested).toEqual([]);
    const load = provider.resolve('Calibri', false, false)!;
    const [first, second] = await Promise.all([load(), load()]);
    expect(first).not.toBe(second);
    expect(first).toEqual(second);
    expect(new DataView(first).getUint32(0)).toBe(0x00010000);
    const cjk = await provider.resolve('MS Mincho', false, false)!();
    expect(new DataView(cjk).getUint32(0)).toBe(0x4f54544f);
    expect(requested).toEqual([latin, japanese]);
  });

  test('Selawik CDN assets load only when a Segoe UI loader is invoked', async () => {
    const requested: string[] = [];
    const baseUrl = 'https://selawik.example/fonts/';
    intercept({
      [`${baseUrl}Carlito-Regular.ttf.zst`]: new URL('../../fonts-latin/assets/Carlito-Regular.ttf.zst', import.meta.url).pathname,
      [`${baseUrl}Selawik-Regular.ttf.zst`]: new URL('../../fonts-latin/assets/Selawik-Regular.ttf.zst', import.meta.url).pathname,
      [`${baseUrl}Selawik-Semilight.ttf.zst`]: new URL('../../fonts-latin/assets/Selawik-Semilight.ttf.zst', import.meta.url).pathname,
    }, requested);
    const provider = createFontProvider({ baseUrl });
    const load = provider.resolve('Segoe UI Semilight', true, false)!;
    expect(requested).toEqual([]);
    await provider.resolveLastResort('Google Sans', false, false)();
    expect(requested).toEqual([`${baseUrl}Carlito-Regular.ttf.zst`]);
    expect((await load()).byteLength).toBeGreaterThan(0);
    expect(requested.slice(1).sort()).toEqual([`${baseUrl}Selawik-Regular.ttf.zst`, `${baseUrl}Selawik-Semilight.ttf.zst`].sort());
  });

  test('custom origins remain isolated and failed downloads can retry', async () => {
    const provider = createFontProvider({
      baseUrl: 'https://retry.example/latin',
      cjkBaseUrl: 'https://retry.example/cjk',
    });
    let calls = 0;
    const requests: string[] = [];
    globalThis.fetch = (async (
      input: RequestInfo | URL,
      options?: RequestInit,
    ) => {
      if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
      calls++;
      requests.push(String(input));
      expect(options?.signal).toBeInstanceOf(AbortSignal);
      expect(options?.credentials).toBe('omit');
      expect(options?.referrerPolicy).toBe('no-referrer');
      if (calls === 1) return new Response('unavailable', { status: 503 });
      if (calls === 2) return new Response(new Uint8Array(4));
      return new Response(
        await Bun.file(
          new URL('../../fonts-latin/assets/Carlito-Regular.ttf.zst', import.meta.url),
        ).arrayBuffer(),
      );
    }) as typeof fetch;
    const load = provider.resolve('Calibri', false, false)!;
    await expect(load()).rejects.toThrow('HTTP 503');
    await expect(load()).rejects.toThrow('expected');
    expect((await load()).byteLength).toBeGreaterThan(0);
    expect(requests).toEqual(
      Array(3).fill('https://retry.example/latin/Carlito-Regular.ttf.zst'),
    );
  });

  test('browser entry bundles without binary asset or add-on package imports', async () => {
    const build = await Bun.build({
      entrypoints: [new URL('./cdn.ts', import.meta.url).pathname],
      target: 'browser',
      minify: true,
      define: { 'import.meta.url': JSON.stringify('https://consumer.example/dist/cdn.js') },
      plugins: [
        {
          name: 'reject-font-assets',
          setup(builder) {
            builder.onResolve(
              { filter: /\.(ttf|otf)(\.zst)?$|^@betteroffice\/fonts-(cjk|intl|latin)/ },
              (args) => {
                throw new Error(`CDN entry imported ${args.path}`);
              },
            );
          },
        },
      ],
    });
    expect(build.success).toBe(true);
    expect(build.outputs).toHaveLength(1);
    expect(build.outputs[0]!.size).toBeLessThan(70_000);
    const entry = await import(`data:text/javascript;base64,${Buffer.from(await build.outputs[0]!.text()).toString('base64')}`);
    const host = globalThis.process as unknown as { getBuiltinModule?: (id: string) => unknown };
    const original = host.getBuiltinModule;
    host.getBuiltinModule = (id: string) => id === 'node:zlib' ? {} : original?.call(host, id);
    const frame = await Bun.file(new URL('../../fonts-intl/assets/NotoSansThai-Regular.ttf.zst', import.meta.url)).arrayBuffer();
    try {
      for (const [index, baseUrl] of [undefined, 'https://custom.example/fonts/', 'http://127.0.0.1:4178/fonts/'].entries()) {
        const pack = `${baseUrl ?? `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-intl@${fontsIntlVersion}/assets/`}NotoSansThai-Regular.ttf.zst`;
        const decoder = baseUrl === undefined
          ? `https://cdn.jsdelivr.net/npm/@betteroffice/fonts@${fontsVersion}/dist/decoder/fontpack.wasm`
          : `${baseUrl}decoder/fontpack.wasm`;
        const requested: string[] = [];
        globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
          const url = String(input);
          requested.push(url);
          expect(options?.credentials).toBe('omit');
          expect(options?.referrerPolicy).toBe('no-referrer');
          expect(options?.signal).toBeInstanceOf(AbortSignal);
          return url === pack
            ? new Response(frame)
            : new Response('missing decoder', { status: 404 });
        }) as typeof fetch;
        const provider = entry.createFontProvider({ baseUrl });
        const bundled = 'https://consumer.example/dist/decoder/fontpack.wasm';
        await expect(provider.resolveScriptFallback('thai', false, false)())
          .rejects.toThrow(`Font decoder wasm not found: ${decoder} (HTTP 404), ${bundled} (HTTP 404)`);
        expect(requested).toEqual(index === 0 ? [pack, decoder, bundled] : [pack, decoder]);
      }
    } finally {
      host.getBuiltinModule = original;
    }
  });
});
