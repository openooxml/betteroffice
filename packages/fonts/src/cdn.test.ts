import { afterEach, describe, expect, test } from 'bun:test';
import { createFontProvider } from './cdn';
import { createFontProvider as createBundledFontProvider } from './index';

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
});

function intercept(files: Record<string, string>, requested: string[]) {
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    requested.push(url);
    const file = files[url];
    if (!file) return new Response('missing', { status: 404 });
    return new Response(await Bun.file(file).arrayBuffer());
  }) as typeof fetch;
}

describe('CDN font provider', () => {
  test('keeps existing bundled fetch behavior separate from CDN request policy', async () => {
    const requests: (RequestInit | undefined)[] = [];
    globalThis.fetch = (async (_input: RequestInfo | URL, options?: RequestInit) => {
      requests.push(options);
      return new Response(await Bun.file(new URL('../assets/Carlito-Regular.ttf', import.meta.url)).arrayBuffer());
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
    const latin =
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.1.0/assets/Carlito-Regular.ttf';
    const japanese =
      'https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@0.1.0/assets/NotoSansJP-Regular.otf';
    intercept(
      {
        [latin]: new URL('../assets/Carlito-Regular.ttf', import.meta.url)
          .pathname,
        [japanese]: new URL(
          '../../fonts-cjk/assets/NotoSansJP-Regular.otf',
          import.meta.url,
        ).pathname,
      },
      requested,
    );
    const provider = createFontProvider();
    expect(requested).toEqual([]);
    const load = provider.resolve('Calibri', false, false)!;
    const [first, second] = await Promise.all([load(), load()]);
    expect(first).toBe(second);
    expect(new DataView(first).getUint32(0)).toBe(0x00010000);
    const cjk = await provider.resolve('MS Mincho', false, false)!();
    expect(new DataView(cjk).getUint32(0)).toBe(0x4f54544f);
    expect(requested).toEqual([latin, japanese]);
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
      calls++;
      requests.push(String(input));
      expect(options?.signal).toBeInstanceOf(AbortSignal);
      expect(options?.credentials).toBe('omit');
      expect(options?.referrerPolicy).toBe('no-referrer');
      if (calls === 1) return new Response('unavailable', { status: 503 });
      if (calls === 2) return new Response(new Uint8Array(4));
      return new Response(
        await Bun.file(
          new URL('../assets/Carlito-Regular.ttf', import.meta.url),
        ).arrayBuffer(),
      );
    }) as typeof fetch;
    const load = provider.resolve('Calibri', false, false)!;
    await expect(load()).rejects.toThrow('HTTP 503');
    await expect(load()).rejects.toThrow('expected 628032');
    expect((await load()).byteLength).toBe(628032);
    expect(requests).toEqual(
      Array(3).fill('https://retry.example/latin/Carlito-Regular.ttf'),
    );
  });

  test('browser entry bundles without binary asset or CJK package imports', async () => {
    const build = await Bun.build({
      entrypoints: [new URL('./cdn.ts', import.meta.url).pathname],
      target: 'browser',
      minify: true,
      plugins: [
        {
          name: 'reject-font-assets',
          setup(builder) {
            builder.onResolve(
              { filter: /\.(ttf|otf)$|^@betteroffice\/fonts-cjk/ },
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
    expect(build.outputs[0]!.size).toBeLessThan(20_000);
  });
});
