import { afterEach, describe, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  DEFAULT_REMOTE_FONT_PROVIDER, WORD_FAMILY_SUBSTITUTES, fetchRemoteFontRules, loadRemoteFontFaces,
  parseRemoteFontCss, remoteFontStylesheetUrl, resolveWordFamilySubstitute,
} from './remote';

const BUNNY_CSS = `/* latin-ext */
@font-face {
  font-family: 'Open Sans';
  font-style: normal;
  font-weight: 400;
  font-stretch: 100%;
  src: url(https://fonts.bunny.net/open-sans/files/open-sans-latin-ext-400-normal.woff2) format('woff2'), url(https://fonts.bunny.net/open-sans/files/open-sans-latin-ext-400-normal.woff) format('woff');
  unicode-range: U+0100-02AF,U+0304,U+0308;
}
/* latin */
@font-face {
  font-family: 'Open Sans';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff2) format('woff2'), url(https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff) format('woff');
  unicode-range: U+0000-00FF,U+0131,U+0152-0153;
}
@font-face {
  font-family: 'Open Sans';
  font-style: italic;
  font-weight: 700;
  src: url(http://tracker.example/x.woff2) format('woff2');
}`;

describe('substitution table', () => {
  test('keys are lowercase, trimmed and name a generic family', () => {
    for (const [key, entry] of Object.entries(WORD_FAMILY_SUBSTITUTES)) {
      expect(key).toBe(key.trim().toLowerCase());
      expect(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy']).toContain(entry.generic);
      if (entry.remote !== null) expect(() => remoteFontStylesheetUrl(entry.remote!)).not.toThrow();
    }
  });

  test('resolves Office families case-insensitively', () => {
    expect(resolveWordFamilySubstitute('  CALIBRI ')).toEqual({ remote: 'Carlito', generic: 'sans-serif' });
    expect(resolveWordFamilySubstitute('Courier New')).toEqual({ remote: 'Cousine', generic: 'monospace' });
    expect(resolveWordFamilySubstitute('ＭＳ 明朝')).toEqual({ remote: 'Noto Serif JP', generic: 'serif' });
    expect(resolveWordFamilySubstitute('toString')).toBeUndefined();
    expect(resolveWordFamilySubstitute('not-a-family')).toBeUndefined();
  });

  test('is exported to the shared catalog unchanged', () => {
    const catalog = JSON.parse(readFileSync(new URL('../../../crates/betteroffice-fonts/data/catalog.json', import.meta.url), 'utf8'));
    expect(catalog.wordFamilySubstitutes).toEqual(WORD_FAMILY_SUBSTITUTES);
  });

  test('stays out of the main entry', () => {
    const index = readFileSync(new URL('./index.ts', import.meta.url), 'utf8');
    expect(index).not.toMatch(/from\s+['"]\.\/(?:substitutes|remote)['"]|import\(\s*['"]\.\/(?:substitutes|remote)['"]/);
  });
});

describe('provider stylesheet URLs', () => {
  const faces = [{ weight: 700, style: 'italic' as const }, { weight: 400, style: 'normal' as const }];

  test('defaults to Bunny Fonts', () => {
    expect(DEFAULT_REMOTE_FONT_PROVIDER).toBe('bunny');
    expect(remoteFontStylesheetUrl('Open Sans', faces))
      .toBe('https://fonts.bunny.net/css?family=Open+Sans:400,700i&display=swap');
  });

  test('builds Google and custom endpoints and is null when off', () => {
    expect(remoteFontStylesheetUrl('Noto Sans JP', faces, 'google'))
      .toBe('https://fonts.googleapis.com/css?family=Noto+Sans+JP:400,700i&display=swap');
    expect(remoteFontStylesheetUrl('Carlito', faces, { cssUrl: 'https://fonts.example.eu/css?x=1#y' }))
      .toBe('https://fonts.example.eu/css?family=Carlito:400,700i&display=swap');
    expect(remoteFontStylesheetUrl('Carlito', faces, 'off')).toBeNull();
  });

  test('rejects plain-http endpoints and family names that could alter the query', () => {
    expect(() => remoteFontStylesheetUrl('Carlito', faces, { cssUrl: 'http://fonts.example.eu/css' })).toThrow(TypeError);
    expect(remoteFontStylesheetUrl('Carlito', faces, { cssUrl: 'http://localhost:8080/css' }))
      .toBe('http://localhost:8080/css?family=Carlito:400,700i&display=swap');
    for (const family of ['Open Sans&text=x', 'a:700', '', 'x|y', '../etc']) {
      expect(() => remoteFontStylesheetUrl(family, faces)).toThrow(TypeError);
    }
  });
});

describe('provider stylesheets', () => {
  afterEach(() => { mock.restore(); });

  test('parses @font-face rules and drops insecure sources', () => {
    const rules = parseRemoteFontCss(BUNNY_CSS, 'https://fonts.bunny.net/css?family=Open+Sans:400');
    expect(rules).toHaveLength(2);
    expect(rules[1]).toEqual({
      family: 'Open Sans', weight: '400', style: 'normal', unicodeRange: 'U+0000-00FF,U+0131,U+0152-0153',
      sources: [
        { url: 'https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff2', format: 'woff2' },
        { url: 'https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff', format: 'woff' },
      ],
    });
  });

  test('fetches without credentials or referrer, caches successes and retries failures', async () => {
    let calls = 0;
    const fetch = mock(async (_url: string | URL | Request, init?: RequestInit) => {
      calls += 1;
      expect(init).toMatchObject({ credentials: 'omit', referrerPolicy: 'no-referrer' });
      return calls === 1 ? new Response('', { status: 503 }) : new Response(BUNNY_CSS);
    });
    const options = { fetch: fetch as unknown as typeof globalThis.fetch };
    const faces = [{ weight: 400, style: 'normal' as const }];
    await expect(fetchRemoteFontRules('Fira Sans', faces, options)).rejects.toThrow('503');
    expect(await fetchRemoteFontRules('Fira Sans', faces, options)).toHaveLength(2);
    expect(await fetchRemoteFontRules('Fira Sans', faces, options)).toHaveLength(2);
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(String(fetch.mock.calls[0]![0])).toBe('https://fonts.bunny.net/css?family=Fira+Sans:400&display=swap');
    expect(await fetchRemoteFontRules('Fira Sans', faces, { ...options, provider: 'off' })).toEqual([]);
  });

  test('a stylesheet that never arrives fails at the timeout and is retried later', async () => {
    const fetch = mock(async (_url: string | URL | Request, init?: RequestInit) =>
      new Promise<Response>((_, reject) => { init!.signal!.addEventListener('abort', () => reject(init!.signal!.reason)); }));
    const started = performance.now();
    await expect(fetchRemoteFontRules('Questrial', [{ weight: 400, style: 'normal' }],
      { fetch: fetch as unknown as typeof globalThis.fetch, timeoutMs: 30 })).rejects.toThrow();
    expect(performance.now() - started).toBeLessThan(2000);
    await expect(fetchRemoteFontRules('Questrial', [{ weight: 400, style: 'normal' }],
      { fetch: fetch as unknown as typeof globalThis.fetch, timeoutMs: 30 })).rejects.toThrow();
    expect(fetch).toHaveBeenCalledTimes(2);
  });

  test('registers every slice of the look-alike from bytes fetched without credentials or referrer', async () => {
    const created: { family: string; source: ArrayBuffer; descriptors: FontFaceDescriptors; loads: number }[] = [];
    class FakeFontFace {
      readonly record: (typeof created)[number];
      constructor(family: string, source: ArrayBuffer, descriptors: FontFaceDescriptors) {
        this.record = { family, source, descriptors, loads: 0 };
        created.push(this.record);
      }
      async load() { this.record.loads += 1; return this; }
    }
    const original = globalThis.FontFace;
    globalThis.FontFace = FakeFontFace as unknown as typeof FontFace;
    try {
      const added: unknown[] = [];
      const fontSet = { add: (face: unknown) => { added.push(face); } } as unknown as FontFaceSet;
      const fetch = mock(async (url: string | URL | Request, init?: RequestInit) => {
        expect(init).toMatchObject({ credentials: 'omit', referrerPolicy: 'no-referrer' });
        return String(url).includes('/css?') ? new Response(BUNNY_CSS) : new Response(new Uint8Array([1, 2, 3]));
      });
      const faces = await loadRemoteFontFaces('Tahoma', 'Tahoma', { weight: 400, style: 'normal' }, fontSet,
        { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.eu/css' } });
      expect(fetch.mock.calls.map((call) => String(call[0]))).toEqual([
        'https://fonts.example.eu/css?family=Open+Sans:400&display=swap',
        'https://fonts.bunny.net/open-sans/files/open-sans-latin-ext-400-normal.woff2',
        'https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff2',
      ]);
      expect(faces).toHaveLength(2);
      expect(added).toEqual(faces);
      expect(created.map((face) => [face.family, face.descriptors.unicodeRange, face.loads, face.source.byteLength])).toEqual([
        ['Tahoma', 'U+0100-02AF,U+0304,U+0308', 1, 3],
        ['Tahoma', 'U+0000-00FF,U+0131,U+0152-0153', 1, 3],
      ]);
      const regularOnly = `@font-face { font-family: 'Anton'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/anton/files/anton-latin-400-normal.woff2) format('woff2'); unicode-range: U+0000-00FF; }`;
      const sparse = mock(async (url: string | URL | Request) => String(url).includes('/css?')
        ? new Response(String(url).includes(':700') ? '' : regularOnly) : new Response(new Uint8Array([4])));
      const bold = await loadRemoteFontFaces('Impact', 'Impact', { weight: 700, style: 'normal' }, fontSet,
        { fetch: sparse as unknown as typeof globalThis.fetch });
      expect(sparse.mock.calls.map((call) => String(call[0]))).toEqual([
        'https://fonts.bunny.net/css?family=Anton:700&display=swap', 'https://fonts.bunny.net/css?family=Anton:400&display=swap',
        'https://fonts.bunny.net/anton/files/anton-latin-400-normal.woff2']);
      expect(bold).toHaveLength(1);
      expect(created[created.length - 1]).toMatchObject({ family: 'Impact', descriptors: { weight: '400', style: 'normal' }, loads: 1 });
      expect(await loadRemoteFontFaces('Unknown Family', 'Unknown Family', { weight: 400, style: 'normal' }, fontSet)).toEqual([]);
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', { weight: 400, style: 'normal' }, fontSet, { provider: 'off' })).toEqual([]);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('one caller aborting leaves a shared stylesheet to the others', async () => {
    let release!: (response: Response) => void;
    const fetch = mock(async (_url: string | URL | Request, init?: RequestInit) => {
      expect(init?.signal).toBeDefined();
      return new Promise<Response>((resolve) => { release = resolve; });
    });
    const faces = [{ weight: 400, style: 'normal' as const }];
    const leaving = new AbortController();
    const first = fetchRemoteFontRules('Archivo Narrow', faces, { fetch: fetch as unknown as typeof globalThis.fetch, signal: leaving.signal });
    const second = fetchRemoteFontRules('Archivo Narrow', faces, { fetch: fetch as unknown as typeof globalThis.fetch });
    leaving.abort();
    await expect(first).rejects.toThrow();
    release(new Response(BUNNY_CSS));
    expect(await second).toHaveLength(2);
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  test('a caller abort reaches font downloads without AbortSignal.any', async () => {
    const any = AbortSignal.any;
    Object.defineProperty(AbortSignal, 'any', { configurable: true, value: undefined });
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const controller = new AbortController();
      const signals: AbortSignal[] = [];
      const fetch = mock(async (url: string | URL | Request, init?: RequestInit) => {
        if (String(url).includes('/css?')) return new Response(BUNNY_CSS);
        signals.push(init!.signal!);
        controller.abort();
        return new Response(new Uint8Array([1]));
      });
      const fontSet = { add: mock(() => {}) } as unknown as FontFaceSet;
      await expect(loadRemoteFontFaces('Tahoma', 'Tahoma', { weight: 400, style: 'normal' }, fontSet,
        { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.org/css' }, signal: controller.signal })).rejects.toThrow();
      expect(signals.length).toBeGreaterThan(0);
      expect(signals.every((signal) => signal.aborted)).toBe(true);
      expect((fontSet.add as unknown as ReturnType<typeof mock>)).not.toHaveBeenCalled();
    } finally {
      Object.defineProperty(AbortSignal, 'any', { configurable: true, value: any });
      globalThis.FontFace = original;
    }
  });

  test('registers nothing when aborted, when a font file fails, or when the face has too many slices', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const added: unknown[] = [];
      const fontSet = { add: (face: unknown) => { added.push(face); } } as unknown as FontFaceSet;
      const controller = new AbortController();
      const aborting = mock(async (url: string | URL | Request) => {
        if (String(url).includes('/css?')) return new Response(BUNNY_CSS);
        controller.abort();
        return new Response(new Uint8Array([1]));
      });
      await expect(loadRemoteFontFaces('Lucida Sans', 'Lucida Sans', { weight: 400, style: 'normal' }, fontSet,
        { fetch: aborting as unknown as typeof globalThis.fetch, signal: controller.signal })).rejects.toThrow();
      const failing = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(BUNNY_CSS) : new Response('', { status: 404 }));
      await expect(loadRemoteFontFaces('Segoe UI', 'Segoe UI', { weight: 400, style: 'normal' }, fontSet,
        { fetch: failing as unknown as typeof globalThis.fetch, provider: 'google' })).rejects.toThrow('404');
      const sliced = Array.from({ length: 13 }, (_, index) => `@font-face { font-family: 'Noto Sans JP'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/n/${index}.woff2) format('woff2'); unicode-range: U+${(0x4e00 + index * 16).toString(16)}-${(0x4e0f + index * 16).toString(16)}; }`).join('\n');
      const cjk = mock(async () => new Response(sliced));
      expect(await loadRemoteFontFaces('MS Gothic', 'MS Gothic', { weight: 400, style: 'normal' }, fontSet,
        { fetch: cjk as unknown as typeof globalThis.fetch })).toEqual([]);
      expect(cjk).toHaveBeenCalledTimes(1);
      expect(added).toEqual([]);
    } finally {
      globalThis.FontFace = original;
    }
  });
});
