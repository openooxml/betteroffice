import { afterEach, describe, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  DEFAULT_REMOTE_FONT_PROVIDER, WORD_FAMILY_SUBSTITUTES, fetchRemoteFontRules, fetchRemoteFontSlices, loadRemoteFontFaces,
  RemoteFontLoadError, isRemoteFontFace, parseRemoteFontCss, remoteFontSlices, remoteFontStylesheetUrl, resolveWordFamilySubstitute, unicodeRangeIncludes,
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

  test('unicode-range values match single points, spans and wildcards', () => {
    expect(unicodeRangeIncludes('U+0000-00FF,U+0131', 0x41)).toBe(true);
    expect(unicodeRangeIncludes('U+0000-00FF,U+0131', 0x131)).toBe(true);
    expect(unicodeRangeIncludes('U+0000-00FF,U+0131', 0x141)).toBe(false);
    expect(unicodeRangeIncludes('U+4??', 0x4ff)).toBe(true);
    expect(unicodeRangeIncludes('U+4??', 0x500)).toBe(false);
    expect(unicodeRangeIncludes(null, 0x1f600)).toBe(true);
  });

  test('the byte path fetches only the slices covering the text', async () => {
    const fetch = mock(async (url: string | URL | Request, init?: RequestInit) => {
      expect(init).toMatchObject({ credentials: 'omit', referrerPolicy: 'no-referrer' });
      return String(url).includes('/css?') ? new Response(BUNNY_CSS) : new Response(new Uint8Array([String(url).includes('-ext-') ? 2 : 1]));
    });
    const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/bytes/css' } };
    const face = { weight: 400, style: 'normal' } as const;
    const latin = await remoteFontSlices('Tahoma', face, { ...options, text: 'Hello' });
    expect(latin.map((rule) => rule.unicodeRange)).toEqual(['U+0000-00FF,U+0131,U+0152-0153']);
    expect((await fetchRemoteFontSlices(latin, options)).map((bytes) => new Uint8Array(bytes)[0])).toEqual([1]);
    const extended = await remoteFontSlices('Tahoma', face, { ...options, text: 'Łódź' });
    expect(extended.map((rule) => rule.unicodeRange)).toEqual(['U+0100-02AF,U+0304,U+0308', 'U+0000-00FF,U+0131,U+0152-0153']);
    expect(await remoteFontSlices('Tahoma', face, { ...options, text: '日' })).toEqual([]);
    expect(await remoteFontSlices('Tahoma', face, options)).toHaveLength(2);
    expect(fetch.mock.calls.map((call) => String(call[0]))).toEqual([
      'https://fonts.example.com/bytes/css?family=Open+Sans:400&display=swap',
      'https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff2',
    ]);
  });

  test('the font set path registers covering slices once and fetches only those newly typed characters add', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { constructor(readonly family: string, readonly source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
      async load() { return this; } } as unknown as typeof FontFace;
    try {
      const added: { descriptors: FontFaceDescriptors }[] = [];
      const fontSet = { add: (face: { descriptors: FontFaceDescriptors }) => { added.push(face); } } as unknown as FontFaceSet;
      const fetch = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(BUNNY_CSS) : new Response(new Uint8Array([1])));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/typed/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Hello' })).toHaveLength(1);
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Hello again' })).toHaveLength(1);
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Ł' })).toHaveLength(1);
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Hello Ł' })).toHaveLength(2);
      expect(fetch.mock.calls.map((call) => String(call[0]))).toEqual([
        'https://fonts.example.com/typed/css?family=Open+Sans:400&display=swap',
        'https://fonts.bunny.net/open-sans/files/open-sans-latin-400-normal.woff2',
        'https://fonts.bunny.net/open-sans/files/open-sans-latin-ext-400-normal.woff2',
      ]);
      expect(added.map((face) => face.descriptors.unicodeRange)).toEqual(['U+0000-00FF,U+0131,U+0152-0153', 'U+0100-02AF,U+0304,U+0308']);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('concurrent requests share one fetch per slice and a face removed from the set is fetched again', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const faces = new Set<unknown>();
      const fontSet = { add: (face: unknown) => { faces.add(face); }, has: (face: unknown) => faces.has(face) } as unknown as FontFaceSet;
      const fetch = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(BUNNY_CSS) : new Response(new Uint8Array([1])));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/shared/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      const [first, second] = await Promise.all([
        loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Hello' }),
        loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'World' }),
      ]);
      expect(first).toEqual(second);
      expect(fetch.mock.calls.filter((call) => !String(call[0]).includes('/css?'))).toHaveLength(1);
      faces.clear();
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Hello' })).toHaveLength(1);
      expect(fetch.mock.calls.filter((call) => !String(call[0]).includes('/css?'))).toHaveLength(2);
      expect(faces.size).toBe(1);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a caller that aborts stops waiting while the shared fetch completes for the others', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const faces = new Set<unknown>();
      const fontSet = { add: (face: unknown) => { faces.add(face); }, has: (face: unknown) => faces.has(face) } as unknown as FontFaceSet;
      let release!: () => void;
      let started!: () => void;
      const downloading = new Promise<void>((resolve) => { started = resolve; });
      const fetch = mock(async (url: string | URL | Request) => {
        if (String(url).includes('/css?')) return new Response(BUNNY_CSS);
        started();
        return new Promise<Response>((resolve) => { release = () => resolve(new Response(new Uint8Array([1]))); });
      });
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/abort/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      const leaving = new AbortController();
      const first = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A', signal: leaving.signal });
      await downloading;
      const second = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'B' });
      leaving.abort();
      await expect(first).rejects.toThrow();
      release();
      expect(await second).toHaveLength(1);
      expect(faces.size).toBe(1);
      expect(fetch.mock.calls.filter((call) => !String(call[0]).includes('/css?'))).toHaveLength(1);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a slice every caller abandoned is not registered and loads again later', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const faces = new Set<unknown>();
      const fontSet = { add: (face: unknown) => { faces.add(face); }, has: (face: unknown) => faces.has(face) } as unknown as FontFaceSet;
      let release!: () => void;
      let started!: () => void;
      const downloading = new Promise<void>((resolve) => { started = resolve; });
      const fetch = mock(async (url: string | URL | Request) => {
        if (String(url).includes('/css?')) return new Response(BUNNY_CSS);
        started();
        return new Promise<Response>((resolve) => { release = () => resolve(new Response(new Uint8Array([1]))); });
      });
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/abandoned/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      const leaving = new AbortController();
      const only = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A', signal: leaving.signal });
      await downloading;
      leaving.abort();
      await expect(only).rejects.toThrow();
      release();
      await new Promise((resolve) => setTimeout(resolve, 20));
      expect(faces.size).toBe(0);
      const again = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A' });
      for (let wait = 0; wait < 100 && fetch.mock.calls.length < 3; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
      release();
      expect(await again).toHaveLength(1);
      expect(faces.size).toBe(1);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('the slice record stops at 1024 slices without registering duplicates', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      let registered = 0;
      const fontSet = { add: () => { registered += 1; } } as unknown as FontFaceSet;
      const sliced = Array.from({ length: 512 }, (_, index) => `@font-face { font-family: 'Noto Sans JP'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/cap/${index}.woff2) format('woff2'); unicode-range: U+${(0x4e00 + index).toString(16)}; }`).join('\n');
      const fetch = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(sliced) : new Response(new Uint8Array([1])));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/cap/css' } };
      const text = Array.from({ length: 512 }, (_, index) => String.fromCodePoint(0x4e00 + index)).join('');
      for (let round = 0; round < 2; round++) {
        for (const alias of ['First', 'Second']) {
          expect(await loadRemoteFontFaces('MS Gothic', alias, { weight: 400, style: 'normal' }, fontSet, { ...options, text })).toHaveLength(512);
        }
        await expect(loadRemoteFontFaces('MS Gothic', 'Third', { weight: 400, style: 'normal' }, fontSet, { ...options, text })).rejects.toThrow('full');
      }
      expect(registered).toBe(1024);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a slice is registered only while a caller that asked for it still waits', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { constructor(readonly family: string, readonly source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
      async load() { return this; } } as unknown as typeof FontFace;
    try {
      const added: { descriptors: FontFaceDescriptors }[] = [];
      const fontSet = { add: (face: { descriptors: FontFaceDescriptors }) => { added.push(face); } } as unknown as FontFaceSet;
      const releases: (() => void)[] = [];
      const fetch = mock(async (url: string | URL | Request) => String(url).includes('/css?') ? new Response(BUNNY_CSS)
        : new Promise<Response>((resolve) => { releases.push(() => resolve(new Response(new Uint8Array([1])))); }));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/per-slice/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      const leaving = new AbortController();
      const both = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'AŁ', signal: leaving.signal });
      for (let wait = 0; wait < 100 && releases.length < 2; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
      const latin = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A' });
      leaving.abort();
      await expect(both).rejects.toThrow();
      for (const release of releases) release();
      expect(await latin).toHaveLength(1);
      await new Promise((resolve) => setTimeout(resolve, 20));
      expect(added.map((face) => face.descriptors.unicodeRange)).toEqual(['U+0000-00FF,U+0131,U+0152-0153']);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a full slice record drops slices whose faces left the set before refusing new ones', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const faces = new Set<unknown>();
      const fontSet = { add: (face: unknown) => { faces.add(face); }, has: (face: unknown) => faces.has(face) } as unknown as FontFaceSet;
      const sliced = Array.from({ length: 512 }, (_, index) => `@font-face { font-family: 'Noto Sans JP'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/sweep/${index}.woff2) format('woff2'); unicode-range: U+${(0x4e00 + index).toString(16)}; }`).join('\n');
      const fetch = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(sliced) : new Response(new Uint8Array([1])));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/sweep/css' } };
      const text = Array.from({ length: 512 }, (_, index) => String.fromCodePoint(0x4e00 + index)).join('');
      for (const alias of ['First', 'Second']) await loadRemoteFontFaces('MS Gothic', alias, { weight: 400, style: 'normal' }, fontSet, { ...options, text });
      faces.clear();
      expect(await loadRemoteFontFaces('MS Gothic', 'Third', { weight: 400, style: 'normal' }, fontSet, { ...options, text })).toHaveLength(512);
      expect(faces.size).toBe(512);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a failing slice leaves the other slices of a shared request registered', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { constructor(readonly family: string, readonly source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
      async load() { return this; } } as unknown as typeof FontFace;
    try {
      const added: { descriptors: FontFaceDescriptors }[] = [];
      const fontSet = { add: (face: { descriptors: FontFaceDescriptors }) => { added.push(face); } } as unknown as FontFaceSet;
      const releases: (() => void)[] = [];
      const fetch = mock(async (url: string | URL | Request) => String(url).includes('/css?') ? new Response(BUNNY_CSS)
        : new Promise<Response>((resolve) => { releases.push(() => resolve(String(url).includes('-ext-')
          ? new Response('', { status: 404 }) : new Response(new Uint8Array([1])))); }));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/isolated/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      const leaving = new AbortController();
      const both = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'AŁ', signal: leaving.signal });
      for (let wait = 0; wait < 100 && releases.length < 2; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
      const latin = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A' });
      leaving.abort();
      await expect(both).rejects.toThrow();
      for (const release of releases) release();
      expect(await latin).toHaveLength(1);
      expect(added.map((face) => face.descriptors.unicodeRange)).toEqual(['U+0000-00FF,U+0131,U+0152-0153']);
      expect(isRemoteFontFace(added[0] as unknown as FontFace)).toBe(true);
      const mixed = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'ÄŁ' });
      for (let wait = 0; wait < 100 && releases.length < 3; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
      releases[2]!();
      const error = await mixed.then(() => null, (reason: unknown) => reason);
      expect(error).toBeInstanceOf(RemoteFontLoadError);
      expect((error as RemoteFontLoadError).faces).toEqual([]);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a slice whose file never finishes fails at the timeout and is fetched again later', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const fontSet = { add: () => {} } as unknown as FontFaceSet;
      let stalled = true;
      const fetch = mock(async (url: string | URL | Request) => {
        if (String(url).includes('/css?')) return new Response(BUNNY_CSS);
        return stalled ? new Response(new ReadableStream({ start() {} })) : new Response(new Uint8Array([1]));
      });
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/stalled/css' }, timeoutMs: 20 };
      const face = { weight: 400, style: 'normal' } as const;
      await expect(loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A' })).rejects.toThrow();
      stalled = false;
      await new Promise((resolve) => setTimeout(resolve, 5));
      expect(await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A' })).toHaveLength(1);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a partial failure lists a shared slice that registered while the call waited', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const fontSet = { add: () => {} } as unknown as FontFaceSet;
      const releases: (() => void)[] = [];
      const fetch = mock(async (url: string | URL | Request) => String(url).includes('/css?') ? new Response(BUNNY_CSS)
        : new Promise<Response>((resolve) => { releases.push(() => resolve(String(url).includes('-ext-')
          ? new Response('', { status: 404 }) : new Response(new Uint8Array([1])))); }));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/joined/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      const leaving = new AbortController();
      const first = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A', signal: leaving.signal });
      for (let wait = 0; wait < 100 && releases.length < 1; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
      const joined = loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'AŁ' });
      for (let wait = 0; wait < 100 && releases.length < 2; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
      leaving.abort();
      await expect(first).rejects.toThrow();
      releases[1]!();
      releases[0]!();
      const error = await joined.then(() => null, (reason: unknown) => reason);
      expect(error).toBeInstanceOf(RemoteFontLoadError);
      expect((error as RemoteFontLoadError).faces).toHaveLength(1);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a face whose files do not all name their characters is not loaded by text', async () => {
    const css = ['U+0000-00FF', null, null].map((range, index) => `@font-face { font-family: 'Noto Sans SC'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/sc/${index}.woff2) format('woff2');${range ? ` unicode-range: ${range};` : ''} }`).join('\n');
    const fetch = mock(async () => new Response(css));
    const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/rangeless/css' } };
    expect(await remoteFontSlices('SimHei', { weight: 400, style: 'normal' }, { ...options, text: 'A中' })).toEqual([]);
    expect(await remoteFontSlices('SimHei', { weight: 400, style: 'normal' }, options)).toHaveLength(3);
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  test('slices sharing a file but not a range register separately', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { constructor(readonly family: string, readonly source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
      async load() { return this; } } as unknown as typeof FontFace;
    try {
      const added: { descriptors: FontFaceDescriptors }[] = [];
      const fontSet = { add: (face: { descriptors: FontFaceDescriptors }) => { added.push(face); } } as unknown as FontFaceSet;
      const css = ['U+0000-00FF', 'U+0100-02AF'].map((range) => `@font-face { font-family: 'Open Sans'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/open-sans/all.woff2) format('woff2'); unicode-range: ${range}; }`).join('\n');
      const fetch = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(css) : new Response(new Uint8Array([1])));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/same-file/css' } };
      const face = { weight: 400, style: 'normal' } as const;
      await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'A' });
      await loadRemoteFontFaces('Tahoma', 'Tahoma', face, fontSet, { ...options, text: 'Ł' });
      expect(added.map((face) => face.descriptors.unicodeRange)).toEqual(['U+0000-00FF', 'U+0100-02AF']);
    } finally {
      globalThis.FontFace = original;
    }
  });

  test('a face in many slices loads only the slices its text touches', async () => {
    const original = globalThis.FontFace;
    globalThis.FontFace = class { async load() { return this; } } as unknown as typeof FontFace;
    try {
      const fontSet = { add: () => {} } as unknown as FontFaceSet;
      const sliced = Array.from({ length: 100 }, (_, index) => `@font-face { font-family: 'Noto Sans JP'; font-style: normal; font-weight: 400; src: url(https://fonts.bunny.net/n/${index}.woff2) format('woff2'); unicode-range: U+${(0x4e00 + index * 16).toString(16)}-${(0x4e0f + index * 16).toString(16)}; }`).join('\n');
      const fetch = mock(async (url: string | URL | Request) =>
        String(url).includes('/css?') ? new Response(sliced) : new Response(new Uint8Array([1])));
      const options = { fetch: fetch as unknown as typeof globalThis.fetch, provider: { cssUrl: 'https://fonts.example.com/cjk/css' } };
      const faces = await loadRemoteFontFaces('MS Gothic', 'MS Gothic', { weight: 400, style: 'normal' }, fontSet,
        { ...options, text: String.fromCodePoint(0x4e00 + 5 * 16 + 3, 0x4e00 + 70 * 16) + 'a' });
      expect(faces).toHaveLength(2);
      expect(fetch.mock.calls.map((call) => String(call[0])).slice(1)).toEqual(['https://fonts.bunny.net/n/5.woff2', 'https://fonts.bunny.net/n/70.woff2']);
    } finally {
      globalThis.FontFace = original;
    }
  });
});
