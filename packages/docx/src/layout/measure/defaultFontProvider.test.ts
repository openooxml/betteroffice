import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';

import { configureDefaultFonts, resolveDefaultFontProvider } from './defaultFontProvider';
import { TextMeasureFontRegistry, type BundledFontProvider } from './fontRegistry';

const CARLITO_REGULAR = new URL('../../../../fonts/assets/Carlito-Regular.ttf', import.meta.url);

function recordingSink() {
  const registered: Uint8Array[] = [];
  return {
    registered,
    sink: {
      registerFont(bytes: Uint8Array): number {
        registered.push(bytes);
        return registered.length;
      },
    },
  };
}

function bytesOf(text: string): ArrayBuffer {
  return new TextEncoder().encode(text).buffer as ArrayBuffer;
}

describe('default font provider', () => {
  beforeEach(() => {
    configureDefaultFonts({});
  });

  afterEach(() => {
    configureDefaultFonts({});
  });

  test('resolves real font bytes with no explicit injection', async () => {
    const provider = await resolveDefaultFontProvider();
    expect(provider).toBeDefined();

    const loader = provider!.resolve('Calibri', false, false);
    expect(loader).toBeDefined();

    const bytes = await loader!();
    // sfnt, not woff2: skrifa/rustybuzz parse raw SFNT only.
    expect(new DataView(bytes).getUint32(0)).toBe(0x00010000);
    expect(new Uint8Array(bytes)).toEqual(new Uint8Array(await Bun.file(CARLITO_REGULAR).bytes()));
  });

  test('feeds those bytes into a chain with nothing injected', async () => {
    const { registered, sink } = recordingSink();
    const registry = new TextMeasureFontRegistry(sink, { bundled: resolveDefaultFontProvider });

    const chain = await registry.getFontIdChain('Calibri', false, false);

    expect(chain.length).toBeGreaterThan(0);
    expect(registered[0]).toEqual(await Bun.file(CARLITO_REGULAR).bytes());
  });

  test('serves the faces from a configured base URL', async () => {
    configureDefaultFonts({ baseUrl: new URL('../../../../fonts/assets/', import.meta.url) });

    const provider = await resolveDefaultFontProvider();
    const bytes = await provider!.resolve('Calibri', false, false)!();

    expect(new Uint8Array(bytes)).toEqual(new Uint8Array(await Bun.file(CARLITO_REGULAR).bytes()));
  });

  test('resolves CJK coverage faces through the optional add-on package', async () => {
    const provider = await resolveDefaultFontProvider();
    const loader = provider!.resolveScriptFallback!('cjk-sc', false, false);
    expect(loader).toBeDefined();

    const bytes = await loader!();
    // OTTO — the static CFF Regular, not a variable TTF.
    expect(new DataView(bytes).getUint32(0)).toBe(0x4f54544f);
  });

  test('an injected provider wins and the default is never loaded', async () => {
    let defaultLoaded = false;
    configureDefaultFonts({
      load: () => {
        defaultLoaded = true;
        return Promise.reject(new Error('should not be reached'));
      },
    });

    const injected: BundledFontProvider = {
      resolve: () => () => Promise.resolve(bytesOf('injected-face')),
    };
    const { registered, sink } = recordingSink();
    const registry = new TextMeasureFontRegistry(sink, { bundled: injected });

    const chain = await registry.getFontIdChain('Calibri', false, false);

    expect(chain).toEqual([1]);
    expect(new TextDecoder().decode(registered[0])).toBe('injected-face');
    expect(defaultLoaded).toBe(false);
  });

  test('retries a failed default-provider load', async () => {
    let attempts = 0;
    configureDefaultFonts({
      load: () => {
        attempts++;
        if (attempts === 1) return Promise.reject(new Error('transient import failure'));
        return Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () => Promise.resolve(bytesOf('recovered-provider')),
          }),
        });
      },
    });

    expect(await resolveDefaultFontProvider()).toBeUndefined();
    const recovered = await resolveDefaultFontProvider();
    expect(recovered).toBeDefined();
    expect(await resolveDefaultFontProvider()).toBe(recovered);
    expect(attempts).toBe(2);
  });

  test('lets the same registry recover from a transient provider-load failure', async () => {
    let attempts = 0;
    configureDefaultFonts({
      load: () => {
        attempts++;
        if (attempts === 1) return Promise.reject(new Error('transient import failure'));
        return Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () => Promise.resolve(bytesOf('recovered-provider')),
          }),
        });
      },
    });
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { registered, sink } = recordingSink();
      const registry = new TextMeasureFontRegistry(sink, { bundled: resolveDefaultFontProvider });

      expect(await registry.getFontIdChain('Calibri', false, false)).toEqual([]);
      expect(await registry.getFontIdChain('Calibri', false, false)).toEqual([1]);
      expect(new TextDecoder().decode(registered[0])).toBe('recovered-provider');
      expect(attempts).toBe(2);
    } finally {
      warn.mockRestore();
    }
  });

  test('pins a relative base URL when it is configured', async () => {
    const originalLocation = Object.getOwnPropertyDescriptor(globalThis, 'location');
    let receivedBaseUrl: string | URL | undefined;
    Object.defineProperty(globalThis, 'location', {
      configurable: true,
      value: new URL('https://example.test/docs/team/report/'),
    });
    try {
      configureDefaultFonts({
        baseUrl: 'fonts/',
        load: () =>
          Promise.resolve({
            createFontProvider: (providerOptions?: { baseUrl?: string | URL }) => {
              receivedBaseUrl = providerOptions?.baseUrl;
              return { resolve: () => undefined };
            },
          }),
      });
      Object.defineProperty(globalThis, 'location', {
        configurable: true,
        value: new URL('https://example.test/'),
      });

      await resolveDefaultFontProvider();

      expect(String(receivedBaseUrl)).toBe('https://example.test/docs/team/report/fonts/');
    } finally {
      if (originalLocation) Object.defineProperty(globalThis, 'location', originalLocation);
      else delete (globalThis as { location?: Location }).location;
    }
  });

  test('explains that server-side base URLs must be absolute', () => {
    expect(() => configureDefaultFonts({ baseUrl: 'fonts/' })).toThrow(
      'baseUrl must be absolute when no browser location exists'
    );
  });

  test('clear re-resolves a reconfigured default provider', async () => {
    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () => Promise.resolve(bytesOf('first-provider')),
          }),
        }),
    });
    const { registered, sink } = recordingSink();
    const registry = new TextMeasureFontRegistry(sink, { bundled: resolveDefaultFontProvider });

    expect(await registry.getFontIdChain('Calibri', false, false)).toEqual([1]);

    configureDefaultFonts({
      load: () =>
        Promise.resolve({
          createFontProvider: () => ({
            resolve: () => () => Promise.resolve(bytesOf('second-provider')),
          }),
        }),
    });
    registry.clear();

    expect(await registry.getFontIdChain('Calibri', false, false)).toEqual([2]);
    expect(new TextDecoder().decode(registered[1])).toBe('second-provider');
  });

  test('warns once when neither an injected nor a bundled provider is available', async () => {
    configureDefaultFonts({ load: () => Promise.reject(new Error('package not installed')) });
    const warn = spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const { registered, sink } = recordingSink();
      const registry = new TextMeasureFontRegistry(sink, { bundled: resolveDefaultFontProvider });

      expect(await registry.getFontIdChain('Calibri', false, false)).toEqual([]);
      expect(await registry.getFontIdChain('Cambria', true, false)).toEqual([]);
      expect(await registry.getFontIdChain('Arial', false, true)).toEqual([]);
      expect(registered).toHaveLength(0);

      const fallbacks = warn.mock.calls.filter((call) => String(call[0]).includes('no font bytes'));
      expect(fallbacks).toHaveLength(1);
      expect(String(fallbacks[0]![0])).toContain('Install @betteroffice/fonts');
      expect(String(fallbacks[0]![0])).toContain('measureText');
      expect(String(fallbacks[0]![0])).toContain('OS fonts');
    } finally {
      warn.mockRestore();
    }
  });

  // Inlining can resolve the provider while breaking its import.meta.url assets,
  // so telling this user to install the package again would be misleading.
  test('warns differently when a provider resolves but every face fails to load', async () => {
    const broken: BundledFontProvider = {
      resolve: () => () => Promise.reject(new Error('fetch failed')),
      resolveLastResort: () => () => Promise.reject(new Error('fetch failed')),
    };
    const warn = spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const { sink } = recordingSink();
      const registry = new TextMeasureFontRegistry(sink, { bundled: broken });

      expect(await registry.getFontIdChain('Calibri', false, false)).toEqual([]);

      const fallbacks = warn.mock.calls
        .map((call) => String(call[0]))
        .filter((line) => line.includes('no font bytes'));
      expect(fallbacks).toHaveLength(1);
      expect(fallbacks[0]).toContain('mark it external');
      expect(fallbacks[0]).not.toContain('Install @betteroffice/fonts');
      expect(fallbacks[0]).toContain('measureText');
    } finally {
      warn.mockRestore();
    }
  });

  test('does not blame bundling when a partial provider has no matching face', async () => {
    const partial: BundledFontProvider = {
      resolve: (family) =>
        family === 'Calibri' ? () => Promise.resolve(bytesOf('calibri')) : undefined,
    };
    const warn = spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const { sink } = recordingSink();
      const registry = new TextMeasureFontRegistry(sink, { bundled: partial });

      expect(await registry.getFontIdChain('Wingdings', false, false)).toEqual([]);

      const fallbacks = warn.mock.calls
        .map((call) => String(call[0]))
        .filter((line) => line.includes('no font bytes'));
      expect(fallbacks).toHaveLength(1);
      expect(fallbacks[0]).not.toContain('mark it external');
      expect(fallbacks[0]).not.toContain('pass your own measurementFontProvider');
      expect(fallbacks[0]).toContain('configured font provider has no matching face');
      expect(fallbacks[0]).toContain('OS fonts');
    } finally {
      warn.mockRestore();
    }
  });

  test('does not warn when the default provider covers the family', async () => {
    const warn = spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const { sink } = recordingSink();
      const registry = new TextMeasureFontRegistry(sink, { bundled: resolveDefaultFontProvider });

      await registry.getFontIdChain('Totally Unknown Face', false, false);

      expect(
        warn.mock.calls.filter((call) => String(call[0]).includes('no font bytes'))
      ).toHaveLength(0);
    } finally {
      warn.mockRestore();
    }
  });
});
