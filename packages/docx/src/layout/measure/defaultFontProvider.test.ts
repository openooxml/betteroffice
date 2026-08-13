import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';

import { configureDefaultFonts, resolveDefaultFontProvider } from './defaultFontProvider';
import { TextMeasureFontRegistry, type BundledFontProvider } from './fontRegistry';

const CARLITO_REGULAR = new URL('../../../../fonts/assets/Carlito-Regular.ttf', import.meta.url);

/** Records every buffer handed to the engine, in registration order. */
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

      const synthetic = warn.mock.calls.filter((call) =>
        String(call[0]).includes('synthetic metrics')
      );
      expect(synthetic).toHaveLength(1);
      expect(String(synthetic[0]![0])).toContain('@betteroffice/fonts');
    } finally {
      warn.mockRestore();
    }
  });

  /**
   * What a bundler that inlined `@betteroffice/fonts` produces: the provider
   * resolves, but its `import.meta.url` asset URLs miss and every load fails.
   * The advice must not be "install the package" — it already is installed.
   */
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

      const synthetic = warn.mock.calls
        .map((call) => String(call[0]))
        .filter((line) => line.includes('no font bytes'));
      expect(synthetic).toHaveLength(1);
      expect(synthetic[0]).toContain('mark it external');
      expect(synthetic[0]).not.toContain('Install @betteroffice/fonts');
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

      const synthetic = warn.mock.calls
        .map((call) => String(call[0]))
        .filter((line) => line.includes('no font bytes'));
      expect(synthetic).toHaveLength(1);
      expect(synthetic[0]).not.toContain('mark it external');
      expect(synthetic[0]).toContain('synthetic metrics');
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
        warn.mock.calls.filter((call) => String(call[0]).includes('synthetic metrics'))
      ).toHaveLength(0);
    } finally {
      warn.mockRestore();
    }
  });
});
