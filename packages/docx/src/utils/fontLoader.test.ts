import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { configureDefaultFonts } = await import('../layout/measure/defaultFontProvider');
const { loadFont, setGoogleFontsEnabled } = await import('./fontLoader');

const added: Array<{ family: string; weight?: string; style?: string }> = [];

beforeAll(() => {
  class FakeFontFace {
    constructor(
      readonly family: string,
      _bytes: ArrayBuffer,
      readonly descriptors: { weight?: string; style?: string } = {}
    ) {}
    async load() {
      return this;
    }
  }
  Object.assign(globalThis, { FontFace: FakeFontFace });
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: {
      add: (face: FakeFontFace) =>
        added.push({ family: face.family, ...face.descriptors }),
      check: () => false,
      addEventListener() {},
      removeEventListener() {},
    },
  });
});

afterEach(() => {
  configureDefaultFonts({});
  setGoogleFontsEnabled(true);
  added.length = 0;
  document.head.innerHTML = '';
});

afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

const bundle = (families: string[]) => ({
  createFontProvider: () => ({
    resolve: (family: string) =>
      families.includes(family) ? async () => new ArrayBuffer(8) : undefined,
    resolveLastResort: () => async () => new ArrayBuffer(8),
  }),
});

const googleLinks = () =>
  [...document.head.querySelectorAll('link')].filter((link) =>
    link.href.includes('fonts.googleapis.com')
  );

test('a configured bundle serves the face without asking Google', async () => {
  configureDefaultFonts({ fonts: bundle(['Carlito']) });
  expect(await loadFont('Carlito')).toBe(true);
  expect(added.map((face) => face.family)).toEqual(['Carlito', 'Carlito', 'Carlito', 'Carlito']);
  expect(googleLinks()).toHaveLength(0);
});

test('a family the bundle lacks stays on its fallback stack without asking Google', async () => {
  configureDefaultFonts({ fonts: bundle([]) });
  expect(await loadFont('Unbundled Serif Display')).toBe(false);
  expect(added).toHaveLength(0);
  expect(googleLinks()).toHaveLength(0);
});

test('with the bundle and Google disabled, bundled faces still register', async () => {
  setGoogleFontsEnabled(false);
  configureDefaultFonts({ fonts: bundle(['Caladea']) });
  expect(await loadFont('Caladea')).toBe(true);
  expect(added).toHaveLength(4);
});
