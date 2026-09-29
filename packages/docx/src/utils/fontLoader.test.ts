import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, jest, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { configureDefaultFonts } = await import('../layout/measure/defaultFontProvider');
const {
  loadFont,
  loadFontFromBuffer,
  loadFontWithMapping,
  onFontError,
  onFontsLoaded,
  setGoogleFontsEnabled,
} = await import('./fontLoader');

type Face = { family: string; weight?: string; style?: string };

const added: Face[] = [];
const renderable = new Set<string>();
const dispatched: Array<{ type: string; fontfaces: Face[] }> = [];
let statusAtConstruction: 'loaded' | 'unloaded' = 'loaded';
let failNextFaceLoad = false;

class FakeFontFace {
  status: string = statusAtConstruction;
  readonly weight: string;
  readonly style: string;
  constructor(
    readonly family: string,
    _bytes: ArrayBuffer,
    descriptors: { weight?: string; style?: string } = {}
  ) {
    this.weight = descriptors.weight ?? 'normal';
    this.style = descriptors.style ?? 'normal';
  }
  async load() {
    if (failNextFaceLoad) {
      failNextFaceLoad = false;
      this.status = 'error';
      throw new Error('bad font data');
    }
    this.status = 'loaded';
    return this;
  }
}

const faceOf = (face: FakeFontFace): Face => ({
  family: face.family,
  weight: face.weight,
  style: face.style,
});

class FakeFontFaceSetLoadEvent extends Event {
  readonly fontfaces: FakeFontFace[];
  constructor(type: string, init: { fontfaces?: FakeFontFace[] } = {}) {
    super(type);
    this.fontfaces = init.fontfaces ?? [];
  }
}

beforeAll(() => {
  Object.assign(globalThis, { FontFace: FakeFontFace, FontFaceSetLoadEvent: FakeFontFaceSetLoadEvent });
  const probe = {
    font: '',
    textBaseline: 'top',
    measureText() {
      return { width: [...renderable].some((family) => probe.font.includes(`"${family}"`)) ? 90 : 60 };
    },
  };
  HTMLCanvasElement.prototype.getContext = (() => probe) as never;
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: {
      add: (face: FakeFontFace) => added.push(faceOf(face)),
      delete: (face: FakeFontFace) => {
        const index = added.findIndex(
          (entry) =>
            entry.family === face.family && entry.weight === face.weight && entry.style === face.style
        );
        if (index >= 0) added.splice(index, 1);
        return index >= 0;
      },
      dispatchEvent: (event: Event & { fontfaces?: FakeFontFace[] }) => {
        dispatched.push({ type: event.type, fontfaces: (event.fontfaces ?? []).map(faceOf) });
        return true;
      },
      check: () => false,
      addEventListener() {},
      removeEventListener() {},
    },
  });
});

afterEach(() => {
  jest.useRealTimers();
  configureDefaultFonts({});
  setGoogleFontsEnabled(true);
  added.length = 0;
  dispatched.length = 0;
  statusAtConstruction = 'loaded';
  failNextFaceLoad = false;
  renderable.clear();
  document.head.innerHTML = '';
});

afterAll(() => {
  if (ownsDom) GlobalRegistrator.unregister();
});

const bytes = async () => new ArrayBuffer(8);

const bundle = (families: string[], bundledNames: string[] = []) => ({
  createFontProvider: () => ({
    resolve: (family: string) => (families.includes(family) ? bytes : undefined),
    resolveFamily: (family: string) =>
      families.includes(family) || bundledNames.includes(family) ? bytes : undefined,
    resolveLastResort: () => bytes,
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

test('a bundled family name registers under that name', async () => {
  configureDefaultFonts({ fonts: bundle([], ['Gelasio']) });
  expect(await loadFont('Gelasio')).toBe(true);
  expect(added).toEqual([
    { family: 'Gelasio', weight: '400', style: 'normal' },
    { family: 'Gelasio', weight: '700', style: 'normal' },
    { family: 'Gelasio', weight: '400', style: 'italic' },
    { family: 'Gelasio', weight: '700', style: 'italic' },
  ]);
});

test('a mapped CJK family registers its bundled equivalent under the equivalent name', async () => {
  configureDefaultFonts({ fonts: bundle([], ['Noto Serif SC']) });
  expect(await loadFontWithMapping('SimSun')).toBe(true);
  expect(new Set(added.map((face) => face.family))).toEqual(new Set(['Noto Serif SC']));
  expect(googleLinks()).toHaveLength(0);
});

test('a family without italics registers only its upright faces, leaving italics to synthesis', async () => {
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: () => bytes,
        resolveFamily: (_family: string, _bold: boolean, italic: boolean) =>
          italic ? undefined : bytes,
      }),
    },
  });
  expect(await loadFont('Upright Sans')).toBe(true);
  expect(added.map((face) => face.style)).toEqual(['normal', 'normal']);
});

test('explicit weights and styles register only those faces', async () => {
  configureDefaultFonts({ fonts: bundle(['Weighted Sans']) });
  expect(await loadFont('Weighted Sans', { weights: [700], styles: ['italic'] })).toBe(true);
  expect(added).toEqual([{ family: 'Weighted Sans', weight: '700', style: 'italic' }]);
});

test('a stalled load settles false at the deadline, without Google, and a later call reuses it', async () => {
  jest.useFakeTimers();
  let release!: () => void;
  const stalled = new Promise<void>((resolve) => {
    release = resolve;
  });
  const stalledBytes = async () => {
    await stalled;
    return new ArrayBuffer(8);
  };
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({ resolve: () => stalledBytes, resolveFamily: () => stalledBytes }),
    },
  });
  const first = loadFont('Stalled Sans');
  jest.advanceTimersByTime(5000);
  expect(await first).toBe(false);
  expect(googleLinks()).toHaveLength(0);
  release();
  expect(await loadFont('Stalled Sans')).toBe(true);
  expect(added).toHaveLength(4);
});

test('a face that fails to load is removed and retried on the next call', async () => {
  const errors: Error[] = [];
  const unsubscribe = onFontError((error) => errors.push(error));
  configureDefaultFonts({ fonts: bundle(['Flaky Sans']) });
  failNextFaceLoad = true;
  expect(await loadFont('Flaky Sans', { weights: [400], styles: ['normal'] })).toBe(false);
  expect(added).toHaveLength(0);
  expect(errors).toHaveLength(1);
  expect(await loadFont('Flaky Sans', { weights: [400], styles: ['normal'] })).toBe(true);
  expect(added).toHaveLength(1);
  unsubscribe();
});

test('faces loaded before joining the set are announced once, after concurrent families finish', async () => {
  configureDefaultFonts({ fonts: bundle(['Announced Sans', 'Announced Serif']) });
  expect(
    await Promise.all([loadFont('Announced Sans'), loadFont('Announced Serif')])
  ).toEqual([true, true]);
  expect(dispatched).toHaveLength(1);
  expect(dispatched[0].type).toBe('loadingdone');
  expect(dispatched[0].fontfaces).toHaveLength(8);
});

test('a face that never loads does not hold back announcing the others', async () => {
  jest.useFakeTimers();
  let release!: () => void;
  const stalled = new Promise<void>((resolve) => {
    release = resolve;
  });
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: () => bytes,
        resolveFamily: (family: string) =>
          family === 'Stuck Sans'
            ? async () => {
                await stalled;
                return new ArrayBuffer(8);
              }
            : bytes,
      }),
    },
  });
  const stuck = loadFont('Stuck Sans');
  expect(await loadFont('Quick Sans')).toBe(true);
  expect(dispatched).toHaveLength(0);
  jest.advanceTimersByTime(500);
  expect(dispatched).toHaveLength(1);
  expect(new Set(dispatched[0].fontfaces.map((face) => face.family))).toEqual(
    new Set(['Quick Sans'])
  );
  jest.advanceTimersByTime(5000);
  expect(await stuck).toBe(false);
  release();
  expect(await loadFont('Stuck Sans')).toBe(true);
});

test('a family with one failed face renders, and the next call registers only that face', async () => {
  const unsubscribe = onFontError(() => undefined);
  configureDefaultFonts({ fonts: bundle(['Partial Serif']) });
  failNextFaceLoad = true;
  expect(await loadFont('Partial Serif')).toBe(true);
  expect(added).toHaveLength(3);
  expect(await loadFont('Partial Serif')).toBe(true);
  expect(added).toHaveLength(4);
  expect(await loadFont('Partial Serif')).toBe(true);
  expect(added).toHaveLength(4);
  unsubscribe();
});

test('a face that arrives after the deadline keeps the family off the system-font shortcut', async () => {
  jest.useFakeTimers();
  let release!: () => void;
  const stalled = new Promise<void>((resolve) => {
    release = resolve;
  });
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: () => bytes,
        resolveFamily: (_family: string, bold: boolean, italic: boolean) =>
          !bold && !italic
            ? async () => {
                await stalled;
                return new ArrayBuffer(8);
              }
            : bold && !italic
              ? bytes
              : undefined,
      }),
    },
  });
  failNextFaceLoad = true;
  const unsubscribe = onFontError(() => undefined);
  const first = loadFont('Late Serif');
  await Promise.resolve();
  jest.advanceTimersByTime(5000);
  expect(await first).toBe(false);
  release();
  await stalled;
  await new Promise((resolve) => setImmediate(resolve));
  renderable.add('Late Serif');
  expect(await loadFont('Late Serif')).toBe(true);
  expect(added.map((face) => face.weight).sort()).toEqual(['400', '700']);
  unsubscribe();
});

test('a mapped family the bundle lacks under its equivalent name registers the original family faces', async () => {
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: () => undefined,
        resolveFamily: (family: string, _bold: boolean, italic: boolean) =>
          family === 'Comic Sans MS' && !italic ? bytes : undefined,
      }),
    },
  });
  expect(await loadFontWithMapping('Comic Sans MS')).toBe(true);
  expect(added).toEqual([
    { family: 'Comic Neue', weight: '400', style: 'normal' },
    { family: 'Comic Neue', weight: '700', style: 'normal' },
  ]);
  expect(googleLinks()).toHaveLength(0);
});

test('a mapped family keeps its bundled fallback when its equivalent is already loading, in either order', async () => {
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: () => undefined,
        resolveFamily: (family: string, _bold: boolean, italic: boolean) =>
          (family === 'Impact' || family === 'Consolas') && !italic ? bytes : undefined,
      }),
    },
  });
  expect(await Promise.all([loadFont('Anton'), loadFontWithMapping('Impact')])).toEqual([false, true]);
  expect(await Promise.all([loadFontWithMapping('Consolas'), loadFont('Inconsolata')])).toEqual([
    true,
    true,
  ]);
  expect(added.map((face) => face.family).sort()).toEqual([
    'Anton',
    'Anton',
    'Inconsolata',
    'Inconsolata',
  ]);
});

test('aliases waiting on one stalled equivalent settle at a single deadline', async () => {
  jest.useFakeTimers();
  let release!: () => void;
  const stalled = new Promise<void>((resolve) => {
    release = resolve;
  });
  const stalledBytes = async () => {
    await stalled;
    return new ArrayBuffer(8);
  };
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: () => undefined,
        resolveFamily: (family: string) => (family === 'Noto Serif TC' ? stalledBytes : undefined),
      }),
    },
  });
  let settled = 0;
  const loads = ['PMingLiU', 'MingLiU', 'DFKai-SB'].map((family) =>
    loadFontWithMapping(family).finally(() => {
      settled += 1;
    })
  );
  await Promise.resolve();
  jest.advanceTimersByTime(5000);
  for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
  expect(settled).toBe(3);
  expect(await Promise.all(loads)).toEqual([false, false, false]);
  release();
});

test('a mapped retry after an absent equivalent keeps the deadline it started with', async () => {
  jest.useFakeTimers();
  let providerReady!: () => void;
  const ready = new Promise<void>((resolve) => {
    providerReady = resolve;
  });
  const stalled = new Promise<ArrayBuffer>(() => undefined);
  configureDefaultFonts({
    load: async () => {
      await ready;
      return {
        createFontProvider: () => ({
          resolve: () => undefined,
          resolveFamily: (family: string) => (family === 'Monaco' ? () => stalled : undefined),
        }),
      };
    },
  });
  let settled = false;
  const direct = loadFont('Fira Code');
  const mapped = loadFontWithMapping('Monaco').finally(() => {
    settled = true;
  });
  jest.advanceTimersByTime(3000);
  providerReady();
  expect(await direct).toBe(false);
  jest.advanceTimersByTime(2000);
  for (let tick = 0; tick < 50; tick += 1) await Promise.resolve();
  expect(settled).toBe(true);
  expect(await mapped).toBe(false);
});

test('an older provider without resolveFamily registers only the Regular face it can vouch for', async () => {
  configureDefaultFonts({
    fonts: {
      createFontProvider: () => ({
        resolve: (family: string) => (family === 'Legacy Sans' ? bytes : undefined),
      }),
    },
  });
  expect(await loadFont('Legacy Sans')).toBe(true);
  expect(added).toEqual([{ family: 'Legacy Sans', weight: '400', style: 'normal' }]);
  expect(await loadFont('Gelasio Legacy')).toBe(false);
});

test('faces the set sees loading are left to the set to announce', async () => {
  statusAtConstruction = 'unloaded';
  configureDefaultFonts({ fonts: bundle(['Spec Sans']) });
  expect(await loadFont('Spec Sans')).toBe(true);
  expect(added).toHaveLength(4);
  expect(dispatched).toHaveLength(0);
});

test('a registered family is announced once and not registered again', async () => {
  const loaded: string[][] = [];
  const unsubscribe = onFontsLoaded((fonts) => loaded.push(fonts));
  configureDefaultFonts({ fonts: bundle(['Once Sans']) });
  expect(await Promise.all([loadFont('Once Sans'), loadFont('Once Sans')])).toEqual([true, true]);
  expect(await loadFont('Once Sans')).toBe(true);
  expect(added).toHaveLength(4);
  expect(loaded).toEqual([['Once Sans']]);
  unsubscribe();
});

test('an embedded family still gets its bundled equivalent as a coverage fallback', async () => {
  configureDefaultFonts({ fonts: bundle(['Cousine']) });
  expect(await loadFontFromBuffer('Courier New', new ArrayBuffer(8))).toBe(true);
  expect(await loadFontWithMapping('Courier New')).toBe(true);
  expect(new Set(added.map((face) => face.family))).toEqual(new Set(['Cousine']));
});

test('without a bundle the Google lookup is unchanged', () => {
  jest.useFakeTimers();
  void loadFont('Unbundled Sans Local');
  expect(googleLinks()).toHaveLength(1);
});

test('without the FontFace API a bundled family settles false', async () => {
  const fontFace = globalThis.FontFace;
  Reflect.deleteProperty(globalThis, 'FontFace');
  try {
    configureDefaultFonts({ fonts: bundle(['Faceless Sans']) });
    expect(await loadFont('Faceless Sans')).toBe(false);
    expect(added).toHaveLength(0);
  } finally {
    Object.assign(globalThis, { FontFace: fontFace });
  }
});
