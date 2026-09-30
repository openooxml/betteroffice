import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, jest, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { configureDefaultFonts } = await import('../layout/measure/defaultFontProvider');
const {
  createFontLoadScope,
  isFontLoaded,
  loadFont,
  loadFontFromBuffer,
  loadFontWithMapping,
  onFontError,
  onFontsLoaded,
  registerDocumentFaces,
  setGoogleFontsEnabled,
} = await import('./fontLoader');
type BufferFaceInput = import('./fontLoader').BufferFaceInput;

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

const faceStyles = () =>
  [...document.head.querySelectorAll('style')].map(
    (style) => /font-family: "([^"]+)"/.exec(style.textContent ?? '')?.[1]
  );
const fontBytes = (...bytes: number[]) => new Uint8Array(bytes).buffer;

test('a family Google does not serve is asked for once, and its link removed', async () => {
  const pending = loadFont('Unserved Office Face');
  const [link] = googleLinks();
  link.onerror?.(new Event('error'));
  expect(await pending).toBe(false);
  expect(googleLinks()).toHaveLength(0);
  expect(await loadFont('Unserved Office Face')).toBe(false);
  expect(googleLinks()).toHaveLength(0);
  const now = Date.now;
  Date.now = () => now() + 5 * 60_000;
  try {
    void loadFont('Unserved Office Face');
    expect(googleLinks()).toHaveLength(1);
    googleLinks()[0].onerror?.(new Event('error'));
  } finally {
    Date.now = now;
  }
});

test('a document declaring thousands of embedded families registers only the first 256', async () => {
  const scope = createFontLoadScope();
  const data = fontBytes(1, 2, 3);
  const faces = Array.from({ length: 5_000 }, (_, index) => ({
    family: `Budget Family ${index}`,
    data,
  }));
  try {
    const registered = await registerDocumentFaces(faces, scope);
    expect(registered.size).toBe(256);
    expect(registered).toEqual(
      new Map(faces.slice(0, 256).map(({ family }) => [family, family]))
    );
    expect(faceStyles()).toEqual(faces.slice(0, 256).map(({ family }) => family));
    expect(faceStyles().length).toBeLessThanOrEqual(256);
    expect(isFontLoaded(faces[256].family)).toBe(false);
  } finally {
    scope.dispose();
  }
});

test('the document face budget also caps faces sharing one family', async () => {
  const scope = createFontLoadScope();
  const data = fontBytes(4, 5, 6);
  const faces: BufferFaceInput[] = Array.from({ length: 257 }, (_, index) => ({
    family: 'Budget Weights',
    data,
    weight: index + 1,
  }));
  faces.push({ family: 'Past Face Budget', data });
  try {
    expect(await registerDocumentFaces(faces, scope)).toEqual(
      new Map([['Budget Weights', 'Budget Weights']])
    );
    expect(faceStyles()).toHaveLength(256);
    const weights = [...document.head.querySelectorAll('style')].map(
      (style) => /font-weight: (\d+);/.exec(style.textContent ?? '')?.[1]
    );
    expect(weights).toEqual(Array.from({ length: 256 }, (_, index) => String(index + 1)));
    expect(isFontLoaded('Past Face Budget')).toBe(false);
  } finally {
    scope.dispose();
  }
});

test('a buffer family stays loaded until its last face is released, even after repeated joins', async () => {
  const first = createFontLoadScope();
  const second = createFontLoadScope();
  const family = 'Counted Faces';
  const regular = { family, data: fontBytes(7), weight: 400 };
  const bold = { family, data: fontBytes(8), weight: 700 };
  try {
    await registerDocumentFaces([regular], first);
    await registerDocumentFaces([regular, bold], second);
    await registerDocumentFaces([regular], first);
    first.dispose();
    expect(isFontLoaded(family)).toBe(true);
    await registerDocumentFaces([bold], second);
    expect(isFontLoaded(family)).toBe(true);
    expect(faceStyles()).toEqual([family]);
    second.dispose();
    expect(isFontLoaded(family)).toBe(false);
    expect(faceStyles()).toEqual([]);
  } finally {
    first.dispose();
    second.dispose();
  }
});

test('documents embedding different faces under one name each keep their own', async () => {
  const first = createFontLoadScope();
  const second = createFontLoadScope();
  const faces = (data: ArrayBuffer) => [{ family: 'Shared Body', data, weight: 'normal' as const }];
  expect(await registerDocumentFaces(faces(fontBytes(1, 2, 3)), first)).toEqual(
    new Map([['Shared Body', 'Shared Body']])
  );
  const aliased = await registerDocumentFaces(faces(fontBytes(4, 5, 6)), second);
  const alias = aliased.get('Shared Body')!;
  expect(alias).not.toBe('Shared Body');
  expect(await registerDocumentFaces(faces(fontBytes(1, 2, 3)), createFontLoadScope())).toEqual(
    new Map([['Shared Body', 'Shared Body']])
  );
  expect(faceStyles()).toEqual(['Shared Body', alias]);
  first.dispose();
  second.dispose();
});

test("a scope's next document takes the family name its previous one held", async () => {
  const scope = createFontLoadScope();
  await registerDocumentFaces([{ family: 'Swapped Face', data: fontBytes(1) }], scope);
  expect(await registerDocumentFaces([{ family: 'Swapped Face', data: fontBytes(2) }], scope)).toEqual(
    new Map([['Swapped Face', 'Swapped Face']])
  );
  expect(faceStyles()).toEqual(['Swapped Face']);
  scope.dispose();
});

test('embedded faces live as long as a scope holds them', async () => {
  const revoked: string[] = [];
  const revoke = URL.revokeObjectURL;
  URL.revokeObjectURL = (url) => void revoked.push(url);
  try {
    const shown = createFontLoadScope();
    const other = createFontLoadScope();
    const face = { family: 'Held Face', data: fontBytes(7, 7, 7) };
    await registerDocumentFaces([face], shown);
    await registerDocumentFaces([face], other);
    shown.dispose();
    expect(faceStyles()).toEqual(['Held Face']);
    await registerDocumentFaces([{ family: 'Next Face', data: fontBytes(8) }], other);
    expect(faceStyles()).toEqual(['Next Face']);
    expect(revoked).toHaveLength(1);
    other.dispose();
    expect(faceStyles()).toEqual([]);
    expect(revoked).toHaveLength(2);
    expect(await loadFontFromBuffer('Kept Face', fontBytes(9))).toBe(true);
    expect(faceStyles()).toEqual(['Kept Face']);
  } finally {
    URL.revokeObjectURL = revoke;
  }
});

test('a scope hears its own loads and module-level ones, not another scope\'s', async () => {
  configureDefaultFonts({ fonts: bundle(['Scoped Serif', 'Module Serif', 'Broken Serif']) });
  const mine = createFontLoadScope();
  const theirs = createFontLoadScope();
  const heard = { mine: [] as string[], theirs: [] as string[], errors: [] as string[] };
  mine.onFontsLoaded((fonts) => heard.mine.push(...fonts));
  theirs.onFontsLoaded((fonts) => heard.theirs.push(...fonts));
  theirs.onFontError((error) => heard.errors.push(error.message));
  const unsubscribe = onFontError(() => {});
  await mine.loadFontsWithMapping(['Scoped Serif']);
  await loadFont('Module Serif');
  failNextFaceLoad = true;
  await mine.loadFontsWithMapping(['Broken Serif']);
  expect(heard).toEqual({ mine: ['Scoped Serif', 'Module Serif', 'Broken Serif'], theirs: ['Module Serif'], errors: [] });
  unsubscribe();
  mine.dispose();
  theirs.dispose();
});

test('a document reclaiming its faces mid-replacement keeps them marked as its name', async () => {
  const scope = createFontLoadScope();
  const face = (byte: number) => [{ family: 'Reclaimed Face', data: fontBytes(byte, byte) }];
  await registerDocumentFaces(face(1), scope);
  const replacing = registerDocumentFaces(face(2), scope);
  const reclaiming = registerDocumentFaces(face(1), scope);
  await Promise.all([replacing, reclaiming]);
  const other = createFontLoadScope();
  const [alias] = (await registerDocumentFaces(face(3), other)).values();
  expect(alias).not.toBe('Reclaimed Face');
  expect(faceStyles()).toEqual(['Reclaimed Face', alias]);
  scope.dispose();
  other.dispose();
});

test('names collide regardless of case, as CSS matches them', async () => {
  const first = createFontLoadScope();
  const second = createFontLoadScope();
  await registerDocumentFaces([{ family: 'Cased Face', data: fontBytes(4) }], first);
  const [alias] = (
    await registerDocumentFaces([{ family: 'cased face', data: fontBytes(5) }], second)
  ).values();
  expect(alias.toLowerCase()).not.toBe('cased face');
  first.dispose();
  second.dispose();
});

test('a scope joining another scope\'s pending face hears it load, and a revived scope hears again', async () => {
  const registering = createFontLoadScope();
  const joining = createFontLoadScope();
  const heard: string[] = [];
  joining.onFontsLoaded((fonts) => heard.push(...fonts));
  const registered = registerDocumentFaces([{ family: 'Joined Face', data: fontBytes(6) }], registering);
  await Promise.all([registered, joining.loadFontsWithMapping(['Joined Face'])]);
  expect(heard).toEqual(['Joined Face']);
  joining.dispose();
  joining.onFontsLoaded((fonts) => heard.push(...fonts));
  expect(joining.disposed).toBe(false);
  await registerDocumentFaces([{ family: 'Revived Face', data: fontBytes(7) }], joining);
  expect(heard).toEqual(['Joined Face', 'Revived Face']);
  registering.dispose();
  joining.dispose();
});

test('a disposed scope\'s face that another scope registers again still loads', async () => {
  const dropped = createFontLoadScope();
  const next = createFontLoadScope();
  const heard: string[] = [];
  next.onFontsLoaded((fonts) => heard.push(...fonts));
  const first = registerDocumentFaces([{ family: 'Churned Face', data: fontBytes(8) }], dropped);
  dropped.dispose();
  const second = registerDocumentFaces([{ family: 'Churned Face', data: fontBytes(8) }], next);
  await Promise.all([first, second]);
  expect(heard).toEqual(['Churned Face']);
  expect(faceStyles()).toEqual(['Churned Face']);
  next.dispose();
});

test('a mapped family is no longer loaded once the scoped face it relied on is released', async () => {
  const scope = createFontLoadScope();
  await registerDocumentFaces([{ family: 'Arimo', data: fontBytes(9) }], scope);
  await scope.loadFontsWithMapping(['Arial']);
  expect(isFontLoaded('Arial')).toBe(true);
  scope.dispose();
  expect(isFontLoaded('Arimo')).toBe(false);
  expect(isFontLoaded('Arial')).toBe(false);
});

test('a face joined by a family load is announced once to each listener', async () => {
  const mine = createFontLoadScope();
  const other = createFontLoadScope();
  const heard = { module: [] as string[], mine: [] as string[], other: [] as string[] };
  const unsubscribe = onFontsLoaded((fonts) => heard.module.push(...fonts));
  mine.onFontsLoaded((fonts) => heard.mine.push(...fonts));
  other.onFontsLoaded((fonts) => heard.other.push(...fonts));
  await Promise.all([
    registerDocumentFaces([{ family: 'Joined Once', data: fontBytes(10) }], mine),
    loadFont('Joined Once'),
    mine.loadFontsWithMapping(['Joined Once']),
  ]);
  expect(heard).toEqual({ module: ['Joined Once'], mine: ['Joined Once'], other: ['Joined Once'] });
  unsubscribe();
  mine.dispose();
  other.dispose();
});

test('a mapped family released while its load waited is not marked loaded', async () => {
  const scope = createFontLoadScope();
  await registerDocumentFaces([{ family: 'Libre Franklin', data: fontBytes(11) }], scope);
  const mapping = loadFontWithMapping('Franklin Gothic');
  scope.dispose();
  await mapping;
  expect(isFontLoaded('Libre Franklin')).toBe(false);
  expect(isFontLoaded('Franklin Gothic')).toBe(false);
});

test('the same face under another spelling of its name is loaded under that spelling too', async () => {
  setGoogleFontsEnabled(false);
  expect(await loadFontFromBuffer('Review Casing', fontBytes(12))).toBe(true);
  expect(await loadFontFromBuffer('review casing', fontBytes(12))).toBe(true);
  expect(faceStyles()).toEqual(['Review Casing']);
  expect(isFontLoaded('review casing')).toBe(true);
  expect(await loadFont('review casing')).toBe(true);
});

test('weights CSS treats as equal collide', async () => {
  const first = createFontLoadScope();
  const second = createFontLoadScope();
  for (const [weight, same] of [['normal', 400], ['bold', '700']] as const) {
    const family = `Weighted ${weight}`;
    await registerDocumentFaces([{ family, data: fontBytes(13), weight }], first);
    const [cssFamily] = (
      await registerDocumentFaces([{ family, data: fontBytes(14), weight: same }], second)
    ).values();
    expect(cssFamily).not.toBe(family);
  }
  first.dispose();
  second.dispose();
});

test('of two documents loading into one scope the later call wins, whichever faces arrive first', async () => {
  const scope = createFontLoadScope();
  await registerDocumentFaces([{ family: 'Shown Face', data: fontBytes(15) }], scope);
  let arrive!: (faces: BufferFaceInput[]) => void;
  const older = registerDocumentFaces(new Promise<BufferFaceInput[]>((resolve) => (arrive = resolve)), scope);
  expect(await registerDocumentFaces([], scope)).toEqual(new Map());
  expect(isFontLoaded('Shown Face')).toBe(false);
  expect(faceStyles()).toEqual([]);
  arrive([{ family: 'Overtaken Face', data: fontBytes(16) }]);
  expect(await older).toEqual(new Map());
  expect(faceStyles()).toEqual([]);
  scope.dispose();
});

test('a parsed document an editor shows takes back the name its previous document held', async () => {
  const shown = createFontLoadScope();
  await registerDocumentFaces([{ family: 'Parsed Collision', data: fontBytes(17) }], shown);
  const parsed = [{ family: 'Parsed Collision', data: fontBytes(18) }];
  const [pageAlias] = (await registerDocumentFaces(parsed)).values();
  expect(pageAlias).not.toBe('Parsed Collision');
  expect(await registerDocumentFaces(parsed, shown)).toEqual(
    new Map([['Parsed Collision', 'Parsed Collision']])
  );
  expect(faceStyles()).toEqual([pageAlias, 'Parsed Collision']);
  shown.dispose();
});

test('a claim made before a scope was disposed stays obsolete after it revives', async () => {
  const scope = createFontLoadScope();
  let arrive!: (faces: BufferFaceInput[]) => void;
  const pending = registerDocumentFaces(new Promise<BufferFaceInput[]>((resolve) => (arrive = resolve)), scope);
  scope.dispose();
  scope.onFontsLoaded(() => {});
  arrive([{ family: 'Disposed Claim Face', data: fontBytes(19) }]);
  expect(await pending).toEqual(new Map());
  expect(faceStyles()).toEqual([]);
  scope.dispose();
});

test('a scope joining a bundled face still loading past its deadline hears it fail', async () => {
  jest.useFakeTimers();
  let fail!: (error: Error) => void;
  const overdue = new Promise<ArrayBuffer>((_, reject) => (fail = reject));
  const load = () => overdue;
  configureDefaultFonts({
    fonts: { createFontProvider: () => ({ resolve: () => load, resolveFamily: () => load }) },
  });
  const first = createFontLoadScope();
  const second = createFontLoadScope();
  const heard: string[] = [];
  first.onFontError(() => {});
  second.onFontError((error) => heard.push(error.message));
  const firstLoad = first.loadFontsWithMapping(['Overdue Sans']);
  jest.advanceTimersByTime(5000);
  await firstLoad;
  first.dispose();
  const secondLoad = second.loadFontsWithMapping(['Overdue Sans']);
  fail(new Error('bundle unreachable'));
  await secondLoad;
  expect(heard).toHaveLength(4);
  second.dispose();
});
