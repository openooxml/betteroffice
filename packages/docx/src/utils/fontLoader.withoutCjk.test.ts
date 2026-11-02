import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';

type CjkAddon = { CJK_FONT_ASSET_URLS?: Record<string, () => URL> };
const missingCjkAddon = (): CjkAddon => {
  throw new Error("Cannot find package '@betteroffice/fonts-cjk'");
};
const importCjk = mock(missingCjkAddon);
mock.module('@betteroffice/fonts-cjk', () => ({
  then(resolve: (addon: CjkAddon) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importCjk).then(resolve, reject);
  },
}));

mock.module('@betteroffice/fonts-intl', () => ({
  then(_resolve: unknown, reject: (error: unknown) => void) {
    Promise.resolve().then(() => reject(new Error("Cannot find package '@betteroffice/fonts-intl'")));
  },
}));

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const fontManager = await import('@betteroffice/fonts');
const { configureDefaultFonts } = await import('../layout/measure/defaultFontProvider');
const { isFontLoaded, isGoogleFontsEnabled, loadFontWithMapping, setGoogleFontsEnabled } = await import('./fontLoader');
const googleFontsEnabled = isGoogleFontsEnabled();
const fontFace = Object.getOwnPropertyDescriptor(globalThis, 'FontFace');
const fontSet = Object.getOwnPropertyDescriptor(document, 'fonts');
const getContext = HTMLCanvasElement.prototype.getContext;

async function prepareCjk(family = 'Calibri', text = '中') {
  const { createRustMeasureSource } = await import('../layout/measure/rustMeasureSource');
  const wasm = await import('../layout/wasm');
  await wasm.preloadLayoutWasm();
  const registered: Uint8Array[] = [];
  const source = createRustMeasureSource({ engine: {
    registerFont(bytes) { registered.push(bytes); return wasm.registerMeasureFont(bytes); },
    clearFonts() {},
  } });
  const requirements: import('../layout/measure/rustMeasureSource').ResidentFontRequirement[] = [{
    key: `${family.toLowerCase()}|0|0`, family, bold: false, italic: false,
    scripts: text.includes('中') ? ['cjk-sc'] : [],
  }];
  await source.prepareFontRequirements(requirements);
  const config = source.measurementConfigForRequirements(requirements)!;
  return {
    config,
    registered,
    measure: (fontChains = config.fontChains) => JSON.parse(wasm.measureParagraphJson(JSON.stringify({
      ...config,
      fontChains,
      block: { kind: 'paragraph', runs: [{ kind: 'text', text, fontFamily: family }] },
      maxWidth: 400,
    }))),
  };
}

const added: Array<{ family: string; bytes: ArrayBuffer }> = [];
const google: string[] = [];
let append: ReturnType<typeof spyOn>;
let fetchSpy: ReturnType<typeof spyOn>;
class FakeFontFace {
  readonly status = 'loaded';
  constructor(readonly family: string, readonly bytes: ArrayBuffer) {}
  async load() { return this; }
}

beforeAll(() => {
  setGoogleFontsEnabled(true);
  Object.assign(globalThis, { FontFace: FakeFontFace });
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: { add: (face: FakeFontFace) => added.push(face), delete() {}, dispatchEvent() {} },
  });
  HTMLCanvasElement.prototype.getContext = (() => ({ measureText: () => ({ width: 60 }) })) as never;
  const appendChild = document.head.appendChild.bind(document.head);
  append = spyOn(document.head, 'appendChild').mockImplementation((node) => {
    if (node instanceof HTMLLinkElement) {
      google.push(node.href);
      queueMicrotask(() => node.dispatchEvent(new Event('error')));
      return node;
    }
    return appendChild(node);
  });
  fetchSpy = spyOn(globalThis, 'fetch').mockImplementation((async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    if (url.protocol !== 'file:') throw new Error(`Unexpected font request: ${url}`);
    return new Response(readFileSync(url));
  }) as unknown as typeof fetch);
});

afterEach(() => {
  configureDefaultFonts({});
  setGoogleFontsEnabled(true);
  added.length = 0;
  google.length = 0;
  document.head.innerHTML = '';
});

afterAll(() => {
  append.mockRestore();
  fetchSpy.mockRestore();
  setGoogleFontsEnabled(googleFontsEnabled);
  if (fontFace) Object.defineProperty(globalThis, 'FontFace', fontFace);
  else Reflect.deleteProperty(globalThis, 'FontFace');
  if (fontSet) Object.defineProperty(document, 'fonts', fontSet);
  else Reflect.deleteProperty(document, 'fonts');
  HTMLCanvasElement.prototype.getContext = getContext;
  if (ownsDom) GlobalRegistrator.unregister();
});

test('a missing CJK module leaves SimSun to Google without registering or caching Latin substitutes', async () => {
  expect(await loadFontWithMapping('SimSun')).toBe(false);
  expect(importCjk).toHaveBeenCalled();
  expect(google.map((url) => new URL(url).searchParams.get('family'))).toEqual([
    'Noto Serif SC:ital,wght@0,400;0,700;1,400;1,700',
  ]);
  expect(added).toEqual([]);
  expect(isFontLoaded('SimSun')).toBe(false);
  expect(isFontLoaded('Noto Serif SC')).toBe(false);
  expect(await loadFontWithMapping('SimSun')).toBe(false);
  expect(google).toHaveLength(1);
});

test('a script family whose pack is not installed uses Google without a managed substitute', async () => {
  expect(await loadFontWithMapping('Noto Sans Devanagari')).toBe(false);
  expect(google.map((url) => new URL(url).searchParams.get('family'))).toEqual([
    'Noto Sans Devanagari:ital,wght@0,400;0,700;1,400;1,700',
  ]);
  expect(added).toEqual([]);
  expect(isFontLoaded('Noto Sans Devanagari')).toBe(false);
});

test('a missing CJK module leaves measurement on the no-provider path', async () => {
  const { measure, registered } = await prepareCjk();
  expect(() => measure()).toThrow(/UNSUPPORTED/);
  expect(registered.length).toBeGreaterThan(0);
  expect(registered.every((bytes) => new DataView(bytes.buffer, bytes.byteOffset).getUint32(0) !== 0x4f54544f)).toBe(true);
});

for (const text of ['abc 123', '中 abc']) {
  test(`missing CJK keeps SimSun ${JSON.stringify(text)} on the no-provider measurement path`, async () => {
    const { config, measure, registered } = await prepareCjk('SimSun', text);
    expect(() => measure({})).toThrow(/UNSUPPORTED/);
    expect(() => measure()).toThrow(/UNSUPPORTED/);
    expect(config.fontChains).toEqual({});
    expect(registered).toEqual([]);
  });
}

test('missing CJK remains absent when Google is disabled', async () => {
  setGoogleFontsEnabled(false);
  expect(await loadFontWithMapping('PMingLiU')).toBe(false);
  expect(added).toEqual([]);
  expect(google).toEqual([]);
  expect(isFontLoaded('PMingLiU')).toBe(false);
});

test('a CJK module that fails to load leaves its family to Google', async () => {
  importCjk.mockImplementation(() => { throw new TypeError('CJK chunk unavailable'); });
  try {
    expect(await loadFontWithMapping('Batang')).toBe(false);
    expect(google.length).toBeGreaterThan(0);
    expect(added).toEqual([]);
    expect(isFontLoaded('Batang')).toBe(false);
  } finally { importCjk.mockImplementation(missingCjkAddon); }
});

test('a configured manager retains its missing-CJK substitution', async () => {
  configureDefaultFonts({ fonts: fontManager });
  expect(await loadFontWithMapping('MS Mincho')).toBe(true);
  expect(added.length).toBeGreaterThan(0);
  const latin = await fontManager.createFontProvider().resolve('Arial', false, false)!();
  expect(added.map((face) => new Uint8Array(face.bytes))).toEqual(
    added.map(() => new Uint8Array(latin)),
  );
  expect(google).toEqual([]);
});

test('an installed CJK module supplies the implicit manager face', async () => {
  importCjk.mockImplementation(() => ({
    CJK_FONT_ASSET_URLS: Object.fromEntries(fontManager.BUNDLED_FONTS.filter(
      (face) => face.script?.startsWith('cjk-'),
    ).map((face) => [face.file, () => new URL(`../../../fonts-cjk/assets/${face.file}`, import.meta.url)])),
  }));
  expect(await loadFontWithMapping('SimHei')).toBe(true);
  expect(added.length).toBeGreaterThan(0);
  expect(added.every((face) => face.family === 'Noto Sans SC')).toBe(true);
  expect(added.every((face) => new DataView(face.bytes).getUint32(0) === 0x4f54544f)).toBe(true);
  expect(google).toEqual([]);
  const { measure, registered } = await prepareCjk();
  expect(measure().kind).toBe('paragraph');
  expect(registered.some((bytes) => new DataView(bytes.buffer, bytes.byteOffset).getUint32(0) === 0x4f54544f)).toBe(true);
});
