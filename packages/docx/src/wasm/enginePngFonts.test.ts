import { expect, mock, test } from 'bun:test';
import * as fonts from '@betteroffice/fonts';
import { resolveCompleteFontProvider } from '../layout/measure';
import type { BundledFontLoader, BundledFontProvider } from '../layout/measure/fontRegistry';
import { faceLoadFailure, recordFaceLoads } from './enginePngFonts';

mock.module('@betteroffice/fonts-intl', () => ({
  then(_resolve: unknown, reject: (error: unknown) => void) {
    Promise.resolve().then(() => reject(new Error("Cannot find package '@betteroffice/fonts-intl'")));
  },
}));

const hosted = Symbol.for('@betteroffice/docx/hostedFontFamilies');
const bytes = new ArrayBuffer(8);

function providerOf(loader: BundledFontLoader): BundledFontProvider {
  const provider = { resolve: () => loader, resolveLastResort: () => loader, [hosted]: new Set(['brand sans']) };
  return provider;
}

test('a wrapped loader keeps every property of the loader it wraps, and the provider keeps its own', async () => {
  const slices = { ref: { file: 'x' } } as unknown as BundledFontLoader['slices'];
  const readOnly = new ArrayBuffer(4);
  const inner: BundledFontLoader = Object.assign(async () => bytes, {
    measureAs: 'Segoe UI', onlyIfPrimaryFails: true, slices, coverage: 'ff', readOnlyBytes: async () => readOnly,
  });
  const provider = providerOf(inner);
  const { provider: watched } = recordFaceLoads(provider, fonts);
  const loader = watched.resolve('Arial', false, false)!;
  expect(loader).not.toBe(inner);
  expect(loader).toMatchObject({ measureAs: 'Segoe UI', onlyIfPrimaryFails: true, slices, coverage: 'ff' });
  expect(await loader()).toBe(bytes);
  expect(await loader.readOnlyBytes!()).toBe(readOnly);
  expect((watched as unknown as Record<symbol, unknown>)[hosted]).toBe((provider as unknown as Record<symbol, unknown>)[hosted]);
});

test('each loader is wrapped once, whichever method returns it', () => {
  const inner: BundledFontLoader = async () => bytes;
  const { provider: watched } = recordFaceLoads(providerOf(inner), fonts);
  const loader = watched.resolve('Arial', false, false);
  expect(watched.resolve('Arial', false, false)).toBe(loader);
  expect(watched.resolveLastResort!('Arial', false, false, 'word')).toBe(loader);
});

test('a failed load is recorded with its face and family and still rejects, through the loader and its read-only bytes', async () => {
  const offline = new Error('offline');
  const inner: BundledFontLoader = Object.assign(() => Promise.reject(offline), { readOnlyBytes: () => Promise.reject(offline) });
  const { provider: watched, failures } = recordFaceLoads(providerOf(inner), fonts);
  const loader = watched.resolve('Arial', false, false)!;
  await expect(loader()).rejects.toBe(offline);
  await expect(loader.readOnlyBytes!()).rejects.toBe(offline);
  const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
  expect(failures).toEqual([
    { face, family: 'Arial', reason: 'offline' },
    { face, family: 'Arial', reason: 'offline' },
  ]);
  expect(faceLoadFailure(failures)).toMatchObject({
    name: 'PngExportError', code: 'font-unavailable', message: `${face.family} for Arial could not load: offline`,
  });
  expect(faceLoadFailure([])).toBeUndefined();
});

test('a face whose optional pack is not installed still rejects but is not recorded', async () => {
  const thai = fonts.BUNDLED_FONTS.find((face) => face.script === 'thai')!;
  const absent = await fonts.loadBundledFontBytes(thai).then(() => undefined, (error: unknown) => error);
  expect(fonts.isMissingFontAddon(absent)).toBe(true);
  expect(fonts.isMissingFontAddon(new Error(`Unknown bundled font asset: ${thai.file}`))).toBe(false);
  const { provider: watched, failures } = recordFaceLoads(providerOf(() => Promise.reject(absent)), fonts);
  await expect(watched.resolve('Arial', false, false)!()).rejects.toBe(absent);
  expect(failures).toEqual([]);
});

test('settling waits for every load in flight, and a failure that lands meanwhile is recorded', async () => {
  let fail!: (error: Error) => void;
  const inner: BundledFontLoader = () => new Promise<ArrayBuffer>((_, reject) => { fail = reject; });
  const { provider: watched, failures, settled } = recordFaceLoads(providerOf(inner), fonts);
  const loading = watched.resolve('Arial', false, false)!();
  loading.catch(() => {});
  let done = false;
  const settling = settled().then(() => { done = true; });
  await Promise.resolve();
  expect(done).toBe(false);
  fail(new Error('offline'));
  await settling;
  expect(failures.map((failure) => failure.reason)).toEqual(['offline']);
});

test('a font package without the pack check rejects with font-unavailable instead of skipping silently', () => {
  const older = { ...fonts, isMissingFontAddon: undefined } as unknown as typeof fonts;
  let thrown: unknown;
  try { recordFaceLoads(providerOf(async () => bytes), older); } catch (error) { thrown = error; }
  expect(thrown).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
});

test('the complete provider wrapped loads the same bytes', async () => {
  const complete = (await resolveCompleteFontProvider())!;
  const { provider: watched, failures } = recordFaceLoads(complete, fonts);
  const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
  const loader = watched.resolve('Arial', false, false)!;
  expect((await loader()).byteLength).toBe(face.byteLength);
  expect((await loader.readOnlyBytes!()).byteLength).toBe(face.byteLength);
  expect(failures).toEqual([]);
});
