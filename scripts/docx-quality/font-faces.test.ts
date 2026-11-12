import { expect, test } from 'bun:test';
import { BUNDLED_FONTS, resolveMetricCompatFace } from '../../packages/fonts/src/manifest';
import { fontProvider } from '../../packages/fonts/src/provider';
import { fontFaceSummary } from './font-faces.mjs';
import { measuredFontProvider, type FontLoad } from './measured-fonts';
import { fontArmLine, fontSummary } from '../office-quality/readme.mjs';

const advertised = [...new Set(BUNDLED_FONTS.map((face) => face.file))];
const faceOf = (kind: string, args: unknown[]) =>
  kind === 'family' ? resolveMetricCompatFace(...(args as [string, boolean, boolean]))?.file : undefined;
const arm = (summary: ReturnType<typeof fontFaceSummary>) =>
  fontSummary([{ comparisons: [{ channel: 'published', fonts: { package: '@betteroffice/fonts@0.4.2', copy: 'local release copy', ...summary } }] }], 'published');

test('an advertised face that does not load fails the arm, with its reason', async () => {
  const carlito = resolveMetricCompatFace('Calibri', true, false)!.file;
  const provider = fontProvider(async (face) => {
    if (face.file === carlito) throw new Error(`Failed to fetch bundled font ${face.file}: HTTP 404`);
    return new ArrayBuffer(face.byteLength);
  });
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(provider, fontLoads, faceOf);
  await measured.resolve('Calibri', false, false)!();
  await expect(measured.resolve('Calibri', true, false)!()).rejects.toThrow('HTTP 404');
  expect(fontLoads.map(({ face, ok }) => ({ face, ok }))).toEqual([
    { face: resolveMetricCompatFace('Calibri', false, false)!.file, ok: true },
    { face: carlito, ok: false },
  ]);
  const summary = arm(fontFaceSummary(fontLoads, advertised));
  expect(summary).toMatchObject({ advertised: advertised.length, loaded: 1, failed: 1, reasons: { 'http 404': 1 } });
  expect(fontArmLine('published', summary)).toBe(`published arm: 1 of ${advertised.length} advertised faces did not load, by reason: http 404 1`);
});

test('a font the pack does not carry falls back without a failure', () => {
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(fontProvider(async () => { throw new Error('never loaded'); }), fontLoads, faceOf);
  expect(measured.resolve('A Family Nobody Bundles', false, false)).toBeUndefined();
  expect(fontLoads).toEqual([]);
  expect(arm(fontFaceSummary(fontLoads, advertised))).toMatchObject({ advertised: advertised.length, loaded: 0, failed: 0, reasons: {} });
});

test('a face whose request the capture policy refused counts as blocked, any other failure as error', () => {
  const summary = fontFaceSummary(
    [
      { kind: 'family', family: 'Calibri', ok: false, error: 'TypeError: Failed to fetch', face: 'Carlito-Regular.ttf' },
      { kind: 'family', family: 'Cambria', ok: false, error: 'TypeError: Failed to fetch', face: 'Caladea-Regular.ttf' },
      { kind: 'family', family: 'Calibri', ok: true, face: 'Carlito-Regular.ttf' },
      { kind: 'script', family: 'arabic', ok: true, face: 'NotoSansArabic-Regular.ttf' },
    ],
    ['Carlito-Regular.ttf', 'Caladea-Regular.ttf', 'NotoSansArabic-Regular.ttf'],
    [{ url: 'https://cdn.jsdelivr.net/npm/@betteroffice/fonts@0.4.2/assets/Carlito-Regular.ttf' }],
  );
  expect(summary).toEqual({
    advertised: 3,
    loaded_faces: ['NotoSansArabic-Regular.ttf'],
    failed_faces: { 'Caladea-Regular.ttf': 'error', 'Carlito-Regular.ttf': 'blocked' },
  });
});
