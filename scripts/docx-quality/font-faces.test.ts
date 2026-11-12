import { expect, test } from 'bun:test';
import { BASE_LATIN_FONTS, BUNDLED_FONTS, resolveMetricCompatFace } from '../../packages/fonts/src/manifest';
import { fontProvider } from '../../packages/fonts/src/provider';
import { fontFaceSummary } from './font-faces.mjs';
import { measuredFontProvider, type FontLoad } from './measured-fonts';
import { fontArmFailures, fontArmLine, fontSummary } from '../office-quality/readme.mjs';

const advertised = [...new Set(BUNDLED_FONTS.map((face) => face.file))];
const baseLatin = BASE_LATIN_FONTS.map((face) => face.file);
const faceOf = (kind: string, args: unknown[]) => {
  if (kind === 'base-latin') {
    const [bold, italic] = args as [boolean, boolean];
    return BASE_LATIN_FONTS.find((face) => face.weight === (bold ? 700 : 400) && face.style === (italic ? 'italic' : 'normal'))?.file;
  }
  return kind === 'family' ? resolveMetricCompatFace(...(args as [string, boolean, boolean]))?.file : undefined;
};
const arm = (summary: ReturnType<typeof fontFaceSummary>) =>
  fontSummary([{ comparisons: [{ channel: 'published', fonts: { package: '@betteroffice/fonts@0.4.2', copy: 'local release copy', measured: true, ...summary } }] }], 'published');

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
  expect(summary).toMatchObject({ advertised: advertised.length, requested: 2, loaded: 1, failed: 1, reasons: { 'http 404': 1 } });
  expect(fontArmLine('published', summary)).toBe(
    "published arm: 1 of 2 faces did not load (M = faces the documents requested that the arm's provider offers, base latin: none in this release), by reason: http 404 1"
  );
});

test('a font the pack does not carry falls back without a failure', () => {
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(fontProvider(async () => { throw new Error('never loaded'); }), fontLoads, faceOf);
  expect(measured.resolve('A Family Nobody Bundles', false, false)).toBeUndefined();
  expect(fontLoads).toEqual([]);
  expect(arm(fontFaceSummary(fontLoads, advertised))).toMatchObject({ requested: 0, loaded: 0, failed: 0, reasons: {} });
});

test('a load whose face is not in the advertised set is reported as unattributed and never fails the arm', () => {
  const summary = fontFaceSummary(
    [
      { kind: 'fallback', family: 'Wingdings', ok: false, error: 'Error: offline' },
      { kind: 'family', family: 'Odd', ok: false, error: 'Error: offline', face: 'NotAdvertised-Regular.ttf' },
      { kind: 'family', family: 'Odd', ok: true, face: 'NotAdvertised-Bold.ttf' },
    ],
    advertised,
  );
  expect(summary.failed_faces).toEqual({});
  expect(summary.unattributed).toEqual({ loaded: 1, failed: 2 });
  const line = fontArmLine('published', arm(summary));
  expect(line).toStartWith('published arm: 0 of 0 faces did not load');
  expect(line).toContain('unattributed loads (not named in the advertised set, never a failure): 1 loaded, 2 failed');
});

test('the commit arm measures the base Latin faces the engine loads and counts them as advertised', async () => {
  const fontLoads: FontLoad[] = [];
  const provider = fontProvider(async (face) => {
    if (face.file === baseLatin.find((file) => file.includes('-Bold.'))) throw new Error('Failed to fetch bundled font: HTTP 404');
    return new ArrayBuffer(face.byteLength);
  });
  const measured = measuredFontProvider(provider, fontLoads, faceOf) as any;
  expect(typeof measured.resolveBaseLatin).toBe('function');
  await measured.resolveBaseLatin(false, false)();
  await expect(measured.resolveBaseLatin(true, false)()).rejects.toThrow('HTTP 404');
  expect(fontLoads.map(({ kind, family, ok }) => ({ kind, family, ok }))).toEqual([
    { kind: 'base-latin', family: 'BetterOffice Base Latin', ok: true },
    { kind: 'base-latin', family: 'BetterOffice Base Latin', ok: false },
  ]);
  const summary = arm(fontFaceSummary(fontLoads, advertised, [], baseLatin));
  expect(summary).toMatchObject({ base_latin: baseLatin.length, requested: 2, failed: 1, reasons: { 'http 404': 1 } });
  expect(fontArmLine('published', summary)).toContain('base latin included');
  const { resolveBaseLatin: _omitted, ...release } = provider;
  expect('resolveBaseLatin' in measuredFontProvider(release, [], faceOf)).toBe(false);
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
    base_latin: 0,
    loaded_faces: ['NotoSansArabic-Regular.ttf'],
    failed_faces: { 'Caladea-Regular.ttf': 'error', 'Carlito-Regular.ttf': 'blocked' },
    unattributed: { loaded: 0, failed: 0 },
  });
});

test('a capture that recorded no font loads is counted, never skipped', () => {
  const summary = fontSummary(
    [{ comparisons: [{ channel: 'commit', fonts: { package: '@betteroffice/fonts@0.5.0', copy: 'workspace', measured: true } }] }],
    'commit',
  );
  expect(summary).toMatchObject({ captures: 0, missing: 1 });
  expect(fontArmLine('commit', summary)).toContain('1 captures recorded no font loads');
});

const capture = (id: string, fonts: Record<string, unknown>) => ({
  id,
  comparisons: [{ channel: 'commit', fonts: { package: '@betteroffice/fonts@0.5.0', copy: 'workspace', measured: true, ...fonts } }],
});
const measuredOk = { capture: 'ok', advertised: 3, loaded_faces: ['Carlito-Regular.ttf'], failed_faces: {} };

test('timed-out and errored captures do not fail the font check and are counted as not measured', () => {
  const samples = [capture('done', measuredOk), capture('slow', { capture: 'timeout' }), capture('broken', { capture: 'error' })];
  const summary = fontSummary(samples, 'commit');
  expect(summary).toMatchObject({ captures: 1, missing: 0, failed: 0, not_measured: { timed_out: 1, errored: 1 } });
  const { line, failures } = fontArmFailures(samples, 'commit');
  expect(failures).toEqual([]);
  expect(line).toEndWith('; captures not measured: 1 timed out, 1 errored');
});

test('a completed capture without a font record fails the arm by name', () => {
  const { failures } = fontArmFailures([capture('done', measuredOk), capture('silent', { capture: 'ok' })], 'commit');
  expect(failures).toEqual(['DOCX commit arm: 1 completed captures recorded no font loads: silent']);
});

test('a completed capture with a failed advertised face fails the arm', () => {
  const { line, failures } = fontArmFailures(
    [capture('done', { ...measuredOk, failed_faces: { 'Caladea-Regular.ttf': 'http 404' } })],
    'commit',
  );
  expect(failures).toEqual([line]);
  expect(line).toContain('1 of 2 faces did not load');
});

test('an arm with no completed capture fails as nothing was measured, never skipped', () => {
  const { failures } = fontArmFailures([capture('slow', { capture: 'timeout' }), capture('broken', { capture: 'error' })], 'commit');
  expect(failures).toEqual(['DOCX commit arm: nothing was measured; captures not measured: 1 timed out, 1 errored']);
});
