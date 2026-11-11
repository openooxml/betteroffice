import { expect, spyOn, test } from 'bun:test';
import * as bundled from '@betteroffice/fonts';
import cases from '../../../../crates/betteroffice-fonts/data/script-fallback-chains.json';
import { documentFaceChain, loadDocumentFace } from './enginePngFonts';

const arabicFaces = bundled.BUNDLED_FONTS.filter((face) => face.script === 'arabic');
const sameFace = (a: bundled.BundledFontFace, b: bundled.BundledFontFace) => a.family === b.family && a.weight === b.weight && a.style === b.style;

async function nativeChain(family: string, bold: boolean, scripts: bundled.BundledFontScript[]): Promise<bundled.BundledFontFace[]> {
  const chain: bundled.BundledFontFace[] = [];
  const push = async (face: bundled.BundledFontFace | undefined) => {
    if (face && await bundled.bundledFaceInstalled(face) && !chain.some((known) => sameFace(known, face))) chain.push(face);
  };
  const primary = bundled.resolveBundledFamilyFace(family, bold, false);
  const installedPrimary = primary && await bundled.bundledFaceInstalled(primary) ? primary : undefined;
  await push(installedPrimary);
  for (const script of scripts) await push(bundled.resolveScriptFallbackFace(script, bold, false, family));
  await push(bundled.resolveGlyphFallbackFace(family, bold, false, 'word'));
  const last = bundled.resolveLastResortChoice(family, bold, false, 'word');
  if (!last.onlyIfPrimaryFails || !installedPrimary) await push(last.face);
  return chain;
}

for (const { family, bold, arabic } of cases) {
  test(`a ${family}${bold ? ' bold' : ''} run falls back for Arabic to the face the native chain gives`, async () => {
    const chain = await documentFaceChain({ family, bold, italic: false, scripts: ['arabic'] });
    const files = chain.flatMap((face) => arabicFaces.filter((known) => known.file === face.file).map((known) => known.file));
    expect(files).toEqual([arabic]);
  });
}

for (const [family, scripts] of [['Calibri', []], ['Calibri', ['hebrew']], ['Arial', ['arabic']], ['Brand Sans', []], ['Brand Sans', ['arabic', 'hebrew']]] as const) {
  test(`the ${family} chain${scripts.length ? ` for ${scripts.join(' and ')}` : ''} runs family face, script faces, glyph fallback, then last resort`, async () => {
    const chain = await documentFaceChain({ family, bold: false, italic: false, scripts: [...scripts] });
    expect(chain.map((face) => face.file)).toEqual((await nativeChain(family, false, [...scripts])).map((face) => face.file));
  });
}

test('a script face whose package is not installed is left out of the chain', async () => {
  const arabic = bundled.resolveScriptFallbackFace('arabic', false, false, 'Arial')!;
  const original = bundled.bundledFaceInstalled;
  const installed = spyOn(bundled, 'bundledFaceInstalled')
    .mockImplementation((face) => face.file === arabic.file ? Promise.resolve(false) : original(face));
  try {
    const chain = await documentFaceChain({ family: 'Arial', bold: false, italic: false, scripts: ['arabic'] });
    expect(chain.map((face) => face.file)).not.toContain(arabic.file);
    expect(chain.length).toBeGreaterThan(0);
  } finally { installed.mockRestore(); }
});

test('a face loads strictly, and one that cannot load rejects with font-unavailable naming it', async () => {
  const face = bundled.resolveScriptFallbackFace('hebrew', false, false, 'Calibri')!;
  const failing = spyOn(bundled, 'loadBundledFontBytes').mockImplementation(() => Promise.reject(new Error('offline')));
  try {
    await expect(loadDocumentFace(face, 'Calibri')).rejects.toMatchObject({
      name: 'PngExportError', code: 'font-unavailable', message: `${face.family} for Calibri could not load: offline`,
    });
    expect(failing.mock.calls.every((call) => call[2]?.strict === true)).toBe(true);
  } finally { failing.mockRestore(); }
});
