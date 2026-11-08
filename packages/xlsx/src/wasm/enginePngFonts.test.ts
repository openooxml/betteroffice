import { expect, test } from 'bun:test';
import { BUNDLED_FONTS } from '@betteroffice/fonts';
import cases from '../../../../crates/betteroffice-fonts/data/script-fallback-chains.json';
import { defaultSheetFaces } from './enginePngFonts';

const arabicFaces = BUNDLED_FONTS.filter((face) => face.script === 'arabic');

for (const { family, bold, arabic } of cases) {
  test(`a ${family}${bold ? ' bold' : ''} cell falls back for Arabic to the face the native chain gives`, async () => {
    const font = { family, bold, italic: false };
    const { faces, fallbacks } = await defaultSheetFaces([{ font, scripts: ['arabic'] }]);
    const chain = [...faces.filter((face) => face.family === family), ...fallbacks(font)];
    const files = chain.flatMap((face) =>
      arabicFaces.filter((known) => known.byteLength === face.bytes.byteLength).map((known) => known.file));
    expect(files).toEqual([arabic]);
  });
}
