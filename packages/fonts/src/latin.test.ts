import { expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { LATIN_FONT_ASSET_URLS } from '../../fonts-latin/src';
import latinPackage from '../../fonts-latin/package.json';
import pack from '../../fonts-latin/assets/manifest.json';
import fontsPackage from '../package.json';
import before from './latin-resolution.json';
import { BUNDLED_FONTS, loadBundledFontBytes, resolveBundledFamilyFace } from './index';
import { isLatinFont } from './packed';

const latin = BUNDLED_FONTS.filter((face) => isLatinFont(face.file));

test('ships Latin at the fixed version as a regular dependency', () => {
  expect(latinPackage.version).toBe(fontsPackage.version);
  expect(fontsPackage.dependencies['@betteroffice/fonts-latin']).toBe('workspace:*');
  expect(latinPackage.exports['.'].import).toBe('./dist/index.js');
  expect(latinPackage.files).toEqual(['dist', 'assets', 'LICENSE', 'LICENSES', 'README.md']);
  expect(pack.schema).toBe(1);
});

test('has exactly one literal URL per original Latin and RTL face', async () => {
  const files = before.map((face) => `${face.file}.zst`).sort();
  expect(files).toHaveLength(70);
  expect(latin.map((face) => face.file).sort()).toEqual(files);
  expect(pack.faces.map((face) => face.file).sort()).toEqual(files);
  expect(Object.keys(LATIN_FONT_ASSET_URLS).sort()).toEqual(files);
  const source = await Bun.file(new URL('../../fonts-latin/src/index.ts', import.meta.url)).text();
  const rows = [...source.matchAll(/'([^']+)': \(\) => new URL\('\.\.\/assets\/([^']+)', import\.meta\.url\)/g)];
  expect(rows.map(([, file]) => file).sort()).toEqual(files);
  for (const [, file, target] of rows) {
    expect(file).toBe(target);
    expect(LATIN_FONT_ASSET_URLS[file]().pathname).toEndWith(`/assets/${file}`);
  }
  const checksums = await Bun.file(new URL('../../../scripts/fonts/latin-sources.sha256', import.meta.url)).text();
  expect(checksums.trim().split('\n').map((line) => line.split(/\s+/)[1]).sort())
    .toEqual(before.map((face) => face.file).sort());
});

test('preserves resolution for every original family and style', () => {
  for (const row of before) {
    const face = latin.find((face) => face.file === `${row.file}.zst`)!;
    expect(face.family).toBe(row.family);
    expect<string>(face.style).toBe(row.style);
    expect<number>(face.weight).toBe(row.weight);
    const resolved = resolveBundledFamilyFace(
      row.metricCompatWith ?? row.family, row.weight === 700, row.style === 'italic',
    );
    expect(resolved?.file).toBe(`${row.file}.zst`);
  }
});

test('decodes every installed Latin face with consumer-owned buffers', async () => {
  for (const face of latin) {
    const entry = pack.faces.find((entry) => entry.file === face.file)!;
    const first = await loadBundledFontBytes(face);
    const second = await loadBundledFontBytes(face);
    expect(entry.family).toBe(face.file.slice(0, face.file.indexOf('-')));
    expect(face.byteLength).toBe(entry.byteLength);
    expect(first.byteLength).toBe(face.byteLength);
    expect(createHash('sha256').update(new Uint8Array(first)).digest('hex')).toBe(entry.sha256);
    expect(new DataView(first).getUint32(0)).toBe(0x00010000);
    expect(first).not.toBe(second);
    structuredClone(first, { transfer: [first] });
    expect(new Uint8Array(await loadBundledFontBytes(face))).toEqual(new Uint8Array(second));
  }
});
