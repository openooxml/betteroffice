import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import * as wasmCore from '../../packages/docx/src/core';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeCore from '@betteroffice/docx-native';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { NATIVE_ENTRIES } from '../../scripts/native-engine';
import { expectSavedPartsEqual } from '../native-parity';
import { documentFixture, exercisePackage } from './package-script';

for (const [entry, source] of Object.entries(NATIVE_ENTRIES.docx)) {
  test(`matches the wasm source exports at ${entry}`, async () => {
    const specifier = '@betteroffice/docx-native' + (entry === '.' ? '' : entry.slice(1));
    const nativeKeys = Object.keys(await import(specifier));
    const wasmKeys = Object.keys(await import(`../../packages/docx/${source}`));
    expect(nativeKeys.sort()).toEqual(wasmKeys.sort());
  });
}

test('the same public package script returns equal results and saved parts', async () => {
  const bytes = await documentFixture();
  const actual = await exercisePackage(nativeYrs, nativeCore, bytes);
  const expected = await exercisePackage(wasmYrs, wasmCore, bytes);
  expect(actual.results).toEqual(expected.results);
  await expectSavedPartsEqual(actual.bytes, expected.bytes);
  const saved = await JSZip.loadAsync(actual.bytes);
  expect(await saved.file('custom/untouched.bin')!.async('uint8array')).toEqual(new Uint8Array([0, 127, 255]));
});

test('structured exports and content controls match through the package API', async () => {
  for (const fixture of [
    '../../crates/docx-edit/tests/fixtures/structured-export/principal.docx',
    '../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx',
  ]) {
    const bytes = new Uint8Array(readFileSync(new URL(fixture, import.meta.url)));
    for (const revisionView of ['accepted', 'original', 'markup'] as const) {
      const options = { revisionView, includeFormatting: false, maxBlocks: 10_000, maxBytes: 8_388_608 };
      expect(await nativeCore.exportDocxStructured(bytes, options)).toEqual(await wasmCore.exportDocxStructured(bytes, options));
      expect(await nativeCore.exportDocxMarkdown(bytes, options)).toEqual(await wasmCore.exportDocxMarkdown(bytes, options));
    }
    expect(await nativeCore.listDocxContentControls(bytes)).toEqual(await wasmCore.listDocxContentControls(bytes));
    const query = { kind: 'tag', tag: 'account.reference' } as const;
    expect(await nativeCore.findDocxContentControls(bytes, query)).toEqual(await wasmCore.findDocxContentControls(bytes, query));
    expect(await nativeCore.fontPreflight(bytes)).toEqual(await wasmCore.fontPreflight(bytes));
  }
});
