import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';

import { decodeFontFace, openDocument } from './index.js';
import { expectFontFace, fontCases, fontErrorCases } from '../font-fixtures';

const fixture = readFileSync(
  new URL('../../apps/demo/public/betteroffice-demo.docx', import.meta.url)
);
const layoutInput = JSON.parse(
  readFileSync(
    new URL(
      '../../crates/docx-layout/tests/fixtures/single-page-multi-paragraph.input.json',
      import.meta.url
    ),
    'utf8'
  )
);
const font = readFileSync(
  new URL('../../crates/docx-raster/tests/assets/Carlito-Regular.ttf', import.meta.url)
);

describe('@betteroffice/docx-native', () => {
  test('linked font decoder returns known answers and shared errors', () => {
    expect(typeof decodeFontFace).toBe('function');
    for (const entry of fontCases()) {
      const decoded = decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput);
      expectFontFace(decoded, entry);
      for (const bytes of [decoded.font, decoded.payload].filter((value) => value !== undefined)) {
        expect(bytes).toBeInstanceOf(Uint8Array);
        expect(Buffer.isBuffer(bytes)).toBe(false);
        expect(Object.getPrototypeOf(bytes)).toBe(Uint8Array.prototype);
      }
    }
    for (const entry of fontErrorCases()) {
      expect(() => decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput))
        .toThrow(new Error(entry.message));
    }
  });

  test('opens, inspects, lays out, renders, and saves a document', async () => {
    const document = await openDocument(fixture, { author: 'Node test' });

    const structure = await document.structure;
    const paragraphIds = await document.paragraphIds;
    const text = await document.text;

    expect(structure.bodyParagraphs).toBeGreaterThan(0);
    expect(document.author).toBe('Node test');
    expect(paragraphIds).toHaveLength(structure.bodyParagraphs);
    expect(text.length).toBeGreaterThan(0);
    const layout = await document.layout(layoutInput);
    await document.registerFont({ family: 'Calibri', data: font });
    const rendered = await document.renderPage(layout.displayList);

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await document.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('rejects invalid edit origins', async () => {
    expect(() => openDocument(fixture, { origin: 'unknown' })).toThrow('origin must be');
  });
});
