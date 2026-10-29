import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';

import { fontPreflight, openDocument } from './index.js';
import {
  fontPreflight as wasmFontPreflight,
  preloadPreflightWasm,
} from '../../packages/docx/src/wasm/preflight';

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
  test('font preflight matches the standalone wasm on synthetic bytes', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:rFonts w:ascii="Run" w:cs="Complex"/><w:iCs/></w:rPr><w:t>Text العربية</w:t></w:r></w:p></w:body></w:document>`);
    zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Base" w:hAnsi="Base"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`);
    const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    await preloadPreflightWasm(readFileSync(new URL(
      '../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm',
      import.meta.url
    )));
    const native = fontPreflight(bytes);
    expect(native).toBe(JSON.stringify(await wasmFontPreflight(bytes)));
    expect(JSON.parse(native)).toEqual({
      complete: false,
      faces: [
        { family: 'Base', bold: false, italic: false, scripts: ['arabic'] },
        { family: 'Base', bold: true, italic: false, scripts: ['arabic'] },
        { family: 'Complex', bold: false, italic: false, scripts: ['arabic'] },
        { family: 'Complex', bold: true, italic: false, scripts: ['arabic'] },
        { family: 'Complex', bold: true, italic: true, scripts: ['arabic'] },
        { family: 'Run', bold: false, italic: false, scripts: ['arabic'] },
        { family: 'Run', bold: true, italic: false, scripts: ['arabic'] },
      ],
    });
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
