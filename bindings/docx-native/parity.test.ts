import { readFileSync } from 'node:fs';
import { beforeAll, expect, test } from 'bun:test';
import JSZip from 'jszip';
import {
  exportDocxMarkdown,
  exportDocxStructured,
  findDocxContentControls,
  listDocxContentControls,
  type DocxExportOptions,
} from '../../packages/docx/src/core';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import {
  fontPreflight as wasmFontPreflight,
  preloadPreflightWasm,
} from '../../packages/docx/src/wasm/preflight';
import { fontPreflight, openDocument } from './index.js';

const exportFixture = readFileSync(
  new URL('../../crates/docx-edit/tests/fixtures/structured-export/principal.docx', import.meta.url)
);
const controlsFixture = readFileSync(
  new URL('../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx', import.meta.url)
);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)
      )
    )
  )
);

test('matches the WASM bytes-level exports and content controls', async () => {
  for (const bytes of [exportFixture, controlsFixture]) {
    const document = await openDocument(bytes);
    for (const revisionView of ['accepted', 'original', 'markup'] as const) {
      const options: DocxExportOptions = {
        revisionView,
        stories: ['body', 'headers', 'footers', 'footnotes', 'endnotes', 'comments'],
        includeFormatting: false,
        maxBlocks: 10_000,
        maxBytes: 8_388_608,
      };
      expect(await document.exportStructured(options)).toEqual(
        await exportDocxStructured(bytes, options)
      );
      expect(await document.exportMarkdown(options)).toEqual(await exportDocxMarkdown(bytes, options));
    }
    expect(await document.listContentControls()).toEqual(await listDocxContentControls(bytes));
    const options = { stories: ['body', 'headers'] as const, maxControls: 10_000, maxBytes: 8_388_608 };
    expect(await document.listContentControls(options)).toEqual(
      await listDocxContentControls(bytes, options)
    );
    const query = { kind: 'tag', tag: 'account.reference' } as const;
    expect(await document.findContentControls(query, options)).toEqual(
      await findDocxContentControls(bytes, query, options)
    );
  }
});

test('native font preflight matches the standalone wasm on synthetic bytes', async () => {
  const zip = new JSZip();
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:rFonts w:ascii="Run" w:cs="Complex"/><w:iCs/></w:rPr><w:t>Text العربية</w:t></w:r></w:p></w:body></w:document>');
  zip.file('word/styles.xml', '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Base" w:hAnsi="Base"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>');
  const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await preloadPreflightWasm(readFileSync(new URL(
    '../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm',
    import.meta.url
  )));
  expect(fontPreflight(bytes)).toBe(JSON.stringify(await wasmFontPreflight(bytes)));
});
