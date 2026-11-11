import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { pictureDocument, type PicturePart } from '../../e2e/docx-png-pictures';
import { typographyRows } from '../../e2e/fixtures/docx-typography';
import { resolveCompleteFontProvider } from '../../packages/docx/src/layout/measure';
import { engineKind } from '../../packages/docx/src/wasm/enginePng';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { preloadOpcWasm } from '../../packages/docx/src/wasm/opc';
import { createYrsSession } from '../../packages/docx/src/yrs';

/**
 * Child of `displayListParity.test.ts`: lays each shared fixture out through PNG
 * export on whichever engine this process runs, and writes every page's display
 * list as PNG export hands it to the engine. Usage: `<out.json> pinned|default`.
 */
const [out, fonts] = process.argv.slice(2);
if (!out || (fonts !== 'pinned' && fonts !== 'default')) throw new Error('usage: displayListParity.dump.ts <out.json> pinned|default');
const root = resolve(import.meta.dir, '../..');
const drawn = (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.displayListParity')] as string[] | undefined;
if (!drawn) throw new Error('displayListParity.capture.ts must be preloaded');

async function typographyDocument(): Promise<Uint8Array> {
  const paragraphs = typographyRows.map(row => {
    const properties = [
      `<w:rFonts w:ascii="${row.family}" w:hAnsi="${row.family}" w:cs="${row.family}"/>`,
      '<w:sz w:val="24"/><w:szCs w:val="24"/>',
      row.kerningMinPt === undefined ? '' : `<w:kern w:val="${row.kerningMinPt * 2}"/>`,
      row.rtl ? '<w:rtl/>' : '',
      row.ligatures === undefined ? '' : `<w14:ligatures w14:val="${row.ligatures}"/>`,
      row.contextualAlternates === undefined ? '' : `<w14:cntxtAlts w14:val="${Number(row.contextualAlternates)}"/>`,
    ].join('');
    return `<w:p>${row.rtl ? '<w:pPr><w:bidi/></w:pPr>' : ''}<w:r><w:rPr>${properties}</w:rPr><w:t>${row.text}</w:t></w:r></w:p>`;
  }).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

await preloadEditWasm(new Uint8Array(readFileSync(resolve(root, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))));
await preloadOpcWasm(new Uint8Array(readFileSync(resolve(root, 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
const fontProvider = fonts === 'pinned' ? await resolveCompleteFontProvider() : undefined;
if (fonts === 'pinned' && !fontProvider) throw new Error('the complete font provider did not resolve');

const fixtures: Record<string, Uint8Array> = {
  demo: new Uint8Array(readFileSync(resolve(root, 'apps/demo/public/betteroffice-demo.docx'))),
  typography: await typographyDocument(),
};
for (const part of ['none', 'drawn', 'missing', 'corrupt'] satisfies PicturePart[]) fixtures[`picture-${part}`] = await pictureDocument(part);

const pages: Record<string, string[]> = {};
for (const [name, bytes] of Object.entries(fixtures)) {
  const session = await createYrsSession();
  try {
    session.openDocx(bytes, true);
    drawn.length = 0;
    for (let page = 0; ; page++) {
      try {
        await session.renderPng({ page }, fontProvider ? { fontProvider } : {});
      } catch (error) {
        if (error instanceof RangeError && page > 0) break;
        throw error;
      }
    }
    pages[name] = [...drawn];
  } finally {
    session.destroy();
  }
}
writeFileSync(out, JSON.stringify({ engine: engineKind, fonts, pages }));
