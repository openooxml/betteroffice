import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { createRustMeasureSource, type ResidentFontRequirement } from '../layout/measure/rustMeasureSource';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm',
)))));

test('resident pagination wraps all content after font provider loads reject', async () => {
  const paragraphs = Array.from({ length: 100 }, (_, index) => {
    const family = index % 2 ? 'Calibri' : 'Arial';
    return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="22"/></w:rPr><w:t>` +
      'The quiet river runs beside the green field. '.repeat(24) + '</w:t></w:r></w:p>';
  }).join('');
  const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paragraphs + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>')],
  ])));
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const counts: number[] = [];
  try {
    for (const blocked of [false, true]) {
      const session = await createYrsSession();
      const measurement = createRustMeasureSource({
        engine: {
          registerFont: data => session.registerFont(data),
          registerSubstituteFont: (id, family) => session.registerSubstituteFont(id, family),
          clearFonts() {},
        },
        bundled: { resolve: family => async () => {
          if (blocked) throw new Error('font CDN blocked');
          const name = family === 'Calibri' ? 'Carlito-Regular.ttf' : 'LiberationSans-Regular.ttf';
          const data = readFileSync(resolve(import.meta.dir, '../../../../packages/fonts/assets', name));
          return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
        } },
      });
      try {
        const { document } = session.openDocx(bytes, true);
        const layout = buildResidentRegionLayoutRequest(document, 24, {});
        const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(layout))) as ResidentFontRequirement[];
        await measurement.prepareFontRequirements(requirements);
        layout.measurement = measurement.measurementConfigForRequirements(requirements)!;
        const result = JSON.parse(session.layoutDocumentWithRegionsRetainedJson(JSON.stringify(layout))) as {
          layout: { pages: { fragments: { blockId: string; toLine?: number }[] }[]; partial?: boolean };
        };
        expect(result.layout.partial).not.toBe(true);
        expect(result.layout.pages.length).toBeGreaterThanOrEqual(6);
        const fragments = result.layout.pages.flatMap(page => page.fragments);
        expect(new Set(fragments.map(fragment => fragment.blockId)).size).toBe(100);
        expect(fragments.at(-1)!.toLine).toBeGreaterThan(1);
        counts.push(result.layout.pages.length);
      } finally { measurement.clear(); session.destroy(); }
    }
    expect(counts[1]!).toBeGreaterThanOrEqual(counts[0]!);
    expect(warn).toHaveBeenCalledTimes(1);
  } finally { warn.mockRestore(); }
}, 30000);
