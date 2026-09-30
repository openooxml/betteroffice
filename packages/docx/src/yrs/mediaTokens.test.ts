import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx, yrsToDocument } from './index';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
const PICTURE = `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/picture.png"/></Relationships>`)],
    ['word/media/unused.png', new Uint8Array([9, 9, 9])],
    ['word/media/picture.png', PNG],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p>${PICTURE}</w:p><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

for (const mediaTokens of [false, true]) {
  test(`an image resolves through its token and saves untouched (tokens: ${mediaTokens})`, async () => {
    const bytes = fixture();
    const session = await createYrsSession({ clientId: 3 });
    try {
      session.openDocx(bytes, true, { mediaTokens });
      const state = Buffer.from(session.encodeState()).toString('latin1');
      expect(state.includes('data:image/png;base64,')).toBe(!mediaTokens);
      expect(session.mediaSource('media:1')).toEqual({ bytes: PNG, mimeType: 'image/png' });
      expect(session.mediaSource('media:2')).toBeNull();
      const dataUrl = `data:image/png;base64,${Buffer.from(PNG).toString('base64')}`;
      expect(session.mediaDataUrl('media:1')).toBe(dataUrl);
      expect(session.mediaDataUrl(dataUrl)).toBeNull();

      const projected = yrsToDocument(session, session.materializeDocx()!);
      const content = JSON.stringify(projected.package.document.content);
      expect(content.match(/"src":"[^"]*"/g)).toEqual([`"src":"${dataUrl}"`]);
      expect(content).toContain('"rId":"rIdImage"');

      const saved = unzipContainer((await saveYrsDocx(session)).bytes);
      const source = unzipContainer(bytes);
      for (const part of ['word/media/picture.png', 'word/media/unused.png']) {
        expect(saved[part]).toEqual(source[part]);
      }
      expect(new TextDecoder().decode(saved['word/document.xml'])).toContain('r:embed="rIdImage"');
    } finally {
      session.destroy();
    }
  });
}
