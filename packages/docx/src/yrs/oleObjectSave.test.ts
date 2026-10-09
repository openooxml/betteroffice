import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"`;
const OBJECT_WITH_FALLBACK =
  '<w:object><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdImage"/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_1" DrawAspect="Content" ObjectID="_1" r:id="rIdOle"/></w:object>';
const WATERMARK =
  '<w:pict><v:shape id="PowerPlusWaterMarkObject1" style="width:72pt;height:36pt"><v:imagedata r:id="rIdImage"/></v:shape></w:pict>';
const DANGLING =
  '<w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id=""/></v:shape></w:pict>';

function fixture(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string | Uint8Array) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/image.png"/></Relationships>`);
  set('word/media/image.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body><w:p><w:r>${OBJECT_WITH_FALLBACK}</w:r></w:p><w:p><w:r>${WATERMARK}</w:r></w:p><w:p><w:r>${DANGLING}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function opaqueKinds(parsed: Awaited<ReturnType<typeof parseDocx>>): string[] {
  const kinds: string[] = [];
  for (const block of parsed.package.document.content) {
    if (block.type !== 'paragraph') continue;
    for (const child of (block as Paragraph).content) {
      if (child.type !== 'run') continue;
      for (const content of child.content) {
        if (content.type === 'opaqueDrawing') kinds.push(content.kind);
      }
    }
  }
  return kinds;
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected']) {
  it(`${seeder} keeps OLE objects and unresolved picts as opaque drawings`, async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    expect(opaqueKinds(parsed)).toEqual(['object', 'pict', 'pict']);
    const session = await createYrsSession({ clientId: 74310 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const embeds = session.storySegments('body').filter((segment) => segment.kind === 'embed');
      expect(embeds.filter((segment) => segment.kind === 'embed' && segment.embedKind === 'opaqueDrawing')).toHaveLength(3);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      const saved = readDocxContainer(await repackDocx(yrsToDocument(session, parsed)));
      const documentXml = saved.text('word/document.xml') ?? '';
      expect(documentXml).toContain('OLEObject');
      expect(documentXml).toContain('PowerPlusWaterMarkObject1');
      expect(documentXml).toContain('r:id=""');
      expect(documentXml).toContain('Edited ');
    } finally {
      session.destroy();
    }
  });
}

it('native and projected seeders agree on OLE and unresolved pict state', async () => {
  const bytes = fixture();
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const native = await createYrsSession({ clientId: 74311 });
  const projected = await createYrsSession({ clientId: 74311 });
  try {
    native.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    expect(native.storySegments('body')).toEqual(projected.storySegments('body'));
  } finally {
    native.destroy();
    projected.destroy();
  }
});
