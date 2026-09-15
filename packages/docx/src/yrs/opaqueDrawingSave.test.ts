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
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
const OBJECT =
  '<w:object><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_1" DrawAspect="Content" ObjectID="_1" r:id="rIdOle"/></w:object>';
const UNREADABLE_CHART =
  '<w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="2" name="Chart X"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart9"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const ALT_CONTENT =
  '<mc:AlternateContent><mc:Choice Requires="wps"><w:pict><v:rect style="width:10pt;height:10pt"/></w:pict></mc:Choice><mc:Fallback><w:pict><v:rect style="width:10pt;height:10pt"/></w:pict></mc:Fallback></mc:AlternateContent>';

/** One text paragraph plus one paragraph per unmodeled drawing kind. */
function fixture(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart9" Type="${R}/chart" Target="charts/chart9.xml"/></Relationships>`);
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:p><w:r>${OBJECT}</w:r></w:p><w:p><w:r>${UNREADABLE_CHART}</w:r></w:p><w:p><w:r>${ALT_CONTENT}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function opaqueXml(parsed: Awaited<ReturnType<typeof parseDocx>>): string[] {
  const xml: string[] = [];
  for (const block of parsed.package.document.content) {
    if (block.type !== 'paragraph') continue;
    for (const child of (block as Paragraph).content) {
      if (child.type !== 'run') continue;
      for (const content of child.content) {
        if (content.type === 'opaqueDrawing') xml.push(content.xml);
      }
    }
  }
  return xml;
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected']) {
  it(`${seeder} keeps unmodeled drawings through an editor save`, async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    expect(opaqueXml(parsed)).toHaveLength(3);
    const session = await createYrsSession({ clientId: 74020 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const embeds = session.storySegments('body').filter((segment) => segment.kind === 'embed');
      expect(embeds.filter((segment) => segment.embedKind === 'opaqueDrawing')).toHaveLength(3);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      const saved = readDocxContainer(await repackDocx(yrsToDocument(session, parsed)));
      const documentXml = saved.text('word/document.xml') ?? '';
      expect(documentXml).toContain('Edited Text');
      for (const xml of opaqueXml(parsed)) expect(documentXml).toContain(xml);
      const reopened = await parseDocx(await repackDocx(yrsToDocument(session, parsed)), {
        preloadFonts: false,
      });
      expect(opaqueXml(reopened)).toEqual(opaqueXml(parsed));
    } finally {
      session.destroy();
    }
  });
}

it('native and projected seeders agree on opaque drawing state', async () => {
  const bytes = fixture();
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const native = await createYrsSession({ clientId: 74021 });
  const projected = await createYrsSession({ clientId: 74021 });
  try {
    native.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    expect(native.storySegments('body')).toEqual(projected.storySegments('body'));
  } finally {
    native.destroy();
    projected.destroy();
  }
});
