import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:o="urn:schemas-microsoft-com:office:office"`;
const IMAGE =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="2" name="IMAGE"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="IMAGE"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
const SHAPE =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Rect 41"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvPr id="41" name="Rect 41"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>';
const CHART =
  '<w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const CHART_SPACE =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/></c:barChart></c:plotArea></c:chart></c:chartSpace>';
const OPAQUE =
  '<w:object><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_1" DrawAspect="Content" ObjectID="_1" r:id="rIdOle"/></w:object>';
const DATE = '2024-01-01T00:00:00Z';
const ins = (id: number, inner: string): string =>
  `<w:ins w:id="${id}" w:author="Ada" w:date="${DATE}"><w:r>${inner}</w:r></w:ins>`;

function fixture(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string | Uint8Array) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="${OFFICE_DOC}.drawingml.chart+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE_DOC}.wordprocessingml.comments+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/image.png"/><Relationship Id="rIdChart1" Type="${R}/chart" Target="charts/chart1.xml"/><Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/></Relationships>`);
  set('word/media/image.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  set('word/charts/chart1.xml', CHART_SPACE);
  set('word/comments.xml', `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="3" w:author="Ada" w:date="${DATE}" w:initials="A"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>`);
  const body = [
    `<w:p><w:r><w:t>ab</w:t></w:r></w:p>`,
    `<w:p>${ins(11, IMAGE)}</w:p>`,
    `<w:p>${ins(12, SHAPE)}</w:p>`,
    `<w:p>${ins(13, CHART)}</w:p>`,
    `<w:p>${ins(14, OPAQUE)}</w:p>`,
    `<w:p><w:del w:id="15" w:author="Ada" w:date="${DATE}"><w:r>${OPAQUE}</w:r></w:del></w:p>`,
    `<w:p><w:commentRangeStart w:id="3"/><w:r>${OPAQUE}</w:r><w:commentRangeEnd w:id="3"/></w:p>`,
  ].join('');
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected']) {
  it(`${seeder} keeps tracked-change wrappers and comment anchors around embed atoms`, async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const paragraphs = parsed.package.document.content.filter(
      (block): block is Paragraph => block.type === 'paragraph'
    );
    expect(paragraphs[1]?.content[0]?.type).toBe('insertion');
    const session = await createYrsSession({ clientId: 74110 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const embeds = session.storySegments('body').filter((segment) => segment.kind === 'embed');
      const embedded = (kind: string): number =>
        embeds.filter((segment) => segment.kind === 'embed' && segment.embedKind === kind).length;
      expect(embedded('image')).toBe(1);
      expect(embedded('shape')).toBe(1);
      expect(embedded('chart')).toBe(1);
      expect(embedded('opaqueDrawing')).toBe(3);
      for (const kind of ['image', 'shape', 'chart'] as const) {
        const segment = embeds.find((entry) => entry.kind === 'embed' && entry.embedKind === kind);
        if (segment?.kind !== 'embed') throw new Error(`missing ${kind} embed`);
        expect(segment.attributes.ins).toMatchObject({ author: 'Ada' });
      }
      const opaque = embeds.filter((entry) => entry.kind === 'embed' && entry.embedKind === 'opaqueDrawing');
      expect(opaque.filter((entry) => entry.kind === 'embed' && entry.attributes.ins !== undefined)).toHaveLength(1);
      expect(opaque.filter((entry) => entry.kind === 'embed' && entry.attributes.del !== undefined)).toHaveLength(1);
      const saved = yrsToDocument(session, parsed).package.document.content.filter(
        (block): block is Paragraph => block.type === 'paragraph'
      );
      expect(saved[1]?.content[0]?.type).toBe('insertion');
      expect(saved[2]?.content[0]?.type).toBe('insertion');
      expect(saved[3]?.content[0]?.type).toBe('insertion');
      expect(saved[4]?.content[0]?.type).toBe('insertion');
      expect(saved[5]?.content[0]?.type).toBe('deletion');
      const commented = saved[6]?.content.map((child) => child.type) ?? [];
      expect(commented).toContain('commentRangeStart');
      expect(commented).toContain('commentRangeEnd');
    } finally {
      session.destroy();
    }
  });
}

it('native and projected seeders agree on tracked embed state', async () => {
  const bytes = fixture();
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const native = await createYrsSession({ clientId: 74111 });
  const projected = await createYrsSession({ clientId: 74111 });
  try {
    native.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    expect(native.storySegments('body')).toEqual(projected.storySegments('body'));
  } finally {
    native.destroy();
    projected.destroy();
  }
});
