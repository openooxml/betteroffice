import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { Document, Paragraph } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:o="urn:schemas-microsoft-com:office:office"`;
const IMAGE =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="2" name="IMAGE"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="IMAGE"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
const CHART =
  '<w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing>';
const CHART_SPACE =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/></c:barChart></c:plotArea></c:chart></c:chartSpace>';
const OPAQUE =
  '<w:object><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_1" DrawAspect="Content" ObjectID="_1" r:id="rIdOle"/></w:object>';
const DATE = '2024-01-01T00:00:00Z';

function anchored(embed: string, bookmarkId: number, name: string): string {
  return `<w:p><w:r><w:t>ab</w:t></w:r><w:r>${embed}</w:r><w:r><w:t>cd</w:t></w:r><w:bookmarkStart w:id="${bookmarkId}" w:name="${name}"/><w:r><w:t>ef</w:t></w:r><w:bookmarkEnd w:id="${bookmarkId}"/></w:p>`;
}

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
    `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>ab</w:t></w:r><w:r>${OPAQUE}</w:r><w:r><w:t>cd</w:t></w:r><w:commentRangeEnd w:id="3"/></w:p>`,
    anchored(OPAQUE, 7, 'afterOpaque'),
    anchored(CHART, 8, 'afterChart'),
    anchored(IMAGE, 9, 'afterImage'),
  ].join('');
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function paragraphs(document: Document): Paragraph[] {
  return document.package.document.content.filter(
    (block): block is Paragraph => block.type === 'paragraph'
  );
}

function bookmarkOffsets(paragraph: Paragraph): Array<{ id: number; offset: number }> {
  const offsets: Array<{ id: number; offset: number }> = [];
  for (const child of paragraph.content) {
    if (child.type === 'bookmarkStart') offsets.push({ id: child.id, offset: child.position?.offset ?? -1 });
  }
  return offsets;
}

function runTexts(paragraph: Paragraph): string[] {
  const texts: string[] = [];
  for (const child of paragraph.content) {
    if (child.type !== 'run') continue;
    for (const content of child.content) {
      if (content.type === 'text') texts.push(content.text);
      else texts.push(`<${content.type}>`);
    }
  }
  return texts;
}

async function saveXml(bytes: Uint8Array<ArrayBuffer>, seeder: 'native' | 'projected', sessionId: number): Promise<{ saved: Document; xml: string }> {
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId: sessionId });
  try {
    if (seeder === 'native') session.seedFromDocx(bytes);
    else documentToYrs(session, parsed);
    const saved = yrsToDocument(session, parsed);
    const xml = readDocxContainer(await repackDocx(saved)).text('word/document.xml') ?? '';
    return { saved, xml };
  } finally {
    session.destroy();
  }
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected'] as const) {
  it(`${seeder} keeps anchor offsets stable across a no-op save after every embed kind`, async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const original = paragraphs(parsed);
    expect(bookmarkOffsets(original[1]!)).toEqual([{ id: 7, offset: 5 }]);
    expect(bookmarkOffsets(original[2]!)).toEqual([{ id: 8, offset: 5 }]);
    expect(bookmarkOffsets(original[3]!)).toEqual([{ id: 9, offset: 5 }]);
    const { saved, xml } = await saveXml(bytes, seeder, 74210);
    const kept = paragraphs(saved);
    expect(bookmarkOffsets(kept[1]!)).toEqual([{ id: 7, offset: 5 }]);
    expect(bookmarkOffsets(kept[2]!)).toEqual([{ id: 8, offset: 5 }]);
    expect(bookmarkOffsets(kept[3]!)).toEqual([{ id: 9, offset: 5 }]);
    expect(kept[0]!.content.map((child) => child.type)).toEqual([
      'commentRangeStart',
      'run',
      'run',
      'run',
      'commentRangeEnd',
    ]);
    expect(runTexts(kept[1]!)).toEqual(['ab', '<opaqueDrawing>', 'cd', 'ef']);
    expect(runTexts(kept[2]!)).toEqual(['ab', '<chart>', 'cd', 'ef']);
    expect(runTexts(kept[3]!)).toEqual(['ab', '<drawing>', 'cd', 'ef']);
    expect(xml).toContain('afterOpaque');
    const again = await parseDocx(new Uint8Array(await repackDocx(saved)).buffer as ArrayBuffer, {
      preloadFonts: false,
    });
    const reparsed = paragraphs(again);
    expect(bookmarkOffsets(reparsed[1]!)).toEqual([{ id: 7, offset: 5 }]);
    expect(bookmarkOffsets(reparsed[2]!)).toEqual([{ id: 8, offset: 5 }]);
    expect(bookmarkOffsets(reparsed[3]!)).toEqual([{ id: 9, offset: 5 }]);
  });

  it(`${seeder} keeps offsets and run boundaries stable when a save is reopened and saved again`, async () => {
    const bytes = fixture();
    const first = await saveXml(bytes, seeder, 74212);
    const second = await saveXml(
      new Uint8Array(await repackDocx(first.saved)),
      seeder,
      74213
    );
    for (const [index, id] of [[1, 7], [2, 8], [3, 9]] as const) {
      expect(bookmarkOffsets(paragraphs(first.saved)[index]!)).toEqual([{ id, offset: 5 }]);
      expect(bookmarkOffsets(paragraphs(second.saved)[index]!)).toEqual([{ id, offset: 5 }]);
      expect(runTexts(paragraphs(second.saved)[index]!)).toEqual(
        runTexts(paragraphs(first.saved)[index]!)
      );
    }
  });
}
