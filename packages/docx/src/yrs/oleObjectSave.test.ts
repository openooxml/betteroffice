import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { OpaqueDrawingContent, Paragraph } from '../types/document';
import type { LayoutBlock } from '../layout/pagination/types';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"`;
const OBJECT_WITH_FALLBACK =
  '<w:object><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdImage"/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_1" DrawAspect="Content" ObjectID="_1" r:id="rIdOle"/></w:object>';
const WATERMARK =
  '<w:pict><v:shape id="PowerPlusWaterMarkObject1" style="width:72pt;height:36pt"><v:imagedata r:id="rIdImage"/></v:shape></w:pict>';
const DANGLING =
  '<w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id=""/></v:shape></w:pict>';

function fixture(hiddenOle = false, content = OBJECT_WITH_FALLBACK): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string | Uint8Array) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/image.png"/></Relationships>`);
  set('word/media/image.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body><w:p><w:r>${hiddenOle ? '<w:rPr><w:vanish/></w:rPr>' : ''}${content}</w:r></w:p><w:p><w:r>${WATERMARK}</w:r></w:p><w:p><w:r>${DANGLING}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function opaqueDrawings(parsed: Awaited<ReturnType<typeof parseDocx>>): OpaqueDrawingContent[] {
  const drawings: OpaqueDrawingContent[] = [];
  for (const block of parsed.package.document.content) {
    if (block.type !== 'paragraph') continue;
    for (const child of (block as Paragraph).content) {
      if (child.type !== 'run') continue;
      for (const content of child.content) {
        if (content.type === 'opaqueDrawing') drawings.push(content);
      }
    }
  }
  return drawings;
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

for (const seeder of ['native', 'projected']) {
  it(`${seeder} keeps OLE objects and unresolved picts as opaque drawings`, async () => {
    const bytes = fixture();
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const drawings = opaqueDrawings(parsed);
    expect(drawings.map((drawing) => drawing.kind)).toEqual(['object', 'pict', 'pict']);
    expect(drawings[0]!.xml).toBe(OBJECT_WITH_FALLBACK);
    expect(drawings[0]!.preview?.size).toEqual({ width: 914400, height: 457200 });
    expect(drawings.slice(1).every((drawing) => drawing.preview === undefined)).toBe(true);
    const session = await createYrsSession({ clientId: 74310 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const embeds = session.storySegments('body').filter((segment) => segment.kind === 'embed');
      expect(embeds).toHaveLength(3);
      expect(embeds.map((segment) => segment.embedKind)).toEqual(['opaqueDrawing', 'opaqueDrawing', 'opaqueDrawing']);
      expect(embeds[0]!.payload.xml).toBe(OBJECT_WITH_FALLBACK);
      expect(embeds[0]!.payload.preview).toMatchObject({ width: 96, height: 48, rId: 'rIdImage', wrapType: 'inline' });
      expect(embeds[1]!.payload.preview).toBeUndefined();
      expect(embeds[2]!.payload.preview).toBeUndefined();
      expect(session.paragraphSpans('body')[0]!.length).toBe(1);
      const second = session.paragraphs('body')[1]!;
      session.insertText({ story: 'body', paraId: second.paraId, offset: 0 }, 'Edited ');
      const projected = yrsToDocument(session, parsed);
      expect(opaqueDrawings(projected)[0]!.preview?.size).toEqual(drawings[0]!.preview!.size);
      expect(session.paragraphSpans('body')[0]!.length).toBe(1);
      const saved = readDocxContainer(await repackDocx(projected));
      const documentXml = saved.text('word/document.xml') ?? '';
      expect(documentXml.split(OBJECT_WITH_FALLBACK)).toHaveLength(2);
      expect(documentXml).toContain(WATERMARK);
      expect(documentXml).toContain(DANGLING);
      expect(documentXml).toContain('Edited ');
    } finally {
      session.destroy();
    }
  });
}

for (const seeder of ['native', 'projected']) {
  it(`${seeder} preserves a VML group when its paragraph text changes`, async () => {
    const group = '<w:pict><v:group coordsize="1000,1000" style="width:100pt;height:100pt"><v:shape style="left:0;top:0;width:500;height:500"><v:imagedata r:id="rIdImage"/></v:shape><v:shape style="left:500;top:500;width:500;height:500"><v:imagedata r:id="rIdImage"/></v:shape></v:group></w:pict>';
    const bytes = fixture(false, group);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    expect(opaqueDrawings(parsed)[0]!.previews).toHaveLength(2);
    const session = await createYrsSession({ clientId: 74315 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      expect(session.paragraphSpans('body')[0]!.length).toBe(1);
      const paragraph = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
      const projected = yrsToDocument(session, parsed);
      expect(opaqueDrawings(projected)[0]!.previews).toHaveLength(2);
      const saved = await repackDocx(projected);
      const documentXml = readDocxContainer(saved).text('word/document.xml') ?? '';
      expect(documentXml).toContain(group);
      expect(documentXml).toContain('Edited ');
      const reopened = await parseDocx(saved, { preloadFonts: false });
      expect(opaqueDrawings(reopened)[0]!.previews).toHaveLength(2);
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

it('native and projected seeders agree on hidden OLE marks and preview visibility', async () => {
  const bytes = fixture(true);
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const native = await createYrsSession({ clientId: 74317 });
  const projected = await createYrsSession({ clientId: 74317 });
  try {
    native.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    expect(native.storySegments('body')).toEqual(projected.storySegments('body'));
    for (const session of [native, projected]) {
      const embeds = session.storySegments('body').filter((segment) => segment.kind === 'embed');
      expect(embeds).toHaveLength(3);
      expect(embeds[0]!.embedKind).toBe('opaqueDrawing');
      expect(embeds[0]!.attributes.hidden).toBe(true);
      expect(embeds[0]!.payload.xml).toBe(OBJECT_WITH_FALLBACK);
      expect(embeds[0]!.payload.preview).toMatchObject({ width: 96, height: 48 });
      for (const showHiddenText of [false, true]) {
        const blocks = session.yrsBlocksForStory('body', { showHiddenText }) as LayoutBlock[];
        const images = blocks.flatMap((block) =>
          block.kind === 'paragraph' ? block.runs.filter((run) => run.kind === 'image') : []
        );
        expect(images).toHaveLength(showHiddenText ? 1 : 0);
      }
      const saved = readDocxContainer(await repackDocx(yrsToDocument(session, parsed)));
      const xml = saved.text('word/document.xml') ?? '';
      expect(xml.split(OBJECT_WITH_FALLBACK)).toHaveLength(2);
      expect(xml).toContain('<w:vanish');
    }
  } finally {
    native.destroy();
    projected.destroy();
  }
});

it('worker save keeps an OLE preview atom and its object XML verbatim', async () => {
  const bytes = fixture();
  const startWorker = await residentWorkerFactory();
  const worker = startWorker(74314);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.open(bytes);
    const resident = worker.sessions[0]!;
    const embeds = resident.storySegments('body').filter((segment) => segment.kind === 'embed');
    expect(embeds).toHaveLength(3);
    expect(embeds[0]!.embedKind).toBe('opaqueDrawing');
    expect(embeds[0]!.payload.preview).toMatchObject({ width: 96, height: 48, rId: 'rIdImage' });
    resident.applyRawOps('body', [{ op: 'insert', index: 2, text: 'Edited ' }]);
    for (let pass = 0; pass < 2; pass++) {
      const saved = await client.save({ comments: [] });
      const xml = readDocxContainer(saved.bytes).text('word/document.xml') ?? '';
      expect(xml.split(OBJECT_WITH_FALLBACK)).toHaveLength(2);
      expect(xml).toContain('Edited ');
      expect(xml).toContain(WATERMARK);
      expect(xml).toContain(DANGLING);
    }
  } finally {
    client.destroy();
  }
});
