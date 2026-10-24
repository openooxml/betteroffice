import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { applyFrameDeltaOwned, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import { displayPrimitiveRect } from '../layout/render/displayListGeometry';
import { preloadEditWasm } from './wasm/index';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createResidentEngineSession } from './residentEngineSession';
import * as projectionModule from './yrsToDocument';
import * as facadeModule from './yrsSessionFacade';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { ResidentSourceWorkerClient } from './residentSourceWorkerClient';
import { createYrsSession } from './index';
import { resolveReservations, type DocxReservation } from './reservations';
import type { DocxCommandAck, DocxCommandEnvelope, DocxCommandFrame, ResidentCommandInitRequest } from './residentCommandProtocol';
import type { ResidentEngineWorkerResponse } from './residentEngineWorkerProtocol';
import type { ResidentReviewTarget } from './residentReview';
import type { YrsContentControlValue } from './index';
import type { SectionProperties } from '../types/document';
import type { LayoutBlock } from '../layout/pagination/types';

const TIMEOUT = 20_000;
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});

function syntheticDocument(headersFooters = false, notes = false, controls = false, bodyXml?: string, sectPr?: string, commentsXml?: string, commentsExtendedXml?: string, imageMedia = false): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ));
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
    (bodyXml ?? '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r>' +
    (notes ? '<w:r><w:footnoteReference w:id="1"/></w:r>' : '') + '</w:p>') +
    (controls ? '<w:p><w:sdt><w:sdtPr><w:id w:val="7"/><w:tag w:val="check"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt></w:p>' +
      '<w:sdt><w:sdtPr><w:id w:val="8"/><w15:repeatingSection xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"/><w:tag w:val="item"/><w15:repeatingSectionItem xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Item text</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>' : '') +
    (sectPr ?? (headersFooters ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>' : '<w:sectPr/>')) +
    '</w:body></w:document>'
  ));
  if (imageMedia) {
    parts.set('word/media/picture.png', new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64')));
    parts.set('word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>'));
    parts.set('[Content_Types].xml', toBytes(new TextDecoder().decode(parts.get('[Content_Types].xml')!).replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')));
  }
  if (headersFooters) {
    const types = new TextDecoder().decode(parts.get('[Content_Types].xml')!);
    parts.set('[Content_Types].xml', toBytes(types.replace('</Types>',
      '<Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
    )));
    parts.set('word/_rels/document.xml.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer.xml"/></Relationships>'
    ));
    parts.set('word/header.xml', toBytes('<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>'));
    parts.set('word/footer.xml', toBytes('<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>F</w:t></w:r></w:p></w:ftr>'));
  }
  if (notes) {
    parts.set('[Content_Types].xml', toBytes(new TextDecoder().decode(parts.get('[Content_Types].xml')!).replace('</Types>',
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')));
    const relationships = parts.get('word/_rels/document.xml.rels');
    const xml = relationships ? new TextDecoder().decode(relationships) : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    parts.set('word/_rels/document.xml.rels', toBytes(xml.replace('</Relationships>', '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')));
    parts.set('word/footnotes.xml', toBytes('<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>'));
  }
  if (commentsXml) {
    parts.set('[Content_Types].xml', toBytes(new TextDecoder().decode(parts.get('[Content_Types].xml')!).replace('</Types>',
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')));
    const related = '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>';
    const rels = parts.get('word/_rels/document.xml.rels');
    parts.set('word/_rels/document.xml.rels', toBytes(rels
      ? new TextDecoder().decode(rels).replace('</Relationships>', `${related}</Relationships>`)
      : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${related}</Relationships>`));
    parts.set('word/comments.xml', toBytes(`<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${commentsXml}</w:comments>`));
  }
  if (commentsExtendedXml) {
    parts.set('[Content_Types].xml', toBytes(new TextDecoder().decode(parts.get('[Content_Types].xml')!).replace('</Types>',
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>')));
    parts.set('word/_rels/document.xml.rels', toBytes(new TextDecoder().decode(parts.get('word/_rels/document.xml.rels')!).replace('</Relationships>',
      '<Relationship Id="rIdCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>')));
    parts.set('word/commentsExtended.xml', toBytes(`<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${commentsExtendedXml}</w15:commentsEx>`));
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export async function openWorker(options: { collaboration?: boolean; headersFooters?: boolean; layoutWithRegions?: boolean; notes?: boolean; controls?: boolean; propertyImages?: boolean; objects?: boolean; imageMedia?: boolean; bodyXml?: string; sectPr?: string; comments?: string; displayWindow?: [number, number]; factory?: typeof startWorker; prepare?: (engine: Awaited<ReturnType<typeof createResidentEngineSession>>) => { story: string; paraId: string; offset: number }; commentsXml?: string; commentsExtendedXml?: string } = {}) {
  const worker = (options.factory ?? startWorker)(8301);
  const client = options.collaboration ? new ResidentSourceWorkerClient(worker) : new ResidentEngineWorkerClient(worker);
  try {
    const bytes = syntheticDocument(options.headersFooters, options.notes, options.controls, options.bodyXml, options.sectPr, options.commentsXml ?? options.comments, options.commentsExtendedXml, options.imageMedia);
    if (client instanceof ResidentSourceWorkerClient) await client.openSource({ bytes, collaboration: { clientId: 8301 } });
    else await client.open(bytes, { generation: 'd1-command-worker' });
    const engine = worker.sessions[0]!;
    if (options.propertyImages || options.objects) engine.applyRawOps('body', [1, 2].map(index => ({
      op: 'insertEmbed', index, kind: 'image', payload: {
        id: 'shared-image', src: 'data:image/png;base64,', ...(options.objects ? { width: 40, height: 20 } : { widthEmu: 381000, heightEmu: 190500 }),
      },
    })));
    const prepared = options.prepare?.(engine);
    const read = engine.proposalEngine.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    const point = prepared ?? { story: 'body', paraId: read.paragraphs[0]!.paraId, offset: Math.min(4, engine.geometryReader.paragraphSpans('body')[0]!.length) };
    const fonts = [new Uint8Array(readFileSync(resolve(
      process.cwd(), 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
    )))];
    const fontId = engine.registerFont(fonts[0]!);
    const measurement = {
      defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
      fontChains: { 'liberation sans|0|0': [fontId] },
    };
    let layoutInput = JSON.stringify({ ...JSON.parse(LAYOUT), measurement });
    if (options.notes) layoutInput = JSON.stringify({ ...JSON.parse(layoutInput), notes: { contents: [{ id: 1, noteKind: 'footnote', height: 0 }] } });
    let extras = JSON.stringify({ fontChains: { ...measurement.fontChains, 'calibri|0|0': [fontId] } });
    if (options.headersFooters) {
      layoutInput = JSON.stringify({ ...JSON.parse(layoutInput), regions: { sections: [{
        sectionId: 'main', headerFooterRefs: { headerDefault: 'rIdHeader', footerDefault: 'rIdFooter' },
      }] } });
      engine.layoutDocumentWithRegionsRetained(layoutInput);
      extras = JSON.stringify({ ...JSON.parse(extras), headersFooters: JSON.parse(engine.retainedHeadersFootersJson()!) });
    }
    if (options.layoutWithRegions === false) {
      const block = engine.yrsBlocksForStory('body')[0];
      const measure = JSON.parse(engine.measureParagraphJson(JSON.stringify({ block, maxWidth: 180, ...measurement })));
      layoutInput = JSON.stringify({
        measured: [{ block, measure: { kind: 'paragraph', ...measure } }],
        options: { pageSize: { w: 200, h: 120 }, margins: { top: 10, right: 10, bottom: 10, left: 10 } },
        measurement, renderEnv: {},
      });
    }
    const init: ResidentCommandInitRequest = {
      session: 'd1-command-worker', incarnation: 1, epoch: 1, readOnly: false,
      selection: { anchor: point, focus: point },
      fonts, layoutInput, layoutWithRegions: options.layoutWithRegions ?? true, layoutRevision: 1, extras,
      ...(options.displayWindow ? { displayWindow: options.displayWindow, retainBuiltPages: false } : {}),
    };
    const initial = await client.commandInit(init);
    return { worker, client, engine, init, initial, point };
  } catch (error) {
    client.destroy();
    throw error;
  }
}

export function intent(seq: number, version: number, kind: string, payload: DocxCommandEnvelope['payload']): DocxCommandEnvelope {
  return { session: 'd1-command-worker', incarnation: 1, seq, basis: { epoch: 1, version }, kind, payload };
}

async function editWithFrame(client: ResidentEngineWorkerClient, command: DocxCommandEnvelope): Promise<DocxCommandAck> {
  let resolveFrame!: (frame: DocxCommandFrame) => void;
  const deferred = new Promise<DocxCommandFrame>((resolve) => { resolveFrame = resolve; });
  const unsubscribe = client.onFrame((frame) => {
    if (frame.coversThroughSeq === command.seq) resolveFrame(frame);
  });
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    const ack = await client.edit(command);
    if (ack.frame) return ack;
    expect(['select', 'navigate', 'selectPoint', 'extendSelection', 'selectTableCells', 'navigateCell']).toContain(command.kind);
    expect(ack.outcomes[0]?.status).toBe('noop');
    expect(ack.throughSeq).toBe(command.seq);
    const frame = await Promise.race([
      deferred,
      new Promise<never>((_, reject) => {
        timeout = setTimeout(() => reject(new Error(`Missing command frame through seq ${ack.throughSeq}`)), 5_000);
      }),
    ]);
    expect(frame.coversThroughSeq).toBe(ack.throughSeq);
    expect(frame.revision).toBe(decodeFrameDelta(frame.payload.frame).frameEpoch);
    expect(frame.payload.caret.frameEpoch).toBe(frame.revision);
    return { ...ack, frame };
  } finally {
    if (timeout !== undefined) clearTimeout(timeout);
    unsubscribe();
  }
}

function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages.flatMap(({ primitives }) => primitives)
    .map((primitive) => primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : '').join('');
}

test('worker frames include highlights for keyboard and point range selections', async () => {
  const { client, point } = await openWorker();
  try {
    const selected = await editWithFrame(client, intent(1, 0, 'select', {
      selection: { anchor: { ...point, offset: 0 }, focus: point },
    }));
    expect(selected.frame!.payload.selectionRects?.length).toBeGreaterThan(0);
    const rect = selected.frame!.payload.selectionRects![0]!;
    const collapsed = await editWithFrame(client, intent(2, 0, 'navigate', { unit: 'character', direction: 'forward', extend: false }));
    expect(collapsed.frame!.payload.selectionRects).toEqual([]);
    const keyboard = await editWithFrame(client, intent(3, 0, 'navigate', { unit: 'character', direction: 'backward', extend: true }));
    expect(keyboard.frame!.payload.selectionRects?.length).toBeGreaterThan(0);
    const pointer = await editWithFrame(client, intent(4, 0, 'extendSelection', {
      page: { pageId: '1', pageIndex: rect.pageIndex }, x: rect.x, y: rect.y + rect.height / 2,
      layoutRevision: keyboard.frame!.payload.layoutRevision, frameEpoch: keyboard.frame!.revision, clickCount: 1,
      modifiers: { shift: true, alt: false, ctrl: false, meta: false },
    }));
    expect(pointer.frame!.payload.selectionRects?.length).toBeGreaterThan(0);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker frames highlight one rect per cell of a selected table', async () => {
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('');
  const { client, engine } = await openWorker({ bodyXml:
    '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p>' +
    `<w:tbl><w:tblPr><w:tblW w:w="5760" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>` +
    `<w:tr>${cell('A')}${cell('B')}</w:tr><w:tr>${cell('C')}${cell('D')}</w:tr></w:tbl>`,
  });
  try {
    const story = engine.geometryReader.storyIds().find((candidate) => engine.paragraphs(candidate)[0]?.text === 'A');
    if (!story) throw new Error('Cell story missing');
    const paraId = engine.paragraphs(story)[0]!.paraId;
    const at = { story, paraId, offset: 0 };
    await client.edit(intent(1, 0, 'select', { selection: { anchor: at, focus: at } }));
    const tabled = await editWithFrame(client, intent(2, 0, 'selectTableCells', { target: 'table' }));
    expect(tabled.outcomes[0]?.status).toBe('noop');
    const rects = tabled.frame!.payload.selectionRects!;
    expect(rects.length).toBe(4);
    for (const rect of rects) {
      expect(rect.pageIndex).toBe(0);
      expect(rect.x).toBeGreaterThanOrEqual(0);
      expect(rect.y).toBeGreaterThanOrEqual(0);
      expect(rect.width).toBeGreaterThan(0);
      expect(rect.height).toBeGreaterThan(0);
    }
    expect(new Set(rects.map((rect) => `${rect.x},${rect.y}`)).size).toBe(4);
    const first = rects[0]!;
    const clicked = await editWithFrame(client, intent(3, 0, 'selectPoint', {
      page: { pageId: '1', pageIndex: first.pageIndex }, x: first.x + first.width / 2, y: first.y + first.height / 2,
      layoutRevision: tabled.frame!.payload.layoutRevision, frameEpoch: tabled.frame!.revision,
      clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
    }));
    expect(clicked.frame!.payload.selectionRects).toEqual([]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker selection reads project tracked deletions before slicing raw offsets', async () => {
  const { client, point } = await openWorker({ bodyXml:
    '<w:p w14:paraId="00000001"><w:del w:id="1" w:author="Writer" w:date="2026-01-01T00:00:00Z">' +
    '<w:r><w:delText>XX</w:delText></w:r></w:del><w:r><w:t>abc</w:t></w:r></w:p>',
  });
  try {
    const result = await client.documentRead({ kind: 'rawSelectionInfo', selection: {
      anchor: { ...point, offset: 2 }, focus: { ...point, offset: 3 },
    } });
    expect(result.value).toMatchObject({ selectedText: 'a', before: '', after: 'bc', paragraphText: 'abc' });
  } finally { client.destroy(); }
}, TIMEOUT);

test('render environment commands change hidden text and theme colors without editing the document', async () => {
  const { client, engine, initial } = await openWorker({ bodyXml:
    '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r>' +
    '<w:r><w:rPr><w:vanish/><w:color w:themeColor="accent1"/></w:rPr><w:t>Hidden</w:t></w:r></w:p>',
  });
  try {
    const nativeVersion = engine.proposalEngine.version();
    const layouts = spyOn(engine, 'layoutDocumentWithRegionsAfterEdit');
    const command = intent(1, 0, 'setRenderEnvironment', {
      themeColors: { accent1: 'FF0000' }, showHiddenText: true,
    } as never);
    const ack = await client.edit(command);
    expect(ack.outcomes[0]?.status).toBe('noop');
    expect(ack.version).toBe(0);
    expect(engine.proposalEngine.version()).toBe(nativeVersion);
    expect(ack.selection ?? initial.selection).toEqual(initial.selection);
    expect(layouts).toHaveBeenCalledTimes(1);
    const initialFrame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame!.payload.frame));
    expect(frameText(initialFrame)).toBe('Seed');
    const frame = applyFrameDeltaOwned(initialFrame, decodeFrameDelta(ack.frame!.payload.frame));
    expect(frameText(frame)).toBe('SeedHidden');
    const coloredText = frame.pages[0]!.page.primitives.flatMap((primitive) =>
      (primitive.kind === 'text' || primitive.kind === 'glyphRun') && primitive.color.toUpperCase() === '#FF0000' ? [primitive.text] : []
    ).join('');
    expect(coloredText).toBe('Hidden');
    await client.edit(command);
    expect(layouts).toHaveBeenCalledTimes(1);
    layouts.mockRestore();
  } finally { client.destroy(); }
}, TIMEOUT);

test('render environment reservations lay out at their document version and never reach the document', async () => {
  const { client, engine, initial, point } = await openWorker({ bodyXml:
    '<w:p w14:paraId="00000001"><w:r><w:t>Alpha beta</w:t></w:r></w:p>' +
    '<w:p w14:paraId="00000002"><w:r><w:t>Gamma</w:t></w:r></w:p>',
  });
  try {
    const at = (offset: number) => ({ paraId: point.paraId, offset });
    const inline = (version: string, width: number) => ({
      id: 'box', kind: 'inline', width, height: 30,
      anchor: { kind: 'range', version, range: { story: 'body', start: at(6), end: at(6), view: 'accepted' } },
    });
    const environment = (reservations: unknown) => ({ themeColors: {}, showHiddenText: false, reservations } as never);
    const nativeVersion = engine.proposalEngine.version();
    const before = new Uint8Array((await client.save({ comments: [] })).bytes);
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame!.payload.frame));
    const boxes = () => frame.displayList.pages.flatMap((page) => page.reservations ?? []);
    expect(boxes()).toEqual([]);

    const set = await client.edit(intent(1, 0, 'setRenderEnvironment', environment({
      version: nativeVersion, sets: [{ owner: 'plugin', items: [inline(nativeVersion, 40)] }],
    })));
    expect(set.outcomes[0]?.status).toBe('noop');
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(set.frame!.payload.frame));
    expect(boxes()).toEqual([expect.objectContaining({ owner: 'plugin', id: 'box', width: 40, height: 30 })]);
    expect(frameText(frame)).toBe('Alpha betaGamma');
    expect(engine.proposalEngine.version()).toBe(nativeVersion);
    expect(new Uint8Array((await client.save({ comments: [] })).bytes)).toEqual(before);

    const stale = await client.edit(intent(2, 0, 'setRenderEnvironment', environment({
      version: '5', sets: [{ owner: 'plugin', items: [inline('5', 80)] }],
    })));
    expect(stale.outcomes[0]?.status).toBe('noop');
    if (stale.frame) frame = applyFrameDeltaOwned(frame, decodeFrameDelta(stale.frame.payload.frame));
    expect(boxes()).toEqual([expect.objectContaining({ id: 'box', width: 40 })]);

    const invalid = await client.edit(intent(3, 0, 'setRenderEnvironment', environment({
      version: nativeVersion, sets: [{ owner: 'plugin', items: [inline(nativeVersion, -1)] }],
    })));
    expect(invalid.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    if (invalid.frame) frame = applyFrameDeltaOwned(frame, decodeFrameDelta(invalid.frame.payload.frame));

    const locked = await client.edit(intent(4, 0, 'setReadOnly', { readOnly: true }));
    expect(locked.outcomes[0]?.status).toBe('noop');
    if (locked.frame) frame = applyFrameDeltaOwned(frame, decodeFrameDelta(locked.frame.payload.frame));
    const readOnly = await client.edit(intent(5, 0, 'setRenderEnvironment', environment({
      version: nativeVersion, sets: [{ owner: 'plugin', items: [{
        id: 'gap', kind: 'block', side: 'after', height: 50,
        anchor: { kind: 'paragraph', paragraph: {
          kind: 'session', sessionId: engine.paragraphIdentities().sessionId, story: 'body', paraId: point.paraId,
        } },
      }] }],
    })));
    expect(readOnly.outcomes[0]?.status).toBe('noop');
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(readOnly.frame!.payload.frame));
    expect(boxes()).toEqual([expect.objectContaining({ id: 'gap', height: 50 })]);

    const cleared = await client.edit(intent(6, 0, 'setRenderEnvironment', environment({ version: nativeVersion, sets: [] })));
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(cleared.frame!.payload.frame));
    expect(boxes()).toEqual([]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('stale render environment reservations retain listed owners and clear removed owners after an edit', async () => {
  const { client, engine, initial, point } = await openWorker();
  const layouts = spyOn(engine, 'layoutDocumentWithRegionsAfterEdit');
  try {
    const version = initial.frame!.payload.documentVersion!;
    const inline = (id: string, width: number, hideAnchor = false) => ({
      id, kind: 'inline', width, height: 20, hideAnchor,
      anchor: { kind: 'range', version, range: {
        story: 'body', view: 'accepted',
        start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 2 },
      } },
    });
    const environment = (sets: unknown[]) => ({
      themeColors: {}, showHiddenText: false, reservations: { version, sets },
    } as never);
    const applied = () => JSON.parse(layouts.mock.calls.at(-1)![0]).renderEnv.reservations;
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame!.payload.frame));
    const set = await client.edit(intent(1, 0, 'setRenderEnvironment', environment([
      { owner: 'A', items: [inline('a', 35, true)] },
      { owner: 'B', items: [inline('b', 20)] },
    ])));
    expect(set.outcomes[0]?.status).toBe('noop');
    expect(Object.values(applied())).toHaveLength(2);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(set.frame!.payload.frame));
    expect(frameText(frame)).toBe('ed');
    const retained = applied()[JSON.stringify(['A', 'a'])];

    const edited = await client.edit(intent(2, 0, 'insertText', { text: '!' }));
    expect(edited.outcomes[0]?.status).toBe('applied');
    expect(edited.version).toBe(1);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(edited.frame!.payload.frame));

    const stale = await client.edit(intent(3, 1, 'setRenderEnvironment', environment([
      { owner: 'A', items: [inline('a', 80)] },
    ])));
    expect(stale.outcomes[0]?.status).toBe('noop');
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(stale.frame!.payload.frame));
    expect(applied()).toEqual({ [JSON.stringify(['A', 'a'])]: retained });
    expect(frame.displayList.pages.flatMap((page) => page.reservations ?? []))
      .toEqual([expect.objectContaining({ owner: 'A', id: 'a', width: 35 })]);
    expect(frameText(frame)).toBe('ed!');

    const cleared = await client.edit(intent(4, 1, 'setRenderEnvironment', environment([])));
    expect(cleared.outcomes[0]?.status).toBe('noop');
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(cleared.frame!.payload.frame));
    expect(applied()).toEqual({});
    expect(frame.displayList.pages.flatMap((page) => page.reservations ?? [])).toEqual([]);
    expect(frameText(frame)).toBe('Seed!');
  } finally {
    layouts.mockRestore();
    client.destroy();
  }
}, TIMEOUT);

test('frames name the reservation set their layout includes, or none current for a stale set', async () => {
  const { client, initial, point } = await openWorker();
  try {
    const environment = (version: string, key: string) => ({ themeColors: {}, showHiddenText: false, reservations: { version, key, sets: [{ owner: 'A', items: [{
      id: 'a', kind: 'inline', width: 30, height: 20, anchor: { kind: 'range', version, range: {
        story: 'body', view: 'accepted', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 0 } } },
    }] }] } } as never);
    expect(initial.frame!.payload.reservations).toBeUndefined();
    const opened = initial.frame!.payload.documentVersion!;
    const set = await client.edit(intent(1, 0, 'setRenderEnvironment', environment(opened, 'k1')));
    expect(set.frame!.payload.reservations).toEqual({ version: opened, key: 'k1' });
    const edited = await client.edit(intent(2, 0, 'insertText', { text: '!' }));
    expect(edited.frame!.payload.reservations).toEqual({ version: opened, key: 'k1' });
    const stale = await client.edit(intent(3, 1, 'setRenderEnvironment', environment(opened, 'k2')));
    expect(stale.frame!.payload.reservations).toEqual({ version: opened, key: null });
    const changed = edited.frame!.payload.documentVersion!;
    const current = await client.edit(intent(4, 1, 'setRenderEnvironment', environment(changed, 'k3')));
    expect(current.frame!.payload.reservations).toEqual({ version: changed, key: 'k3' });
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('a render environment command without reservations keeps the set its frames name', async () => {
  const { client, initial, point } = await openWorker();
  try {
    const version = initial.frame!.payload.documentVersion!;
    const set = await client.edit(intent(1, 0, 'setRenderEnvironment', { themeColors: {}, showHiddenText: false, reservations: {
      version, key: 'k1', sets: [{ owner: 'A', items: [{ id: 'a', kind: 'inline', width: 30, height: 20, anchor: { kind: 'range', version,
        range: { story: 'body', view: 'accepted', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 0 } } } }] }],
    } } as never));
    expect(set.frame!.payload.reservations).toEqual({ version, key: 'k1' });
    const themed = await client.edit(intent(2, 0, 'setRenderEnvironment', { themeColors: { accent1: 'FF0000' }, showHiddenText: true } as never));
    expect(themed.outcomes[0]?.status).toBe('noop');
    expect(themed.frame!.payload.reservations).toEqual({ version, key: 'k1' });
    const frame = [set, themed].reduce((retained, ack) => applyFrameDeltaOwned(retained, decodeFrameDelta(ack.frame!.payload.frame)),
      applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame!.payload.frame)));
    expect(frame.displayList.pages.flatMap((page) => page.reservations ?? [])).toEqual([expect.objectContaining({ owner: 'A', id: 'a', width: 30 })]);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test.each([true, false])('render environment reservations survive a review preview relayout (regions=%s)', async (layoutWithRegions) => {
  const { client, initial, point } = await openWorker({ layoutWithRegions });
  try {
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame!.payload.frame));
    const boxes = () => frame.displayList.pages.flatMap((page) => page.reservations ?? []);
    const version = initial.frame!.payload.documentVersion!;
    const set = await client.edit(intent(1, 0, 'setRenderEnvironment', {
      themeColors: {}, showHiddenText: false, reservations: { version, sets: [{ owner: 'plugin', items: [{
        id: 'box', kind: 'inline', width: 40, height: 20,
        anchor: { kind: 'range', version, range: {
          story: 'body', view: 'accepted',
          start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 0 },
        } },
      }] }] },
    } as never));
    expect(set.outcomes[0]?.status).toBe('noop');
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(set.frame!.payload.frame));
    expect(boxes()).toEqual([expect.objectContaining({ owner: 'plugin', id: 'box', width: 40 })]);

    const ids = await client.documentRead({ kind: 'paragraphIdentities' });
    const proposed = await client.edit(intent(2, 0, 'propose', { expectVersion: ids.version, proposals: [{
      id: 'one', op: 'replaceText', paragraph: ids.value.paragraphs[0]!.session!, search: 'Seed', replaceWith: 'Bud',
      suggest: { author: 'Ada', date: '2026-01-01' },
    }] }));
    expect(proposed.outcomes[0]?.status).toBe('applied');
    frame = applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame));
    expect(boxes()).toEqual([expect.objectContaining({ owner: 'plugin', id: 'box', width: 40 })]);

    const review = await client.documentRead({ kind: 'reviewSnapshot' });
    const previewed = await client.edit(intent(3, 1, 'setProposalStates', {
      expectVersion: review.version, expectPreviewVersion: review.value!.proposals.previewVersion,
      changes: [{ id: 'one', state: 'accepted' }],
    }));
    expect(previewed.outcomes[0]?.status).toBe('applied');
    frame = applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame));
    expect(frameText(frame)).toBe('Bud');
    expect(boxes()).toEqual([expect.objectContaining({ owner: 'plugin', id: 'box', width: 40 })]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('the worker and a main-thread session lay out one reservation set identically', async () => {
  const bodyXml = '<w:p w14:paraId="00000001"><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p>' +
    '<w:p w14:paraId="00000002"><w:r><w:t>Next</w:t></w:r></w:p>';
  const { client, engine, init, initial, point } = await openWorker({ bodyXml });
  const host = await createYrsSession({ clientId: 8302 });
  try {
    host.seedFromDocx(syntheticDocument(false, false, false, bodyXml));
    const items = (sessionId: string): DocxReservation[] => {
      const paragraph = { kind: 'session' as const, sessionId, story: 'body', paraId: point.paraId };
      return [
        { id: 'chip', kind: 'inline', width: 40, height: 24, anchor: { kind: 'search', paragraph, text: 'beta' } },
        { id: 'gap', kind: 'block', side: 'after', height: 50, anchor: { kind: 'paragraph', paragraph } },
        { id: 'cut', kind: 'hide', anchor: { kind: 'search', paragraph, text: 'gamma' } },
      ];
    };
    const set = await client.edit(intent(1, 0, 'setRenderEnvironment', {
      themeColors: {}, showHiddenText: false,
      reservations: { version: initial.frame!.payload.documentVersion!, sets: [{ owner: 'plugin', items: items(engine.paragraphIdentities().sessionId) }] },
    } as never));
    const worker = applyFrameDeltaOwned(
      applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame!.payload.frame)),
      decodeFrameDelta(set.frame!.payload.frame)
    );

    const font = host.registerFont(new Uint8Array(readFileSync(resolve(
      process.cwd(), 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
    ))));
    const reservations = resolveReservations(host, {
      version: host.version(),
      sets: [{ owner: 'plugin', items: items(host.paragraphIdentities().sessionId) }],
    });
    const request = JSON.parse(init.layoutInput);
    request.measurement.fontChains = { 'liberation sans|0|0': [font] };
    request.renderEnv = { ...request.renderEnv, reservations };
    host.layoutDocumentWithRegionsRetainedJson(JSON.stringify(request));
    const main = applyFrameDeltaOwned(null, decodeFrameDelta(host.buildDisplayListFrame(JSON.stringify({
      fontChains: { 'liberation sans|0|0': [font], 'calibri|0|0': [font] },
    }), 0)));

    const shape = (frame: RetainedFrame) => frame.displayList.pages.map((page) => ({
      reservations: page.reservations,
      text: page.primitives.flatMap((primitive) => primitive.kind === 'glyphRun' || primitive.kind === 'text'
        ? [[primitive.text, Math.round(primitive.kind === 'text' ? primitive.x : primitive.glyphs[0]?.x ?? 0)]] : []),
    }));
    expect(shape(worker)[0]!.reservations?.map(({ id }) => id).sort()).toEqual(['chip', 'gap']);
    expect(frameText(worker)).toBe('Alpha beta Next');
    expect(shape(main)).toEqual(shape(worker));
  } finally {
    host.destroy();
    client.destroy();
  }
}, TIMEOUT);

function expectFrameRegions(frame: RetainedFrame, headerText: string, footerText: string): void {
  expect(frameText(frame)).toBe('Seed');
  expect(frame.displayList.pages.length).toBeGreaterThan(0);
  for (const page of frame.displayList.pages) {
    expect(page.primitives.length).toBeGreaterThan(0);
    expect(page.header?.kind).toBe('header');
    expect(page.footer?.kind).toBe('footer');
    expect(page.header?.primitives.length).toBeGreaterThan(0);
    expect(page.footer?.primitives.length).toBeGreaterThan(0);
    for (const [region, text] of [[page.header!, headerText], [page.footer!, footerText]] as const) {
      expect(region.primitives.map((primitive) => primitive.kind === 'glyphRun' || primitive.kind === 'text'
        ? primitive.text : '').join('')).toBe(text);
    }
  }
}

test('d2c_worker_fixture_clones_staged_page_and_background_bitmaps', async () => {
  const { client, initial } = await openWorker();
  try {
    const inserted = await client.edit(intent(1, 0, 'insertText', { text: '!' }));
    const full = await client.requestFullFrame();
    for (const frame of [initial.frame, inserted.frame!, full.frame!]) {
      const staged = frame.payload.stagedPresentation;
      const echo = frame.payload.echoContext;
      expect(staged?.frameEpoch).toBe(frame.revision);
      expect(staged?.pages.length).toBeGreaterThan(0);
      expect(echo).toBeDefined();
      const page = staged!.pages.find(({ page }) => page.pageId === echo!.page.pageId)!;
      expect(page.frameEpoch).toBe(frame.revision);
      expect(page.bitmap.width).toBeGreaterThan(0);
      expect(page.bitmap.height).toBeGreaterThan(0);
      expect(page.backgrounds?.map(({ id, rect }) => ({ id, rect }))).toEqual(
        echo!.damage.map(({ backgroundId, rect }) => ({ id: backgroundId, rect })));
      for (const bitmap of [page.bitmap, ...page.backgrounds!.map(({ bitmap }) => bitmap)]) {
        expect(typeof bitmap.close).toBe('function');
        expect(bitmap.width).toBeGreaterThan(0);
        expect(bitmap.height).toBeGreaterThan(0);
        bitmap.close();
        expect([bitmap.width, bitmap.height]).toEqual([0, 0]);
        bitmap.close();
        expect([bitmap.width, bitmap.height]).toEqual([0, 0]);
      }
    }
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('d1a_real_worker_init_insert_delete_acknowledge_current_layout_and_state', async () => {
  const { client, engine, initial, point } = await openWorker();
  const layouts = spyOn(engine, 'layoutDocumentWithRegionsAfterEdit');
  try {
    expect(initial.selection).toMatchObject({ anchor: point, focus: point });
    expect(initial.editState).toMatchObject({ readOnly: false, canUndo: false });
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    expect(frameText(frame)).toBe('Seed');
    const inserted = await client.edit(intent(1, 0, 'insertText', { text: '!' }));
    expect(inserted).toMatchObject({
      throughSeq: 1, version: 1, outcomes: [{ seq: 1, status: 'applied' }],
      selection: { anchor: { ...point, offset: 5 }, focus: { ...point, offset: 5 } },
      editState: { canUndo: true, canRedo: false }, frame: { coversThroughSeq: 1 },
    });
    expect(inserted.frame).toBeDefined();
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(inserted.frame!.payload.frame));
    expect(frameText(frame)).toBe('Seed!');
    expect(layouts).toHaveBeenCalledTimes(1);
    const deleted = await client.edit(intent(2, 1, 'deleteBackward', {}));
    expect(deleted).toMatchObject({
      throughSeq: 2, version: 2, outcomes: [{ seq: 2, status: 'applied' }],
      selection: { anchor: point, focus: point }, frame: { coversThroughSeq: 2 },
    });
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(deleted.frame!.payload.frame));
    expect(frameText(frame)).toBe('Seed');
    expect(layouts).toHaveBeenCalledTimes(2);
    const toggled = await client.edit(intent(3, 2, 'toggleMark', { mark: { type: 'bold' } }));
    expect(toggled).toMatchObject({ version: 2, outcomes: [{ status: 'noop' }], editState: {
      storedMarks: { delta: { bold: true } },
    } });
    const refused = await client.edit(intent(4, 0, 'insertText', { text: 'stale' }));
    expect(refused).toMatchObject({ version: 2, outcomes: [{ status: 'refused', reason: 'stale-anchor' }] });
    expect(layouts).toHaveBeenCalledTimes(2);
  } finally {
    layouts.mockRestore();
    client.destroy();
  }
}, TIMEOUT);

test('d1a_worker_keeps_sticky_provenance_until_document_epoch_replacement', async () => {
  const { client, worker, engine, init, initial, point } = await openWorker();
  try {
    const sticky = structuredClone(initial.selection!);
    const checkpoint = await client.checkpoint();
    await expect(client.commandInit(init)).rejects.toThrow('already initialized');
    const atStart = { ...point, offset: 0 };
    expect((await client.edit(intent(1, 0, 'insertText', {
      text: '-', selection: { anchor: atStart, focus: atStart },
    }))).outcomes[0]?.status).toBe('applied');
    await client.requestFullFrame();
    await client.checkpoint();
    const anchored = await client.edit(intent(2, 0, 'insertText', { text: '!', selection: sticky }));
    expect(anchored.outcomes[0]?.status).toBe('applied');
    expect(anchored.selection?.focus.offset).toBe(6);
    expect(worker.sessions).toHaveLength(1);
    const replacement = await client.bootstrap({
      clientId: 8302, state: checkpoint.state, selection: null,
      fonts: init.fonts, fontsRevision: 0, renderInputs: [], measureInputs: [],
      layoutInput: LAYOUT, layoutWithRegions: true, layoutRevision: 2,
    }, '{}');
    expect(worker.sessions).toHaveLength(2);
    expect(worker.sessions[1]).not.toBe(engine);
    expect(replacement.commandInitial).toMatchObject({ epoch: 2, version: 0, throughSeq: 0 });
    const current = worker.sessions[1]!;
    const before = current.encodeStateVector();
    expect(current.resolveStickyPosition({ story: point.story, encoded: sticky.focus.sticky! })).not.toBeNull();
    const oldEpoch = await client.edit(intent(3, 0, 'insertText', { text: 'old', selection: sticky }));
    expect(oldEpoch.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-epoch' });
    const oldSticky = await client.edit({
      ...intent(1, 0, 'insertText', { text: 'old', selection: sticky }), basis: { epoch: 2, version: 0 },
    });
    expect(oldSticky.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-anchor' });
    expect(current.encodeStateVector()).toEqual(before);
    expect(oldSticky.version).toBe(0);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test.each([true, false])('d1a_header_footer_command_frames_use_current_story_content (regions=%s)', async (layoutWithRegions) => {
  const { client, engine, initial, init } = await openWorker({ headersFooters: true, layoutWithRegions });
  try {
    expect(init.extras).toContain('H');
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    expectFrameRegions(frame, 'H', 'F');
    for (const [index, text] of ['H', 'F'].entries()) {
      const story = engine.geometryReader.storyIds().find((story) =>
        story !== 'body' && engine.paragraphs(story)[0]?.text === text
      );
      if (!story) throw new Error(`Missing ${text} story`);
      const at = { story, paraId: engine.paragraphs(story)[0]!.paraId, offset: 1 };
      const edited = await client.edit(intent(index + 1, index, 'insertText', {
        text: '!', selection: { anchor: at, focus: at },
      }));
      expect(edited.outcomes[0]?.status).toBe('applied');
      expect(edited.frame?.coversThroughSeq).toBe(index + 1);
      frame = applyFrameDeltaOwned(frame, decodeFrameDelta(edited.frame!.payload.frame));
      expectFrameRegions(frame, 'H!', index === 0 ? 'F' : 'F!');
      expect(engine.paragraphs(story)[0]?.text).toBe(`${text}!`);
    }
    const full = await client.requestFullFrame();
    frame = applyFrameDeltaOwned(null, decodeFrameDelta(full.frame!.payload.frame));
    expectFrameRegions(frame, 'H!', 'F!');
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('d1a_non_region_command_frames_refresh_measured_content', async () => {
  const { client, engine, initial } = await openWorker({ layoutWithRegions: false });
  const layouts = spyOn(engine, 'layoutDocumentJson');
  try {
    let frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    expect(frameText(frame)).toBe('Seed');
    const inserted = await client.edit(intent(1, 0, 'insertText', { text: '!' }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(inserted.frame?.coversThroughSeq).toBe(1);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(inserted.frame!.payload.frame));
    expect(frameText(frame)).toBe('Seed!');
    expect(layouts.mock.results[0]?.value).toBe('');
    expect(engine.paragraphs('body')[0]?.text).toBe('Seed!');
    const deleted = await client.edit(intent(2, 1, 'deleteBackward', {}));
    expect(deleted.outcomes[0]?.status).toBe('applied');
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(deleted.frame!.payload.frame));
    expect(frameText(frame)).toBe('Seed');
    const full = await client.requestFullFrame();
    expect(frameText(applyFrameDeltaOwned(null, decodeFrameDelta(full.frame!.payload.frame)))).toBe('Seed');
  } finally {
    layouts.mockRestore();
    client.destroy();
  }
}, TIMEOUT);

test('d1a_worker_read_sticky_anchors_edit_and_rebase_after_remote_update', async () => {
  const { client, engine, point } = await openWorker();
  const remote = await createResidentEngineSession(undefined, 8302);
  try {
    const readSelection = async (offset: number) => {
      const at = { ...point, offset };
      const read = await client.documentRead({
        kind: 'stickyAnchors', locs: [at, at], version: engine.proposalEngine.version(),
      });
      const [anchor, focus] = read.value;
      expect(anchor).not.toBeNull();
      expect(focus).not.toBeNull();
      return {
        anchor: { ...at, sticky: anchor!.encoded },
        focus: { ...at, sticky: focus!.encoded },
      };
    };
    const inserted = await client.edit(intent(1, 0, 'insertText', {
      text: '!', selection: await readSelection(1),
    }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]?.text).toBe('S!eed');
    const sticky = await readSelection(3);
    remote.loadState(engine.encodeState());
    const vector = engine.encodeStateVector();
    remote.insertText({ ...point, offset: 0 }, '--');
    engine.applyUpdate(remote.encodeStateAsUpdate(vector));
    const rebased = await client.edit(intent(2, 0, 'insertText', { text: '?', selection: sticky }));
    expect(rebased.outcomes[0]?.status).toBe('applied');
    expect(rebased.selection?.focus.offset).toBe(6);
    expect(engine.paragraphs('body')[0]?.text).toBe('--S!e?ed');
    const full = await client.requestFullFrame();
    expect(frameText(applyFrameDeltaOwned(null, decodeFrameDelta(full.frame!.payload.frame)))).toBe('--S!e?ed');
  } finally {
    remote.destroy();
    client.destroy();
  }
}, TIMEOUT);

test('history_commands_order_typing_undo_redo_and_readonly_without_duplicates', async () => {
  const { client, engine } = await openWorker();
  try {
    const typing = client.edit(intent(1, 0, 'insertText', { text: '!' }));
    const undo = client.edit({ ...intent(2, 0, 'undo', {}), afterSeq: 1 });
    expect((await typing).outcomes[0]?.status).toBe('applied');
    const undone = await undo;
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]?.text).toBe('Seed');
    expect(undone.editState).toMatchObject({ canUndo: false, canRedo: true });
    const duplicate = await client.edit(intent(2, 0, 'undo', {}));
    expect(duplicate.outcomes).toEqual(undone.outcomes);
    expect(frameText(applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame)))).toBe('Seed');
    const redone = await client.edit(intent(3, 2, 'redo', {}));
    expect(redone.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]?.text).toBe('Seed!');
    expect(redone.editState).toMatchObject({ canUndo: true, canRedo: false });
    expect(frameText(applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame)))).toBe('Seed!');
    await client.edit(intent(4, 3, 'setReadOnly', { readOnly: true }));
    const blocked = await client.edit(intent(5, 3, 'undo', {}));
    expect(blocked.outcomes[0]).toMatchObject({ status: 'refused', reason: 'read-only' });
    expect(engine.paragraphs('body')[0]?.text).toBe('Seed!');
  } finally { client.destroy(); }
}, TIMEOUT);

test('replace_matches_and_range_highlights_share_native_version_and_layout_identity', async () => {
  const { client, engine, point, initial } = await openWorker();
  try {
    const range = { story: 'body', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 4 } };
    const version = engine.proposalEngine.version();
    const highlights = await client.documentReadAt({ kind: 'rangeHighlights', ranges: [range], expectVersion: version }, version);
    expect(highlights.status).toBe('ok');
    if (highlights.status !== 'ok') throw new Error('Highlight read was superseded');
    const geometry = highlights.value!;
    expect(geometry).toMatchObject({ version, layoutRevision: 1, frameEpoch: initial.frame!.revision });
    expect(geometry.rects[0].length).toBeGreaterThan(0);
    expect(geometry.rects[0][0].pageId).toBeDefined();
    const find = await client.documentRead({ kind: 'findText', request: { text: 'Seed', within: { kind: 'story', story: 'body' }, view: 'accepted' } });
    if (!find.value.ok) throw new Error(find.value.failure.message);
    const replaced = await client.edit(intent(1, 0, 'replaceMatches', { targets: find.value.matches.map(match => match.range), text: 'Bud', expectVersion: find.version }));
    expect(replaced.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]?.text).toBe('Bud');
    expect(await client.documentReadAt({ kind: 'rangeHighlights', ranges: [range], expectVersion: version }, version)).toEqual({ status: 'superseded' });
    const current = engine.proposalEngine.version();
    expect((await client.documentReadAt({ kind: 'rangeHighlights', ranges: [], expectVersion: current, expectFrameEpoch: initial.frame.revision }, current))).toMatchObject({ status: 'ok', value: null });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker batch validation and ordered edits return receipts, refuse stale requests, and undo once', async () => {
  const { client, initial, point } = await openWorker();
  try {
    const token = (await client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } })).version;
    const request = { expectVersion: token, steps: [{ op: 'insertText' as const, target: { kind: 'paragraph' as const, story: 'body', paraId: point.paraId }, at: 'end' as const, text: ' plugin' }] };
    expect((await client.documentRead({ kind: 'validateEdits', request })).value).toMatchObject({ ok: true, wouldApply: true });
    const edit = intent(1, initial.version, 'applyEdits', { request });
    const applied = await client.edit(edit);
    expect(applied.outcomes[0]).toMatchObject({ status: 'applied' });
    expect(applied.editState?.batch?.result).toMatchObject({ ok: true, applied: true, receipts: [{ changed: true }] });
    expect((await client.edit(edit)).editState?.batch?.result).toEqual(applied.editState?.batch?.result);
    expect((await client.edit(intent(2, applied.version, 'applyEdits', { request }))).outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-version' });
    const undo = await client.edit(intent(3, applied.version, 'undo', {}));
    expect(undo.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } })).value).toMatchObject({ paragraphs: [{ text: 'Seed' }] });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker creates opens edits and removes parts through ordered commands', async () => {
  const { client, initial } = await openWorker();
  try {
    const created = await client.edit(intent(1, 0, 'ensurePart', { kind: 'header', section: 'final', page: 0, expectVersion: initial.frame.payload.documentVersion! }));
    expect(created).toMatchObject({ version: 1, outcomes: [{ status: 'applied' }], editState: { activePart: { kind: 'header' } } });
    expect(created.selection!.focus.story).toMatch(/^hf:/);
    const typed = await client.edit({ ...intent(2, 1, 'insertText', { text: 'Header' }), afterSeq: 1 });
    expect(typed.outcomes[0]?.status).toBe('applied');
    expect(typed.frame!.payload.parts!.some(part => part.kind === 'header' && part.story === created.selection!.focus.story)).toBe(true);
    const { parseDocx } = await import('../docx/parser');
    const saved = await client.save({ comments: [] });
    const reopened = await parseDocx(saved.bytes);
    expect(JSON.stringify([...reopened.package.headers!.values()][0]!.content)).toContain('"text":"Header"');
    const removed = await client.edit(intent(3, 2, 'removePart', { kind: 'header', section: 'final', page: 0, expectVersion: saved.version }));
    expect(removed).toMatchObject({ version: 3, outcomes: [{ status: 'applied' }], selection: { focus: { story: 'body' } } });
    expect((await client.edit(intent(4, 3, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    const opened = await client.edit(intent(5, 4, 'openPart', { kind: 'header', section: 'final', page: 0, expectVersion: (await client.documentRead({ kind: 'partSnapshot' })).version }));
    expect(opened.outcomes[0]?.status).toBe('noop');
    expect(opened.selection!.focus.story).toBe(created.selection!.focus.story);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker part creation is undoable while its story is active', async () => {
  const { client, initial } = await openWorker();
  try {
    await client.edit(intent(1, 0, 'ensurePart', { kind: 'footer', section: 'final', page: 0, expectVersion: initial.frame.payload.documentVersion! }));
    const undone = await client.edit(intent(2, 1, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(undone.selection!.focus.story).toBe('body');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker note parts edit and remove their references in the same history step', async () => {
  const { client, initial, engine } = await openWorker({ notes: true });
  try {
    const opened = await client.edit(intent(1, 0, 'openPart', { kind: 'footnote', noteId: 1, section: 'final', page: 0, expectVersion: initial.frame.payload.documentVersion! }));
    expect(opened).toMatchObject({ outcomes: [{ status: 'noop' }], selection: { focus: { story: 'fn:1' } } });
    const typed = await client.edit({ ...intent(2, 0, 'insertText', { text: 'Edited ' }), afterSeq: 1 });
    expect(typed.outcomes[0]?.status).toBe('applied');
    expect(typed.frame!.payload.parts!.some(part => part.story === 'fn:1' && part.height > 0)).toBe(true);
    const removed = await client.edit(intent(3, 1, 'removePart', { kind: 'footnote', noteId: 1, section: 'final', page: 0, expectVersion: typed.frame!.payload.documentVersion! }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body').some(part => part.kind === 'embed' && part.embedKind === 'noteRef')).toBe(false);
    expect((await client.edit(intent(4, 2, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect(engine.geometryReader.hasStory('fn:1')).toBe(true);
    expect(engine.storySegments('body').some(part => part.kind === 'embed' && part.embedKind === 'noteRef')).toBe(true);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker handles order control values and repeating item lifecycle with save and undo', async () => {
  const { client, initial } = await openWorker({ controls: true });
  try {
    const read = await client.documentReadAt({ kind: 'contentControlHandles' }, initial.frame.payload.documentVersion!);
    if (read.status !== 'ok') throw new Error('Control snapshot superseded');
    const checkbox = read.value.controls.find(control => control.control.tag === 'check')!;
    const item = read.value.controls.find(control => control.control.tag === 'item')!;
    expect(checkbox.handle).not.toBe(item.handle);
    expect(read.value.frameEpoch).toBe(initial.frame.payload.caret.frameEpoch);
    expect(checkbox.rects.length).toBeGreaterThan(0);
    expect(item.rects.length).toBeGreaterThan(0);
    expect(checkbox.widget).toEqual({ kind: 'checkbox', checked: false });
    const checked = await client.edit(intent(1, 0, 'setContentControlValue', {
      handle: checkbox.handle, value: { kind: 'checkbox', checked: true }, expectVersion: read.version,
    }));
    expect(checked.outcomes[0]?.status).toBe('applied');
    const added = await client.edit(intent(2, 1, 'addRepeatingItem', { handle: item.handle, expectVersion: checked.frame!.payload.documentVersion! }));
    expect(added.outcomes[0]?.status).toBe('applied');
    const after = await client.documentRead({ kind: 'contentControlHandles' });
    const items = after.value.controls.filter(control => control.control.tag === 'item');
    expect(items).toHaveLength(2);
    expect(new Set(items.map(control => control.handle)).size).toBe(2);
    expect(items.every(control => control.rects.length > 0)).toBe(true);
    expect(items.some(control => control.handle === item.handle)).toBe(true);
    const newItem = items.find(control => control.handle !== item.handle)!;
    const removed = await client.edit(intent(3, 2, 'removeRepeatingItem', { handle: newItem.handle, expectVersion: after.version }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect((await client.edit(intent(4, 3, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'contentControlHandles' })).value.controls.filter(control => control.control.tag === 'item')).toHaveLength(2);
    const { parseDocx } = await import('../docx/parser');
    const saved = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(JSON.stringify(saved.package.document)).toContain('Item text');
    const last = (await client.documentRead({ kind: 'contentControlHandles' })).version;
    const stale = await client.edit(intent(5, 4, 'setContentControlValue', { handle: checkbox.handle, value: { kind: 'checkbox', checked: false }, expectVersion: read.version }));
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-version' });
    expect((await client.documentRead({ kind: 'contentControlHandles' })).version).toBe(last);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker widgets of one shared-id paragraph each paint and save checkbox dropdown and date values in turn', async () => {
  const w15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
  const checkbox = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Flag"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>';
  const dropdown = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Choose option"/><w:dropDownList><w:listItem w:displayText="First" w:value="one"/><w:listItem w:displayText="Second" w:value="two"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>First</w:t></w:r></w:sdtContent></w:sdt>';
  const date = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Choose date"/><w:date w:fullDate="2026-10-12T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>2026-10-12</w:t></w:r></w:sdtContent></w:sdt>';
  const repeat = `<w:sdt><w:sdtPr><w:id w:val="8"/><w15:repeatingSection ${w15}/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"/><w:alias w:val="Line"/><w15:repeatingSectionItem ${w15}/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Item text</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>`;
  const { client } = await openWorker({ bodyXml: `<w:p w14:paraId="00000001">${checkbox}<w:r><w:t>gap</w:t></w:r>${checkbox}${dropdown}${date}</w:p>${repeat}` });
  try {
    const widgets = async () => {
      const read = await client.documentRead({ kind: 'contentControlHandles' });
      return { version: read.version, controls: read.value.controls.filter(control => control.widget) };
    };
    const summary = (controls: Awaited<ReturnType<typeof widgets>>['controls']) => controls.map(control => ({ alias: control.control.alias, kind: control.widget!.kind,
      painted: control.rects.length > 0, editable: control.control.effectiveLock.known && !control.control.effectiveLock.content && !control.control.dataBound }));
    const kinds = ['checkbox', 'checkbox', 'dropdown', 'date', 'repeatingItem'] as const;
    const expected: ReturnType<typeof summary> = ['Flag', 'Flag', 'Choose option', 'Choose date', 'Line'].map((alias, index) => ({ alias,
      kind: kinds[index]!, painted: true, editable: true }));
    let current = await widgets();
    expect(summary(current.controls)).toEqual(expected);
    const set = async (seq: number, alias: string, nth: number, value: YrsContentControlValue) => {
      const target = current.controls.filter(control => control.control.alias === alias)[nth]!;
      const applied = await client.edit(intent(seq, seq - 1, 'setContentControlValue', { handle: target.handle, value, expectVersion: current.version }));
      expect(applied.outcomes[0]).toMatchObject({ status: 'applied' });
      current = await widgets();
      expect(summary(current.controls)).toEqual(expected);
    };
    await set(1, 'Flag', 1, { kind: 'checkbox', checked: true });
    expect(current.controls.find(control => control.control.alias === 'Choose option')!.widget).toMatchObject({ kind: 'dropdown', value: 'one',
      items: [{ value: 'one', displayText: 'First' }, { value: 'two', displayText: 'Second' }] });
    await set(2, 'Choose option', 0, { kind: 'dropdown', value: 'two' });
    expect(current.controls.find(control => control.control.alias === 'Choose option')!.widget).toMatchObject({ value: 'two' });
    await set(3, 'Choose date', 0, { kind: 'date', date: '2026-11-03' });
    const { unzipContainer } = await import('../docx/wasm');
    const xml = new TextDecoder().decode(unzipContainer(new Uint8Array((await client.save({ comments: [] })).bytes))['word/document.xml']);
    expect([...xml.matchAll(/w14:checked w14:val="(\d)"/g)].map(match => match[1])).toEqual(['0', '1']);
    expect(xml).toContain('Second');
    expect(xml).toContain('2026-11-03');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker control handles survive unrelated edits and refuse stale missing and last-item writes', async () => {
  const { client, initial, point } = await openWorker({ controls: true });
  try {
    const before = await client.documentReadAt({ kind: 'contentControlHandles' }, initial.frame.payload.documentVersion!);
    if (before.status !== 'ok') throw new Error('Control snapshot unavailable');
    const item = before.value.controls.find(control => control.control.tag === 'item')!;
    const checkbox = before.value.controls.find(control => control.control.tag === 'check')!;
    const last = await client.edit(intent(1, 0, 'removeRepeatingItem', { handle: item.handle, expectVersion: before.version }));
    expect(last).toMatchObject({ version: 0, outcomes: [{ status: 'refused', reason: 'last-repeating-item' }] });
    const typed = await client.edit(intent(2, 0, 'insertText', { text: '+', selection: { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 0 } } }));
    expect(typed.outcomes[0]?.status).toBe('applied');
    const after = await client.documentRead({ kind: 'contentControlHandles' });
    expect(after.value.controls.find(control => control.control.tag === 'check')!.handle).toBe(checkbox.handle);
    expect(after.value.controls.find(control => control.control.tag === 'item')!.handle).toBe(item.handle);
    expect((await client.edit(intent(3, 1, 'setContentControlValue', { handle: checkbox.handle, value: { kind: 'checkbox', checked: true }, expectVersion: before.version }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'stale-version' });
    expect((await client.edit(intent(4, 1, 'setContentControlValue', { handle: 'missing', value: { kind: 'checkbox', checked: true }, expectVersion: after.version }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'stale-control' });
    await client.edit(intent(5, 1, 'setReadOnly', { readOnly: true }));
    expect((await client.edit(intent(6, 1, 'setContentControlValue', { handle: checkbox.handle, value: { kind: 'checkbox', checked: true }, expectVersion: after.version }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'read-only' });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker part and property commands preserve each others metadata', async () => {
  const { client, initial } = await openWorker();
  try {
    const header = await client.edit(intent(1, 0, 'ensurePart', {
      kind: 'header', section: 'final', page: 0, expectVersion: initial.frame.payload.documentVersion!,
    }));
    expect(header.outcomes[0]?.status).toBe('applied');
    const changed = await client.edit(intent(2, 1, 'setSectionProperties', {
      section: 'final', patch: { marginLeft: 2880, pageWidth: 10000 }, expectVersion: header.frame!.payload.documentVersion!,
    }));
    expect(changed.outcomes[0]?.status).toBe('applied');
    const footer = await client.edit(intent(3, 2, 'ensurePart', {
      kind: 'footer', section: 'final', page: 0, expectVersion: changed.frame!.payload.documentVersion!,
    }));
    expect(footer.outcomes[0]?.status).toBe('applied');
    const properties = await client.documentRead({ kind: 'properties', target: { kind: 'section', section: 'final' } });
    expect(properties.value).toMatchObject({ properties: { marginLeft: 2880, pageWidth: 10000 } });
    const parts = await client.documentRead({ kind: 'partSnapshot' });
    expect(parts.value.metadata.headers).toHaveLength(1);
    expect(parts.value.metadata.footers).toHaveLength(1);
    const { parseDocx } = await import('../docx/parser');
    const saved = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(saved.package.document.finalSectionProperties).toMatchObject({ marginLeft: 2880, pageWidth: 10000 });
    expect(saved.package.headers!.size).toBe(1);
    expect(saved.package.footers!.size).toBe(1);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker property queries and section mutations preserve versions undo layout and save', async () => {
  const { client, initial, worker } = await openWorker();
  try {
    const version = initial.frame.payload.documentVersion!;
    const read = await client.documentReadAt({ kind: 'properties', target: { kind: 'section', section: 'final' } }, version);
    expect(read.status).toBe('ok');
    if (read.status !== 'ok') throw new Error('Property read superseded');
    expect(read.value).toMatchObject({ kind: 'section', section: 'final' });
    const changed = await client.edit(intent(1, 0, 'setSectionProperties', {
      section: 'final', patch: { marginLeft: 2880, pageWidth: 10000 }, expectVersion: version,
    }));
    expect(changed).toMatchObject({ version: 1, outcomes: [{ status: 'applied' }] });
    expect((await client.documentReadAt({ kind: 'properties', target: { kind: 'section', section: 'final' } }, version)).status).toBe('superseded');
    const retained = applyFrameDeltaOwned(applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame)), decodeFrameDelta(changed.frame!.payload.frame));
    expect(retained.displayList.pages[0]!.width).toBeCloseTo(10000 / 15);
    const current = changed.frame!.payload.documentVersion!;
    const refused = await new Promise<ResidentEngineWorkerResponse>(resolve => {
      const handler = worker.onmessage;
      worker.onmessage = event => {
        if (event.data.id !== 1000000) { handler?.(event); return; }
        worker.onmessage = handler;
        resolve(event.data);
      };
      worker.postMessage({ id: 1000000, type: 'proposal', operation: { kind: 'removeComment', id: 'unissued' } });
    });
    expect(refused).toMatchObject({ ok: false, error: 'Legacy mutations are unavailable in a resident command session' });
    expect((await client.documentRead({ kind: 'properties', target: { kind: 'section', section: 'final' } })).value)
      .toMatchObject({ properties: { marginLeft: 2880, pageWidth: 10000 } });
    const noop = await client.edit(intent(2, 1, 'setSectionProperties', {
      section: 'final', patch: { marginLeft: 2880 }, expectVersion: current,
    }));
    expect(noop).toMatchObject({ version: 1, outcomes: [{ status: 'noop' }] });
    const saved = (await client.save({ comments: [] })).bytes;
    const { parseDocx } = await import('../docx/parser');
    const reopened = await parseDocx(saved);
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 2880, pageWidth: 10000 });
    expect((await client.edit(intent(3, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'properties', target: { kind: 'section', section: 'final' } })).value)
      .not.toHaveProperty('properties.marginLeft', 2880);
    expect((await client.edit(intent(4, 2, 'redo', {}))).outcomes[0]?.status).toBe('applied');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker hyperlink properties expand contiguous runs at both caret boundaries', async () => {
  const { client, point } = await openWorker({ bodyXml: '<w:p w14:paraId="00000001"><w:hyperlink w:anchor="target"><w:r><w:t>First</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Second</w:t></w:r></w:hyperlink><w:r><w:t> plain</w:t></w:r></w:p>' });
  try {
    for (const offset of [0, 3, 5, 11]) {
      const at = { ...point, offset };
      const result = await client.documentRead({ kind: 'properties', target: { kind: 'hyperlink', selection: { anchor: at, focus: at } } });
      expect(result.value).toMatchObject({ kind: 'hyperlink', hyperlink: { href: '#target' }, range: {
        story: 'body', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 11 },
      } });
    }
    const at = { ...point, offset: 14 };
    expect((await client.documentRead({ kind: 'properties', target: { kind: 'hyperlink', selection: { anchor: at, focus: at } } })).value)
      .toMatchObject({ hyperlink: null, range: { story: 'body', start: { paraId: point.paraId, offset: 14 }, end: { paraId: point.paraId, offset: 14 } } });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker setHyperlink inserts and replaces display text', async () => {
  const { client, engine, initial, point } = await openWorker();
  try {
    const at = { ...point, offset: 0 };
    const selection = { anchor: at, focus: at };
    const version = initial.frame.payload.documentVersion!;
    const refused = await client.edit(intent(1, 0, 'setHyperlink', { selection, hyperlink: { href: 'https://example.com/link' }, expectVersion: version }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-range' });
    const inserted = await client.edit(intent(2, 0, 'setHyperlink', { selection, hyperlink: { href: 'https://example.com/link', tooltip: 'Tip' }, displayText: 'Linked', expectVersion: version }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body')[0]).toMatchObject({ text: 'Linked', attributes: { hyperlink: { href: 'https://example.com/link' } } });
    for (const offset of [0, 3, 6]) {
      const caret = { ...at, offset };
      const read = await client.documentRead({ kind: 'properties', target: { kind: 'hyperlink', selection: { anchor: caret, focus: caret } } });
      expect(read.value).toMatchObject({ kind: 'hyperlink', range: { story: 'body', start: { paraId: at.paraId, offset: 0 }, end: { paraId: at.paraId, offset: 6 } }, hyperlink: { href: 'https://example.com/link' } });
    }
    const changed = await client.edit(intent(3, 1, 'setHyperlink', { selection: { anchor: at, focus: { ...at, offset: 6 } }, hyperlink: { href: '#target' }, displayText: 'Changed', expectVersion: inserted.frame!.payload.documentVersion! }));
    expect(changed.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body')[0]).toMatchObject({ text: 'Changed', attributes: { hyperlink: { href: '#target' } } });
    const saved = await client.save({ comments: [] });
    const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(saved.bytes))['word/document.xml']!);
    expect(xml).toContain('w:anchor="target"');
    expect(xml).toContain('Changed');
    expect(xml).not.toContain('Linked');
    const removed = await client.edit(intent(4, 2, 'setHyperlink', { selection: { anchor: at, focus: { ...at, offset: 7 } }, hyperlink: null, expectVersion: saved.version }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body')[0]?.attributes.hyperlink).toBeUndefined();
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker orders hyperlink note and watermark properties and refuses stale requests', async () => {
  const { client, engine, initial, point } = await openWorker();
  try {
    const version = initial.frame.payload.documentVersion!;
    const linked = await client.edit(intent(1, 0, 'setHyperlink', {
      selection: { anchor: { ...point, offset: 0 }, focus: point }, hyperlink: { href: 'https://example.com/link' }, expectVersion: version,
    }));
    expect(linked.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body')[0]).toHaveProperty('attributes.hyperlink.href', 'https://example.com/link');
    const stale = await client.edit(intent(2, 1, 'setNoteProperties', {
      section: 'final', footnotePr: { numStart: 2 }, endnotePr: { numStart: 3 }, expectVersion: version,
    }));
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-version' });
    const notes = await client.edit(intent(3, 1, 'setNoteProperties', {
      section: 'final', footnotePr: { numStart: 2 }, endnotePr: { numStart: 3 }, expectVersion: linked.frame!.payload.documentVersion!,
    }));
    expect(notes.outcomes[0]?.status).toBe('applied');
    const marked = await client.edit(intent(4, 2, 'setWatermark', {
      watermark: { kind: 'text', text: 'Draft', font: 'Liberation Sans', color: '#cccccc', semitransparent: true, layout: 'diagonal' },
      expectVersion: notes.frame!.payload.documentVersion!,
    }));
    expect(marked.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'properties', target: { kind: 'watermark' } })).value)
      .toMatchObject({ watermark: { kind: 'text', text: 'Draft' } });
    const { parseDocx } = await import('../docx/parser');
    const saved = await client.save({ comments: [] });
    const reopened = await parseDocx(saved.bytes);
    expect([...reopened.package.headers!.values()][0]).toHaveProperty('watermark.text', 'Draft');
    const removed = await client.edit(intent(5, 3, 'setWatermark', { watermark: null, expectVersion: saved.version }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'properties', target: { kind: 'watermark' } })).value).toMatchObject({ watermark: null });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker property image handles distinguish shared ids and refuse deleted targets', async () => {
  const { client, engine, initial, point } = await openWorker({ propertyImages: true });
  try {
    const version = initial.frame.payload.documentVersion!;
    const first = (await client.documentReadAt({ kind: 'properties', target: { kind: 'object', at: { ...point, offset: 1 } } }, version));
    const second = (await client.documentReadAt({ kind: 'properties', target: { kind: 'object', at: { ...point, offset: 2 } } }, version));
    if (first.status !== 'ok' || second.status !== 'ok' || first.value?.kind !== 'object' || second.value?.kind !== 'object') throw new Error('Image properties unavailable');
    expect(first.value.handle).not.toBe(second.value.handle);
    const envelope = intent(1, 0, 'setObjectProperties', {
      handle: second.value.handle, patch: { alt: 'Second', widthEmu: 762000 }, expectVersion: version,
    });
    const changed = await client.edit(envelope);
    expect(changed.outcomes[0]?.status).toBe('applied');
    expect((await client.edit(envelope)).version).toBe(1);
    expect(engine.storySegments('body').filter(part => part.kind === 'embed').map(part => part.payload.alt)).toEqual([undefined, 'Second']);
    const deleted = await client.edit(intent(2, 1, 'deleteForward', { selection: { anchor: { ...point, offset: 1 }, focus: { ...point, offset: 1 } } }));
    expect((await client.edit(intent(3, 2, 'setObjectProperties', {
      handle: first.value.handle, patch: { alt: 'Deleted' }, expectVersion: deleted.frame!.payload.documentVersion!,
    }))).outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-object' });
    expect((await client.edit(intent(4, 2, 'setSectionProperties', {
      section: 'final', patch: { marginLeft: -1 }, expectVersion: deleted.frame!.payload.documentVersion!,
    }))).outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-properties' });
    await client.edit(intent(5, 2, 'setReadOnly', { readOnly: true }));
    expect((await client.edit(intent(6, 2, 'setSectionProperties', {
      section: 'final', patch: { marginLeft: 1440 }, expectVersion: deleted.frame!.payload.documentVersion!,
    }))).outcomes[0]).toMatchObject({ status: 'refused', reason: 'read-only' });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker image property reads normalize adjacent image ranges in either direction', async () => {
  const { client, initial, point } = await openWorker({ propertyImages: true });
  try {
    const version = initial.frame.payload.documentVersion!;
    const first = await client.documentReadAt({ kind: 'properties', target: { kind: 'object', at: { ...point, offset: 1 } } }, version);
    if (first.status !== 'ok' || first.value?.kind !== 'object') throw new Error('Image unavailable');
    for (const [anchor, focus] of [[1, 2], [2, 1]]) {
      const selected = await client.documentReadAt<'properties'>({ kind: 'properties', target: { kind: 'objectSelection', selection: {
        anchor: { ...point, offset: anchor! }, focus: { ...point, offset: focus! },
      } } } as never, version);
      expect(selected).toMatchObject({ status: 'ok', value: { kind: 'object', handle: first.value.handle } });
    }
    const multiple = await client.documentReadAt<'properties'>({ kind: 'properties', target: { kind: 'objectSelection', selection: {
      anchor: { ...point, offset: 1 }, focus: { ...point, offset: 3 },
    } } } as never, version);
    expect(multiple).toMatchObject({ status: 'ok', value: null });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker point reads use the native pointer mapping without selecting', async () => {
  const { client, initial, engine } = await openWorker();
  try {
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const page = frame.displayList.pages[0]!;
    const glyph = page.primitives.find(primitive => primitive.kind === 'glyphRun' || primitive.kind === 'text');
    if (!glyph) throw new Error('Glyph unavailable');
    const box = displayPrimitiveRect(glyph);
    const point = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + box.w / 2, y: box.y + box.h / 2,
      frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
      modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
    const version = initial.frame.payload.documentVersion!;
    const before = JSON.stringify(engine.selection());
    const query = await client.documentReadAt<'pointPosition'>({ kind: 'pointPosition', expectVersion: version,
      hit: { position: -1, pageIndex: 0, region: 'body' }, nativePoint: point } as never, version);
    expect(JSON.stringify(engine.selection())).toBe(before);
    expect(engine.proposalEngine.version()).toBe(version);
    const ack = await client.edit(intent(1, 0, 'selectPoint', point));
    const selected = ack.selection!.focus;
    const text = engine.selectionText({ story: selected.story, start: selected, end: selected });
    expect(query).toMatchObject({ status: 'ok', value: { target: { story: selected.story, start: { paraId: selected.paraId, offset: text.before.length } } } });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker object properties preserve dialog units borders and native size', async () => {
  const { client, engine, initial, point } = await openWorker({ propertyImages: true });
  try {
    const read = await client.documentReadAt({ kind: 'properties', target: { kind: 'object', at: { ...point, offset: 1 } } }, initial.frame.payload.documentVersion!);
    if (read.status !== 'ok' || read.value?.kind !== 'object') throw new Error('Image unavailable');
    const payload = { handle: read.value.handle, patch: { width: 80, height: 40, borderWidth: 2, borderColor: '#123456', borderStyle: 'solid' },
      expectVersion: initial.frame.payload.documentVersion! };
    const applied = await client.edit(intent(1, 0, 'setObjectProperties', payload));
    expect(applied.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body').filter(segment => segment.kind === 'embed')[0]).toMatchObject({ payload: {
      width: 80, height: 40, widthEmu: 762000, heightEmu: 381000, borderWidth: 2,
    } });
    const noop = await client.edit(intent(2, 1, 'setObjectProperties', { ...payload, expectVersion: applied.frame!.payload.documentVersion! }));
    expect(noop).toMatchObject({ version: 1, outcomes: [{ status: 'noop' }] });
  } finally { client.destroy(); }
}, TIMEOUT);
test('worker margin and property commands preserve current metadata through undo and save', async () => {
  const { client, initial } = await openWorker();
  try {
    const margins = await client.edit(intent(1, 0, 'setSectionMargins', {
      section: 'final', margins: { marginLeft: 2880, marginRight: 720 }, expectVersion: initial.frame.payload.documentVersion!,
    }));
    expect(margins.outcomes[0]?.status).toBe('applied');
    const properties = await client.edit(intent(2, 1, 'setSectionProperties', {
      section: 'final', patch: { pageWidth: 10000, marginLeft: 3600 }, expectVersion: margins.frame!.payload.documentVersion!,
    }));
    expect(properties.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'sectionProperties' })).value.sections.at(-1)!.properties)
      .toMatchObject({ pageWidth: 10000, marginLeft: 3600, marginRight: 720 });
    const header = await client.edit(intent(3, 2, 'ensurePart', {
      kind: 'header', section: 'final', page: 0, expectVersion: properties.frame!.payload.documentVersion!,
    }));
    expect(header.outcomes[0]?.status).toBe('applied');
    const changed = await client.edit(intent(4, 3, 'setSectionMargins', {
      section: 'final', margins: { marginLeft: 4320 }, expectVersion: header.frame!.payload.documentVersion!,
    }));
    expect(changed.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'properties', target: { kind: 'section', section: 'final' } })).value)
      .toMatchObject({ properties: { pageWidth: 10000, marginLeft: 4320 } });
    const final = await client.edit(intent(5, 4, 'setSectionProperties', {
      section: 'final', patch: { marginLeft: 5040 }, expectVersion: changed.frame!.payload.documentVersion!,
    }));
    expect(final.outcomes[0]?.status).toBe('applied');
    const undone = await client.edit(intent(6, 5, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'sectionProperties' })).value.sections.at(-1)!.properties)
      .toMatchObject({ pageWidth: 10000, marginLeft: 4320, marginRight: 720 });
    const { parseDocx } = await import('../docx/parser');
    const saved = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(saved.package.document.finalSectionProperties).toMatchObject({ pageWidth: 10000, marginLeft: 4320, marginRight: 720 });
    expect(saved.package.headers!.size).toBe(1);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker section snapshots margins and tab stops use native versions and history', async () => {
  const { client, initial, engine, point } = await openWorker();
  try {
    const version = initial.frame.payload.documentVersion!;
    const read = await client.documentReadAt({ kind: 'sectionProperties' }, version);
    expect(read.status).toBe('ok');
    if (read.status !== 'ok') throw new Error('Section snapshot superseded');
    expect(read.value.sections.at(-1)).toMatchObject({ section: 'final' });
    const margins = await client.edit(intent(1, 0, 'setSectionMargins', {
      section: 'final', margins: { marginLeft: 2880, marginRight: 720 }, expectVersion: version,
    }));
    expect(margins.outcomes[0]?.status).toBe('applied');
    const retained = applyFrameDeltaOwned(applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame)), decodeFrameDelta(margins.frame!.payload.frame));
    expect(retained.displayList.pages[0]!.contentBounds!.x).toBeCloseTo(2880 / 15);
    expect((await client.documentReadAt({ kind: 'sectionProperties' }, version)).status).toBe('superseded');
    expect((await client.documentRead({ kind: 'sectionProperties' })).value.sections.at(-1))
      .toHaveProperty('properties.marginLeft', 2880);
    const tabs = await client.edit(intent(2, 1, 'setTabStops', {
      tabs: [{ position: 720, alignment: 'left' }, { position: 1440, alignment: 'right', leader: 'dot' }],
      expectVersion: margins.frame!.payload.documentVersion!,
    }));
    expect(tabs.outcomes[0]?.status).toBe('applied');
    const removed = await client.edit(intent(3, 2, 'removeTabStop', {
      positionTwips: 720, expectVersion: tabs.frame!.payload.documentVersion!,
    }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'sectionProperties' })).value.paragraphs[0]!.tabs)
      .toEqual([{ position: 1440, alignment: 'right', leader: 'dot' }]);
    expect((await client.edit(intent(4, 3, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]!.properties.tabs).toHaveLength(2);
    const { parseDocx } = await import('../docx/parser');
    const saved = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(saved.package.document.finalSectionProperties).toMatchObject({ marginLeft: 2880, marginRight: 720 });
    expect(saved.package.document.content[0]).toHaveProperty('formatting.tabs', [
      { position: 720, alignment: 'left' }, { position: 1440, alignment: 'right', leader: 'dot' },
    ]);
    expect(point.story).toBe('body');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker ruler commands refuse stale invalid and unknown targets without changing versions', async () => {
  const { client, initial } = await openWorker();
  try {
    const expectVersion = initial.frame.payload.documentVersion!;
    expect((await client.edit(intent(1, 0, 'setSectionMargins', { section: 'missing', margins: { marginLeft: 720 }, expectVersion }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'unknown-section' });
    expect((await client.edit(intent(2, 0, 'setSectionMargins', { section: 'final', margins: { marginLeft: -1 }, expectVersion }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'invalid-margins' });
    const changed = await client.edit(intent(3, 0, 'setSectionMargins', { section: 'final', margins: { marginLeft: 1440 }, expectVersion }));
    expect(changed.outcomes[0]?.status).toBe('applied');
    expect((await client.edit(intent(4, 1, 'setTabStops', { tabs: [], expectVersion }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'stale-version' });
    const noop = await client.edit(intent(5, 1, 'setSectionMargins', {
      section: 'final', margins: { marginLeft: 1440 }, expectVersion: changed.frame!.payload.documentVersion!,
    }));
    expect(noop).toMatchObject({ version: 1, outcomes: [{ status: 'noop' }] });
    await client.edit(intent(6, 1, 'setReadOnly', { readOnly: true }));
    expect((await client.edit(intent(7, 1, 'setTabStops', { tabs: [], expectVersion: changed.frame!.payload.documentVersion! }))).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'read-only' });
  } finally { client.destroy(); }
}, TIMEOUT);

test('review_commands_own_comment_crud_and_versioned_snapshots', async () => {
  const { client, engine, point } = await openWorker();
  try {
    const range = { story: 'body', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 4 } };
    const version = engine.proposalEngine.version();
    const added = await client.edit(intent(1, 0, 'addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Before', expectVersion: version }));
    expect(added.outcomes[0]?.status).toBe('applied');
    const duplicate = await client.edit(intent(1, 0, 'addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Before', expectVersion: version }));
    expect(duplicate.outcomes).toEqual(added.outcomes);
    const read = await client.documentRead({ kind: 'reviewSnapshot' });
    const review = read.value!;
    expect(review.comments).toHaveLength(1);
    const id = review.comments[0].id;
    expect(review.comments[0].body).toBe('Before');
    expect(review.commentGeometry[id].length).toBeGreaterThan(0);
    await expect(client.proposal({ kind: 'removeComment', id })).rejects.toThrow('Legacy mutations');
    const updated = await client.edit(intent(2, 1, 'updateComment', { id, patch: { body: 'After', done: true }, expectVersion: read.version }));
    expect(updated.outcomes[0]?.status).toBe('applied');
    expect(await client.documentReadAt({ kind: 'reviewSnapshot' }, read.version)).toEqual({ status: 'superseded' });
    const current = await client.documentRead({ kind: 'reviewSnapshot' });
    expect(current.value!.comments[0]).toMatchObject({ body: 'After', done: true });
    const removed = await client.edit(intent(3, 2, 'removeComment', { id, expectVersion: current.version }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'reviewSnapshot' })).value).toMatchObject({ comments: [] });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker native comment save retains the body and the same range ID across repeated saves', async () => {
  const { client, engine, point } = await openWorker({ bodyXml: '<w:p w14:paraId="00000001"><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Seed</w:t></w:r></w:p>' });
  try {
    const range = { story: 'body', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 4 } };
    const added = await client.edit(intent(1, 0, 'addComment', { ranges: [range], author: 'A', date: '2026-01-01',
      body: 'Saved note', expectVersion: engine.proposalEngine.version() }));
    expect(added.outcomes[0]?.status).toBe('applied');
    const { parseDocx } = await import('../docx/parser');
    const { unzipContainer } = await import('../docx/wasm');
    let previous: number | undefined;
    for (let save = 0; save < 2; save++) {
      const bytes = (await client.save({ comments: [] })).bytes;
      const reopened = await parseDocx(bytes);
      const comments = reopened.package.document.comments!;
      expect(comments).toHaveLength(1);
      const id = comments[0]!.id;
      expect(JSON.stringify(comments[0]!.content)).toContain('Saved note');
      if (previous !== undefined) expect(id).toBe(previous);
      previous = id;
      const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(bytes))['word/document.xml']);
      expect(xml).toContain(`<w:commentRangeStart w:id="${id}"`);
      expect(xml).toContain(`<w:commentRangeEnd w:id="${id}"`);
      expect(xml).toContain('<w:highlight w:val="yellow"');
    }
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker save retains untouched source highlights across repeated saves', async () => {
  const { client } = await openWorker({ bodyXml: '<w:p w14:paraId="00000001"><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Seed</w:t></w:r></w:p>' });
  try {
    const { unzipContainer } = await import('../docx/wasm');
    for (let save = 0; save < 2; save++) {
      const bytes = (await client.save({ comments: [] })).bytes;
      const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(bytes))['word/document.xml']);
      expect(xml).toContain('<w:highlight w:val="yellow"');
    }
  } finally { client.destroy(); }
}, TIMEOUT);

const SOURCE_COMMENT = {
  body: '<w:p w14:paraId="00000001"><w:commentRangeStart w:id="0"/><w:r><w:t>Seed</w:t></w:r><w:commentRangeEnd w:id="0"/>' +
    '<w:r><w:commentReference w:id="0"/></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Other</w:t></w:r></w:p>',
  comments: '<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Parent</w:t></w:r></w:p></w:comment>',
};

async function savedWorkerDocumentXml(client: ResidentEngineWorkerClient): Promise<string> {
  const { unzipContainer } = await import('../docx/wasm');
  return new TextDecoder().decode(unzipContainer(new Uint8Array((await client.save({ comments: [] })).bytes))['word/document.xml']);
}

test('worker reply saves write the same comment range markers on every save', async () => {
  const { client, engine } = await openWorker({ bodyXml: SOURCE_COMMENT.body, commentsXml: SOURCE_COMMENT.comments });
  try {
    const [parent] = engine.listComments();
    if (!parent) throw new Error('Source comment unavailable');
    const other = engine.paragraphs('body')[1]!.paraId;
    const range = { story: 'body', start: { paraId: other, offset: 0 }, end: { paraId: other, offset: 5 } };
    const added = await client.edit(intent(1, 0, 'addComment', { ranges: [range], author: 'B', date: '2026-01-02', body: 'Reply', expectVersion: engine.proposalEngine.version() }));
    expect(added.outcomes[0]?.status).toBe('applied');
    const reply = engine.listComments().find(comment => comment.id !== parent.id)!.id;
    const linked = await client.edit(intent(2, 1, 'updateComment', { id: reply, patch: { parentId: parent.id }, expectVersion: engine.proposalEngine.version() }));
    expect(linked.outcomes[0]?.status).toBe('applied');
    expect(engine.listComments().find(comment => comment.id === reply)?.parentId).toBe(parent.id);
    const markers = async () => (await savedWorkerDocumentXml(client)).match(/<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*\bw:id="\d+"/g);
    const first = await markers();
    expect(first?.length).toBeGreaterThan(0);
    expect(await markers()).toEqual(first);
    expect(await markers()).toEqual(first);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker comment saves keep command page and column breaks in commented paragraphs', async () => {
  const { client, engine, point } = await openWorker({
    bodyXml: '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Other</w:t></w:r></w:p>',
    prepare(engine) {
      engine.applyRawOps('body', [{ op: 'insertEmbed', index: 5, kind: 'columnBreak', payload: {} }]);
      return { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 };
    },
  });
  try {
    const breaks = async () => {
      const xml = await savedWorkerDocumentXml(client);
      return [/<w:br w:type="page"\/>/g, /<w:br w:type="column"\/>/g, /<w:commentRangeStart\b/g].map(pattern => xml.match(pattern)?.length ?? 0);
    };
    const inserted = await client.edit(intent(1, 0, 'insertPageBreak', { selection: { anchor: point, focus: point } }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    const [first, second] = engine.paragraphs('body').map(paragraph => paragraph.paraId);
    const ranges = [{ paraId: first!, end: 5 }, { paraId: second!, end: 6 }]
      .map(({ paraId, end }) => ({ story: 'body', start: { paraId, offset: 1 }, end: { paraId, offset: end } }));
    for (const [index, range] of ranges.entries()) {
      const added = await client.edit(intent(index + 2, index + 1, 'addComment', { ranges: [range], author: 'A', date: '2026-01-01',
        body: `Note ${index}`, expectVersion: engine.proposalEngine.version() }));
      expect(added.outcomes[0]?.status).toBe('applied');
    }
    expect(await breaks()).toEqual([1, 1, 2]);
    const id = engine.listComments()[0]!.id;
    const updated = await client.edit(intent(4, 3, 'updateComment', { id, patch: { body: 'Edited' }, expectVersion: engine.proposalEngine.version() }));
    expect(updated.outcomes[0]?.status).toBe('applied');
    expect(await breaks()).toEqual([1, 1, 2]);
  } finally { client.destroy(); }
}, TIMEOUT);

const RESOLVED_THREAD = {
  comments: '<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="0C000001"><w:r><w:t>Parent</w:t></w:r></w:p></w:comment>' +
    '<w:comment w:id="1" w:author="B" w:date="2026-01-02T00:00:00Z"><w:p w14:paraId="0C000002"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment>',
  extended: '<w15:commentEx w15:paraId="0C000001" w15:done="1"/><w15:commentEx w15:paraId="0C000002" w15:paraIdParent="0C000001" w15:done="1"/>',
  bodies: [
    '<w:p w14:paraId="00000001"><w:commentRangeStart w:id="0"/><w:r><w:t>Seed</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
    '<w:p w14:paraId="00000001"><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>Seed</w:t></w:r><w:commentRangeEnd w:id="0"/>' +
      '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:commentReference w:id="1"/></w:r></w:p>',
  ],
};

test('worker saves an untouched resolved source thread with its reply', async () => {
  const { parseDocx } = await import('../docx/parser');
  for (const bodyXml of RESOLVED_THREAD.bodies) {
    const { client } = await openWorker({ bodyXml, commentsXml: RESOLVED_THREAD.comments, commentsExtendedXml: RESOLVED_THREAD.extended });
    try {
      const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
      const comments = reopened.package.document.comments ?? [];
      const parent = comments.find(comment => JSON.stringify(comment.content).includes('Parent'));
      const reply = comments.find(comment => JSON.stringify(comment.content).includes('Reply'));
      expect(parent).toMatchObject({ id: 0, done: true });
      expect(reply).toMatchObject({ parentId: parent!.id });
    } finally { client.destroy(); }
  }
}, TIMEOUT);

test('worker comment acknowledgements carry the assigned id after an identity-assigning save', async () => {
  const { client, engine, point } = await openWorker();
  try {
    expect((await client.edit(intent(1, 0, 'paragraphBreak', {}))).outcomes[0]?.status).toBe('applied');
    await client.save({ comments: [] });
    const read = engine.proposalEngine.readParagraphs({ view: 'accepted' });
    const seed = read.ok ? read.paragraphs.find(paragraph => paragraph.text === 'Seed')!.paraId : point.paraId;
    const range = { story: 'body', start: { paraId: seed, offset: 0 }, end: { paraId: seed, offset: 4 } };
    const added = await client.edit(intent(2, 1, 'addComment', { ranges: [range], author: 'A', date: '2026-01-01', body: 'Note', expectVersion: engine.proposalEngine.version() }));
    expect(added.outcomes[0]?.status).toBe('applied');
    const comment = added.editState?.comment;
    const id = comment?.id;
    expect(comment?.seq).toBe(2);
    expect(typeof id).toBe('number');
    expect(comment!.comments.map(entry => [entry.id, entry.author])).toEqual([[id!, 'A']]);
    const { parseDocx } = await import('../docx/parser');
    const saved = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(saved.package.document.comments?.map(entry => entry.id)).toContain(id);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker command saves without comments keep the incremental story projection', async () => {
  const { client } = await openWorker({ headersFooters: true });
  try {
    expect((await client.edit(intent(1, 0, 'insertText', { text: 'X' }))).outcomes[0]?.status).toBe('applied');
    const wrap = facadeModule.wrapOpenedEditSession;
    const read: string[] = [];
    const opened = spyOn(facadeModule, 'wrapOpenedEditSession').mockImplementation((...args) => {
      const result = wrap(...args);
      const paragraphIds = result.session.storyParagraphIds;
      result.session.storyParagraphIds = story => { read.push(story); return paragraphIds(story); };
      return result;
    });
    const project = spyOn(projectionModule, 'yrsToDocument');
    try {
      await client.save({ comments: [] });
      const selected = project.mock.calls[0]?.[2]?.storyIds;
      expect(selected).toEqual(new Set(['body']));
      expect(read).toContain('body');
      expect(read.filter(story => !selected?.has(story))).toEqual([]);
    } finally {
      project.mockRestore();
      opened.mockRestore();
    }
  } finally { client.destroy(); }
}, TIMEOUT);

test('review_commands_order_proposals_preview_withdrawal_and_revision_resolution', async () => {
  const { client, engine, point } = await openWorker();
  try {
    const ids = await client.documentRead({ kind: 'paragraphIdentities' });
    const paragraph = ids.value.paragraphs[0]!.session!;
    const proposed = await client.edit(intent(1, 0, 'propose', { expectVersion: ids.version, proposals: [{ id: 'one', op: 'replaceText', paragraph, search: 'Seed', replaceWith: 'Bud', suggest: { author: 'Ada', date: '2026-01-01' } }] }));
    expect(proposed.outcomes[0]?.status).toBe('applied');
    const read = await client.documentRead({ kind: 'reviewSnapshot' });
    const review = read.value!;
    expect(review.proposals.proposals).toHaveLength(1);
    expect(review.proposalHighlights.one.length).toBeGreaterThan(0);
    const decided = await client.edit(intent(2, 1, 'setProposalStates', { expectVersion: read.version, expectPreviewVersion: review.proposals.previewVersion, changes: [{ id: 'one', state: 'accepted' }] }));
    expect(decided.outcomes[0]?.status).toBe('applied');
    expect(engine.proposalEngine.version()).toBe(read.version);
    expect((await client.documentReadAt({ kind: 'reviewSnapshot', expectPreviewVersion: review.previewVersion }, read.version))).toMatchObject({ status: 'ok', value: null });
    expect((await client.documentRead({ kind: 'reviewSnapshot' })).value?.previewVersion).toBe(review.previewVersion + 1);
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame));
    expect(frameText(frame)).toBe('Bud');
    const withdrawn = await client.edit(intent(3, 2, 'withdrawProposals', { expectVersion: read.version, ids: ['one'] }));
    expect(withdrawn.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]?.text).toBe('Bud');
    const inserted = await client.edit(intent(4, 3, 'insertText', { text: '!', selection: { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 0 } }, suggesting: { name: 'Ada', date: '2026-01-01' } }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    const revision = engine.geometryReader.listRevisions()[0]!.revisionId;
    expect((await client.documentRead({ kind: 'reviewSnapshot' })).value!.revisionGeometry[revision].length).toBeGreaterThan(0);
    const rejected = await client.edit(intent(5, 4, 'rejectRevision', { revisionId: revision, expectVersion: engine.proposalEngine.version() }));
    expect(rejected.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body')[0]?.text).toBe('Bud');
  } finally { client.destroy(); }
}, TIMEOUT);

test('identity assigning save presents its native version before version checked reads', async () => {
  const { client, engine, point } = await openWorker();
  const frames: DocxCommandFrame[] = [];
  const unsubscribe = client.onFrame((frame) => { frames.push(frame); });
  try {
    const broken = await client.edit(intent(1, 0, 'paragraphBreak', {}));
    expect(broken.outcomes[0]?.status).toBe('applied');
    const saved = await client.save({ comments: [] });
    expect(saved.version).not.toBe(broken.frame!.payload.documentVersion);
    expect(saved.version).toBe(engine.proposalEngine.version());
    expect(frames.map((frame) => frame.payload.documentVersion)).toEqual([saved.version]);
    const frameEpoch = frames[0]!.revision;
    expect(await client.documentReadAt({ kind: 'reviewSnapshot' }, saved.version)).toMatchObject({
      status: 'ok', version: saved.version, value: { version: saved.version, frameEpoch },
    });
    const range = { story: 'body', start: { paraId: point.paraId, offset: 0 }, end: { paraId: point.paraId, offset: 4 } };
    expect(await client.documentReadAt({ kind: 'rangeHighlights', ranges: [range], expectVersion: saved.version }, saved.version))
      .toMatchObject({ status: 'ok', value: { version: saved.version, frameEpoch } });
    const commented = await client.edit(intent(2, 1, 'addComment', {
      ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Saved', expectVersion: saved.version,
    }));
    expect(commented.outcomes[0]?.status).toBe('applied');
  } finally {
    unsubscribe();
    client.destroy();
  }
}, TIMEOUT);

test('save that assigns no identities keeps its version frame and layout', async () => {
  const { client, engine } = await openWorker();
  const frames: DocxCommandFrame[] = [];
  const unsubscribe = client.onFrame((frame) => { frames.push(frame); });
  const layouts = spyOn(engine, 'layoutDocumentWithRegionsRetained');
  try {
    expect((await client.edit(intent(1, 0, 'paragraphBreak', {}))).outcomes[0]?.status).toBe('applied');
    const first = await client.save({ comments: [] });
    const published = frames.length;
    const laidOut = layouts.mock.calls.length;
    const parts = await client.documentReadAt({ kind: 'partSnapshot' }, first.version);
    const second = await client.save({ comments: [] });
    expect(second.version).toBe(first.version);
    expect(frames).toHaveLength(published);
    expect(layouts).toHaveBeenCalledTimes(laidOut);
    expect(await client.documentReadAt({ kind: 'partSnapshot' }, second.version)).toEqual(parts);
  } finally {
    layouts.mockRestore();
    unsubscribe();
    client.destroy();
  }
}, TIMEOUT);

test('toolbar_metadata_and_insertion_commands_are_versioned_worker_services', async () => {
  const { client, engine, point } = await openWorker();
  try {
    const metadata = await client.documentRead({ kind: 'toolbarMetadata' });
    expect(metadata.value).toMatchObject({ styles: [], fontFamilies: expect.any(Array) });
    const selected = { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 4 } };
    const format = await client.edit(intent(1, 0, 'formatSelection', { selection: selected, delta: { bold: true, fontSize: 14 } }));
    expect(format.outcomes[0]?.status).toBe('applied');
    expect(format.editState?.toolbar).toMatchObject({ bold: true, fontSize: 28 });
    const list = await client.edit(intent(2, 1, 'toggleList', { kind: 'numbered' }));
    expect(list.outcomes[0]?.status).toBe('applied');
    expect(list.editState?.toolbar?.paragraphProperties.numPr).toMatchObject({ numId: 2, ilvl: 0 });
    const listedText = frameText(applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame)));
    const table = await client.edit(intent(3, 2, 'insertTable', { selection: { anchor: point, focus: point }, rows: 2, columns: 2 }));
    expect(table.outcomes[0]?.status).toBe('applied');
    expect(table.editState?.tableContext).toMatchObject({ rowCount: 2, columnCount: 2 });
    expect(table.selection?.focus.story).toMatch(/:t\d+:r0c0~\d+$/);
    expect(engine.geometryReader.storyIds().length).toBe(5);
    expect(await client.documentReadAt({ kind: 'toolbarMetadata' }, metadata.version)).toEqual({ status: 'superseded' });
    const undone = await client.edit(intent(4, 3, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(undone).toMatchObject({ version: 4, selection: { focus: { story: 'body' } }, editState: { canRedo: true } });
    expect((await client.edit(intent(4, 3, 'undo', {}))).outcomes).toEqual(undone.outcomes);
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed']);
    expect(engine.geometryReader.storyIds()).toEqual(['body']);
    expect(frameText(applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame)))).toBe(listedText);
    const staleCell = await client.edit(intent(5, 4, 'insertText', { text: '?', selection: table.selection! }));
    expect(staleCell).toMatchObject({ version: 4, outcomes: [{ status: 'refused', reason: 'stale-anchor' }] });
    const redone = await client.edit(intent(6, 4, 'redo', {}));
    expect(redone.outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed', '']);
    expect(engine.geometryReader.storyIds().length).toBe(5);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker frames publish selection and distinct stable image handles', async () => {
  const { client, initial, point } = await openWorker({ objects: true });
  try {
    const geometry = initial.frame.payload.geometry!;
    expect(geometry).toMatchObject({ version: initial.frame.payload.documentVersion, frameEpoch: initial.frame.revision });
    const images = geometry.objects.filter(object => object.kind === 'image');
    expect(images).toHaveLength(2);
    expect(new Set(images.map(object => object.handle)).size).toBe(2);
    expect(images.every(object => object.rects.length > 0)).toBe(true);
    const selected = await editWithFrame(client, intent(1, 0, 'select', { selection: {
      anchor: { ...point, offset: 0 }, focus: { ...point, offset: 1 },
    } }));
    expect(selected.frame!.payload.geometry!.selection.length).toBeGreaterThan(0);
    const typed = await client.edit(intent(2, 0, 'insertText', {
      text: '+', selection: { anchor: { ...point, offset: 0 }, focus: { ...point, offset: 0 } },
    }));
    expect(typed.frame!.payload.geometry!.objects.filter(object => object.kind === 'image').map(object => object.handle))
      .toEqual(images.map(object => object.handle));
  } finally { client.destroy(); }
}, TIMEOUT);

test('the public session has no embed move, so collaborating replicas cannot duplicate a picture', async () => {
  const session = await createYrsSession({ clientId: 8321 });
  try { expect('moveEmbed' in session).toBe(false); } finally { session.destroy(); }
});

for (const destination of ['forward', 'backward', 'self'] as const) {
  test(`worker moveEmbed same paragraph ${destination} preserves the picture`, async () => {
    const { client, engine, initial } = await openWorker({ prepare(engine) {
      engine.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'image', payload: { id: 'inline', src: 'data:image/png;base64,', width: 40, height: 20 } }]);
      return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
    } });
    try {
      const image = initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!;
      const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
      const glyphs = frame.displayList.pages[0]!.primitives.filter(p => p.kind === 'glyphRun' || p.kind === 'text');
      const box = destination === 'self' ? { x: image.rects[0]!.x, y: image.rects[0]!.y, w: image.rects[0]!.width, h: image.rects[0]!.height }
        : displayPrimitiveRect(glyphs[destination === 'backward' ? 0 : glyphs.length - 1]!);
      const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + (destination === 'forward' ? box.w : 0), y: box.y + box.h / 2,
        frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
      const moved = await client.edit(intent(1, 0, 'moveEmbed', { handle: image.handle, expectVersion: initial.frame.payload.documentVersion!, drop } as never));
      expect(moved.outcomes[0]?.status).toBe(destination === 'self' ? 'noop' : 'applied');
      expect(moved.version).toBe(destination === 'self' ? 0 : 1);
      if (destination === 'self') expect(engine.canUndo()).toBe(false);
      else {
        expect(moved.selection!.anchor.offset).toBe(destination === 'backward' ? 0 : 4);
        expect(moved.selection!.focus.offset).toBe(moved.selection!.anchor.offset + 1);
      }
    } finally { client.destroy(); }
  }, TIMEOUT);
}

for (const refusal of ['unsupported-suggesting', 'read-only', 'stale-version', 'unsupported-object', 'stale-object', 'stale-epoch', 'locked-target'] as const) {
  test(`worker moveEmbed refuses ${refusal === 'locked-target' ? 'tracked deletion with locked-target' : refusal} without mutation`, async () => {
    const { client, engine, initial } = await openWorker({ prepare(engine) {
      engine.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'image', payload: {
        id: 'inline', src: 'data:image/png;base64,', width: 40, height: 20,
        ...(refusal === 'unsupported-object' ? { displayMode: 'float', wrapType: 'square', position: { horizontal: { relativeTo: 'page', posOffset: 0 }, vertical: { relativeTo: 'page', posOffset: 0 } } } : {}),
      } }]);
      return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
    } });
    try {
      const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
      const image = initial.frame.payload.geometry!.objects.find(o => o.kind === 'image');
      const glyph = frame.displayList.pages[0]!.primitives.find(p => p.kind === 'glyphRun' || p.kind === 'text')!;
      const box = displayPrimitiveRect(glyph);
      const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x, y: box.y + box.h / 2,
        frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
      if (refusal === 'read-only') await client.edit(intent(1, 0, 'setReadOnly', { readOnly: true }));
      if (refusal === 'locked-target') engine.applyRawOps('body', [{ op: 'format', index: 1, len: 1, attrs: { del: { id: 'pending-removal', author: 'Ada', date: '2026-01-01' } } }]);
      const before = engine.encodeStateVector();
      const request = intent(2, 0, 'moveEmbed', { handle: refusal === 'stale-object' ? 'missing' : image?.handle ?? 'missing',
        expectVersion: refusal === 'stale-version' ? 'stale' : engine.proposalEngine.version(), drop,
        ...(refusal === 'unsupported-suggesting' ? { suggesting: { name: 'Ada', date: '2026-01-01' } } : {}) } as never);
      if (refusal === 'stale-epoch') request.basis.epoch = 0;
      const result = await client.edit(request);
      expect(result.outcomes[0]).toMatchObject({ status: 'refused', reason: refusal });
      expect(result.version).toBe(0);
      expect(engine.encodeStateVector()).toEqual(before);
    } finally { client.destroy(); }
  }, TIMEOUT);
}

for (const stale of ['layoutRevision', 'frameEpoch', 'pageId', 'pageIndex', 'collaboration'] as const) {
  test(`worker moveEmbed ${stale === 'frameEpoch' ? 'accepts a drop from an earlier frame of the same layout' : `refuses stale drop ${stale}`}`, async () => {
    const { client, engine, initial } = await openWorker({ collaboration: stale === 'collaboration', objects: true });
    try {
      const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
      const glyph = frame.displayList.pages[0]!.primitives.find(p => p.kind === 'glyphRun' || p.kind === 'text')!;
      const box = displayPrimitiveRect(glyph);
      const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + box.w, y: box.y + box.h / 2,
        frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
      if (stale === 'layoutRevision') drop.layoutRevision++;
      if (stale === 'frameEpoch') drop.frameEpoch++;
      if (stale === 'pageId') drop.page.pageId = 'obsolete';
      if (stale === 'pageIndex') drop.page.pageIndex++;
      const before = engine.encodeStateVector();
      const result = await client.edit(intent(1, 0, 'moveEmbed', { handle: initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!.handle,
        expectVersion: initial.frame.payload.documentVersion!, drop } as never));
      if (stale === 'frameEpoch') {
        // A selection-only frame does not move text, so a drop captured on the previous frame still applies.
        expect(result.outcomes[0]!.status).not.toBe('refused');
        return;
      }
      expect(result.outcomes[0]).toMatchObject({ status: 'refused', reason: stale === 'collaboration' ? 'unsupported-command' : 'stale-anchor' });
      expect(result.version).toBe(0);
      expect(engine.encodeStateVector()).toEqual(before);
    } finally { client.destroy(); }
  }, TIMEOUT);
}

for (const tracked of ['source', 'destination'] as const) {
  test(`worker moveEmbed refuses a tracked row at the ${tracked}`, async () => {
    const { client, engine, initial } = await openWorker({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="7" w:author="Ada" w:date="2026-01-01T00:00:00Z"/></w:trPr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', prepare(engine) {
      const story = tracked === 'source' ? 'body:t0:r0c0' : 'body';
      engine.applyRawOps(story, [{ op: 'insertEmbed', index: 1, kind: 'image', payload: { id: 'inline', src: 'data:image/png;base64,', width: 40, height: 20 } }]);
      return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
    } });
    try {
      const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
      const story = tracked === 'source' ? 'body' : 'body:t0:r0c0';
      const rect = initial.frame.payload.geometry!.objects.find(o => o.kind === 'text' && o.story === story)!.rects[0]!;
      const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: rect.x + rect.width / 2, y: rect.y + rect.height / 2,
        frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
      const before = engine.encodeStateVector();
      const result = await client.edit(intent(1, 0, 'moveEmbed', { handle: initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!.handle,
        expectVersion: initial.frame.payload.documentVersion!, drop } as never));
      expect(result.outcomes[0]).toMatchObject({ status: 'refused', reason: 'tracked-revision-conflict' });
      expect(result.version).toBe(0);
      expect(engine.encodeStateVector()).toEqual(before);
    } finally { client.destroy(); }
  }, TIMEOUT);
}

for (const tracked of [false, true]) {
  test(`worker moveEmbed moves across paragraphs in one undo step with a tracked insertion: ${tracked}`, async () => {
    const { client, engine, initial } = await openWorker({ bodyXml: '<w:p><w:r><w:t>AB</w:t></w:r></w:p><w:p><w:r><w:t>CD</w:t></w:r></w:p>', prepare(engine) {
      engine.applyRawOps('body', [{ op: 'insertEmbed', index: tracked ? 0 : 1, kind: 'image', payload: { id: 'inline', src: 'data:image/png;base64,', width: 40, height: 20, nested: { locks: [true, false] } }, attrs: { bold: true, ...(tracked ? { ins: { id: 'pending-image', author: 'Ada', date: '2026-01-01' } } : {}) } }]);
      return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
    } });
    try {
      const before = engine.storySegments('body');
      const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
      const glyphs = frame.displayList.pages[0]!.primitives.filter(p => p.kind === 'glyphRun' || p.kind === 'text');
      const box = displayPrimitiveRect(glyphs[glyphs.length - 1]!);
      const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + box.w, y: box.y + box.h / 2,
        frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
      const moved = await client.edit(intent(1, 0, 'moveEmbed', { handle: initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!.handle, expectVersion: initial.frame.payload.documentVersion!, drop } as never));
      expect(moved.outcomes[0]).toMatchObject({ status: 'applied' });
      expect(moved.selection!.focus.offset - moved.selection!.anchor.offset).toBe(1);
      expect(moved.selection!.anchor.paraId).toBe(engine.geometryReader.paragraphSpans('body')[1]!.paraId);
      const after = engine.storySegments('body');
      expect(after.filter(s => s.kind === 'embed')).toEqual(before.filter(s => s.kind === 'embed'));
      const undone = await client.edit(intent(2, 1, 'undo', {}));
      expect(undone.outcomes[0]?.status).toBe('applied');
      expect(engine.storySegments('body')).toEqual(before);
      const redone = await client.edit(intent(3, 2, 'redo', {}));
      expect(redone.outcomes[0]?.status).toBe('applied');
      expect(engine.storySegments('body')).toEqual(after);
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('worker moveEmbed moves from body text into a table cell in the body part', async () => {
  const { client, engine, initial } = await openWorker({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', prepare(engine) {
    engine.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'image', payload: { id: 'inline', src: 'data:image/png;base64,', width: 40, height: 20 } }]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
  } });
  try {
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const rect = initial.frame.payload.geometry!.objects.find(o => o.kind === 'text' && o.story === 'body:t0:r0c0')?.rects[0];
    if (!rect) throw new Error('Cell text rectangle unavailable');
    const box = { x: rect.x, y: rect.y, w: rect.width, h: rect.height };
    const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + box.w / 2, y: box.y + box.h / 2,
      frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
      modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
    const moved = await client.edit(intent(1, 0, 'moveEmbed', { handle: initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!.handle, expectVersion: initial.frame.payload.documentVersion!, drop } as never));
    expect(moved.outcomes[0]?.status).toBe('applied');
    expect(moved.selection!.anchor.story).toBe('body:t0:r0c0');
    expect(engine.storySegments('body').filter(s => s.kind === 'embed' && s.embedKind === 'image')).toHaveLength(0);
    expect(engine.storySegments('body:t0:r0c0').filter(s => s.kind === 'embed' && s.embedKind === 'image')).toHaveLength(1);
    await client.edit(intent(2, 1, 'undo', {}));
    expect(engine.storySegments('body').filter(s => s.kind === 'embed' && s.embedKind === 'image')).toHaveLength(1);
    expect(engine.storySegments('body:t0:r0c0').filter(s => s.kind === 'embed' && s.embedKind === 'image')).toHaveLength(0);
  } finally { client.destroy(); }
}, TIMEOUT);

for (const locked of ['source', 'destination'] as const) {
  test(`worker moveEmbed refuses a locked content control at the ${locked}`, async () => {
    const { client, engine, initial } = await openWorker({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="7"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked</w:t></w:r></w:p></w:sdtContent></w:sdt>', prepare(engine) {
      const story = locked === 'source' ? engine.geometryReader.storyIds().find(id => id.startsWith('body:'))! : 'body';
      engine.applyRawOps(story, [{ op: 'insertEmbed', index: 1, kind: 'image', payload: { id: 'inline', src: 'data:image/png;base64,', width: 40, height: 20 } }]);
      return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
    } });
    try {
      const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
      const story = locked === 'source' ? 'body' : engine.geometryReader.storyIds().find(id => id.startsWith('body:'));
      const rect = initial.frame.payload.geometry!.objects.find(o => o.kind === 'text' && o.story === story)?.rects[0];
      if (!rect) throw new Error('Control text rectangle unavailable');
      const box = { x: rect.x, y: rect.y, w: rect.width, h: rect.height };
      const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + box.w / 2, y: box.y + box.h / 2,
        frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
      const before = engine.encodeStateVector();
      const moved = await client.edit(intent(1, 0, 'moveEmbed', { handle: initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!.handle, expectVersion: initial.frame.payload.documentVersion!, drop } as never));
      expect(moved.outcomes[0]).toMatchObject({ status: 'refused', reason: 'locked-target' });
      expect(engine.encodeStateVector()).toEqual(before);
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('worker moveEmbed save and reopen retain one drawing and its relationship', async () => {
  const drawing = '<w:r><w:rPr><w:b/></w:rPr><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="381000" cy="190500"/><wp:docPr id="9" name="Picture"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const bodyXml = `<w:p><w:r><w:t>A</w:t></w:r>${drawing}</w:p><w:p><w:r><w:t>Destination</w:t></w:r></w:p>`;
  const { client, engine, initial, init } = await openWorker({ imageMedia: true, bodyXml });
  const baseline = await openWorker({ imageMedia: true, bodyXml });
  const reopened = new ResidentEngineWorkerClient(startWorker(8302));
  try {
    const original = engine.storySegments('body').filter(s => s.kind === 'embed');
    expect(original.every(s => s.kind === 'embed' && s.payload.docPr === undefined)).toBe(true);
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const glyphs = frame.displayList.pages[0]!.primitives.filter(p => p.kind === 'glyphRun' || p.kind === 'text');
    const box = displayPrimitiveRect(glyphs[glyphs.length - 1]!);
    const drop = { page: { pageId: String(frame.pages[0]!.pageId), pageIndex: 0 }, x: box.x + box.w, y: box.y + box.h / 2,
      frameEpoch: frame.frameEpoch, layoutRevision: initial.frame.payload.layoutRevision, clickCount: 1,
      modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
    const moved = await client.edit(intent(1, 0, 'moveEmbed', { handle: initial.frame.payload.geometry!.objects.find(o => o.kind === 'image')!.handle, expectVersion: initial.frame.payload.documentVersion!, drop } as never));
    expect(moved.outcomes[0]?.status).toBe('applied');
    expect(engine.storySegments('body').filter(s => s.kind === 'embed')).toEqual(original);
    const saved = await client.save({ comments: [] });
    const parts = unzipContainer(new Uint8Array(saved.bytes));
    const xml = new TextDecoder().decode(parts['word/document.xml']!);
    expect(xml.match(/<w:drawing[ >]/g)).toHaveLength(1);
    expect(xml.match(/r:embed="rIdImage"/g)).toHaveLength(1);
    const ids = [...xml.matchAll(/<wp:docPr\b[^>]*\bid="([^"]*)"/g)].map(match => match[1]);
    expect(ids).toHaveLength(1);
    expect(new Set(ids).size).toBe(ids.length);
    await baseline.client.edit(intent(1, 0, 'insertText', { text: 'X' }));
    const edited = unzipContainer(new Uint8Array((await baseline.client.save({ comments: [] })).bytes));
    const editedXml = new TextDecoder().decode(edited['word/document.xml']!);
    expect(xml.match(/<w:drawing[ >][\s\S]*?<\/w:drawing>/g)).toEqual(editedXml.match(/<w:drawing[ >][\s\S]*?<\/w:drawing>/g));
    expect(xml).toContain('noChangeAspect="1"');
    const rels = new TextDecoder().decode(parts['word/_rels/document.xml.rels']!);
    expect(rels.match(/Id="rIdImage"/g)).toHaveLength(1);
    expect(Object.keys(parts).filter(key => key === 'word/media/picture.png')).toHaveLength(1);
    await reopened.open(new Uint8Array(saved.bytes), { generation: 'd1-command-worker' });
    const restored = await reopened.commandInit({ ...init, selection: null });
    expect(restored.frame.payload.geometry!.objects.filter(o => o.kind === 'image')).toHaveLength(1);
  } finally { baseline.client.destroy(); reopened.destroy(); client.destroy(); }
}, TIMEOUT);

test('worker orders versioned image resize and move with undo and stale-handle refusal', async () => {
  const { client, engine, initial, point } = await openWorker({ objects: true });
  try {
    const handle = initial.frame.payload.geometry!.objects.find(object => object.kind === 'image')!.handle;
    const expectVersion = initial.frame.payload.documentVersion!;
    const request = intent(1, 0, 'resizeObject', { handle, widthEmu: 762000, heightEmu: 381000, expectVersion });
    const resized = await client.edit(request);
    expect(resized).toMatchObject({ version: 1, outcomes: [{ status: 'applied' }] });
    expect((await client.edit(request)).version).toBe(1);
    const imageRects = (response: typeof resized) => response.frame!.payload.geometry!.objects
      .filter(object => object.kind === 'image').map(object => object.rects[0]!);
    expect(imageRects(resized).map(rect => [rect.width, rect.height])).toEqual([[80, 40], [40, 20]]);
    const stale = await client.edit(intent(2, 0, 'moveObject', { handle, hOffsetEmu: 95250, vOffsetEmu: 190500, expectVersion }));
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-version' });
    const moved = await client.edit(intent(3, 1, 'moveObject', {
      handle, hOffsetEmu: 95250, vOffsetEmu: 190500, expectVersion: resized.frame!.payload.documentVersion!,
    }));
    expect(moved.outcomes[0]?.status).toBe('applied');
    expect(imageRects(moved)[0]).toMatchObject({ x: 10, y: 20, width: 80, height: 40 });
    const undone = await client.edit(intent(4, 2, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(imageRects(undone)).toEqual(imageRects(resized));
    const deleted = await client.edit(intent(5, 3, 'deleteForward', { selection: {
      anchor: { ...point, offset: 1 }, focus: { ...point, offset: 1 },
    } }));
    const refused = await client.edit(intent(6, 4, 'resizeObject', {
      handle, widthEmu: 381000, heightEmu: 190500, expectVersion: deleted.frame!.payload.documentVersion!,
    }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-object' });
    expect(engine.storySegments('body').filter(segment => segment.kind === 'embed')).toHaveLength(1);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker moves a floating image without changing its wrap and flags which images float', async () => {
  const image = { src: 'data:image/png;base64,', width: 40, height: 20 };
  const { client, engine, initial } = await openWorker({ prepare: engine => {
    engine.applyRawOps('body', [
      { op: 'insertEmbed', index: 1, kind: 'image', payload: { ...image, id: 'floating-image', wrapType: 'behind', displayMode: 'float', cssFloat: 'none',
        position: { behindDoc: true, horizontal: { relativeTo: 'margin', posOffset: 0 }, vertical: { relativeTo: 'margin', posOffset: 0 } } } },
      { op: 'insertEmbed', index: 2, kind: 'image', payload: { ...image, id: 'inline-image' } },
    ]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
  } });
  try {
    const images = (response: { frame?: typeof initial.frame }) => response.frame!.payload.geometry!.objects.filter(object => object.kind === 'image');
    expect(images(initial).map(object => object.floating)).toEqual([true, false]);
    expect(images(initial).map(object => object.behindText)).toEqual([true, undefined]);
    const payload = (id: string) => engine.storySegments('body').find(segment => segment.kind === 'embed' && segment.payload.id === id)!;
    const original = structuredClone(payload('floating-image'));
    const moved = await client.edit(intent(1, 0, 'moveObject', {
      handle: images(initial)[0]!.handle, hOffsetEmu: 95250, vOffsetEmu: 190500, expectVersion: initial.frame.payload.documentVersion!,
    }));
    expect(moved.outcomes[0]?.status).toBe('applied');
    expect(payload('floating-image')).toMatchObject({ payload: { wrapType: 'behind', displayMode: 'float', position: {
      behindDoc: true, horizontal: { relativeTo: 'page', posOffset: 95250 }, vertical: { relativeTo: 'page', posOffset: 190500 },
    } } });
    const undone = await client.edit(intent(2, 1, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(payload('floating-image')).toEqual(original);
    const floated = await client.edit(intent(3, 2, 'moveObject', {
      handle: images(undone)[1]!.handle, hOffsetEmu: 0, vOffsetEmu: 0, expectVersion: undone.frame!.payload.documentVersion!,
    }));
    expect(floated.outcomes[0]?.status).toBe('applied');
    expect(payload('inline-image')).toMatchObject({ payload: { wrapType: 'square', displayMode: 'float' } });
    expect(images(floated).map(object => object.floating)).toEqual([true, true]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker preserves picture aspect locks through resize, move, save and reopen', async () => {
  const picture = (container: 'inline' | 'anchor', id: number, lock: 'frame' | 'picture' | undefined) =>
    `<w:p><w:r><w:drawing><wp:${container} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    (container === 'anchor' ? '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' : '') +
    '<wp:extent cx="381000" cy="190500"/>' + (container === 'anchor' ? '<wp:wrapNone/>' : '') +
    `<wp:docPr id="${id}"/>` + (lock === 'frame' ? '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' : '') +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPicPr>' +
    (lock === 'picture' ? '<a:picLocks noChangeAspect="1"/>' : '') +
    '</pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic>' +
    `</wp:${container}></w:drawing></w:r></w:p>`;
  const { client, engine, initial, init } = await openWorker({ bodyXml:
    picture('inline', 1, 'frame') + picture('anchor', 2, 'picture') + picture('inline', 3, undefined) + picture('anchor', 4, undefined),
  });
  const reopened = new ResidentEngineWorkerClient(startWorker(8302));
  try {
    const images = (response: { frame?: typeof initial.frame }) => response.frame!.payload.geometry!.objects.filter(object => object.kind === 'image');
    const expectLocks = (response: { frame?: typeof initial.frame }) => {
      expect(images(response)).toHaveLength(4);
      expect(images(response).map(object => object.noChangeAspect)).toEqual([true, true, undefined, undefined]);
      for (const object of images(response).slice(2)) expect(object).not.toHaveProperty('noChangeAspect');
    };
    expectLocks(initial);
    const { readResidentFrameGeometry: referenceGeometry } = await import('./residentObjectGeometry.reference');
    const reference = await referenceGeometry(engine,
      applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame)), init.selection ?? null, init.layoutRevision);
    expect(reference.objects.filter(object => object.kind === 'image').map(object => object.noChangeAspect))
      .toEqual(images(initial).map(object => object.noChangeAspect));
    expect(engine.storySegments('body').flatMap(segment => segment.kind === 'embed' && segment.embedKind === 'image'
      ? [segment.payload.noChangeAspect] : [])).toEqual([true, true, undefined, undefined]);
    expect((engine.yrsBlocksForStory('body') as LayoutBlock[])
      .flatMap(block => block.kind === 'paragraph' ? block.runs : [])
      .flatMap(run => run.kind === 'image' ? [run.noChangeAspect] : [])).toEqual([true, true, undefined, undefined]);
    const resized = await client.edit(intent(1, 0, 'resizeObject', {
      handle: images(initial)[0]!.handle, widthEmu: 762000, heightEmu: 381000, expectVersion: initial.frame.payload.documentVersion!,
    }));
    expect(resized.outcomes[0]?.status).toBe('applied');
    expectLocks(resized);
    const moved = await client.edit(intent(2, 1, 'moveObject', {
      handle: images(resized)[1]!.handle, hOffsetEmu: 95250, vOffsetEmu: 190500, expectVersion: resized.frame!.payload.documentVersion!,
    }));
    expect(moved.outcomes[0]?.status).toBe('applied');
    expectLocks(moved);
    // Edits rewrite the unlocked pictures' paragraphs too; saving them must not invent a lock.
    const resizedUnlocked = await client.edit(intent(3, 2, 'resizeObject', {
      handle: images(moved)[2]!.handle, widthEmu: 762000, heightEmu: 95250, expectVersion: moved.frame!.payload.documentVersion!,
    }));
    expect(resizedUnlocked.outcomes[0]?.status).toBe('applied');
    const movedUnlocked = await client.edit(intent(4, 3, 'moveObject', {
      handle: images(resizedUnlocked)[3]!.handle, hOffsetEmu: 190500, vOffsetEmu: 285750, expectVersion: resizedUnlocked.frame!.payload.documentVersion!,
    }));
    expect(movedUnlocked.outcomes[0]?.status).toBe('applied');
    expectLocks(movedUnlocked);
    const saved = await client.save({ comments: [] });
    const { parseDocx } = await import('../docx/parser');
    const { documentToYrs } = await import('./documentToYrs');
    const parsed = await parseDocx(saved.bytes, { preloadFonts: false });
    const projected = await createYrsSession();
    try {
      documentToYrs(projected, parsed);
      expect(projected.storySegments('body').flatMap(segment => segment.kind === 'embed' && segment.embedKind === 'image'
        ? [segment.payload.noChangeAspect] : [])).toEqual([true, true, undefined, undefined]);
      expect(projectionModule.yrsToDocument(projected, parsed).package.document.content
        .flatMap(block => block.type === 'paragraph' ? block.content : [])
        .flatMap(content => content.type === 'run' ? content.content : [])
        .flatMap(content => content.type === 'drawing' ? [content.image.noChangeAspect] : [])).toEqual([true, true, undefined, undefined]);
    } finally { projected.destroy(); }
    await reopened.open(new Uint8Array(saved.bytes), { generation: 'd1-command-worker' });
    const restored = await reopened.commandInit(init);
    expectLocks(restored);
    expect(images(restored)[0]!.rects[0]).toMatchObject({ width: 80, height: 40 });
  } finally { reopened.destroy(); client.destroy(); }
}, TIMEOUT);

test('worker selects a picture by its handle, refuses a stale version, and flags the selected embed', async () => {
  const drawing = (id: number) => `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<wp:extent cx="381000" cy="190500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  // The second picture is alone in its paragraph, whose selection rects also carry an empty-line marker.
  const { client, engine, initial } = await openWorker({ bodyXml: `<w:p><w:r><w:t>ab</w:t></w:r>${drawing(1)}</w:p><w:p>${drawing(2)}</w:p>` });
  try {
    const images = (response: { frame?: typeof initial.frame }) => response.frame!.payload.geometry!.objects.filter(object => object.kind === 'image');
    expect(images(initial).map(object => object.selected)).toEqual([undefined, undefined]);
    const [first, second] = images(initial).map(object => object.handle);
    const staleVersion = initial.frame.payload.documentVersion!;
    const paraId = engine.geometryReader.paragraphSpans('body')[0]!.paraId;
    const start = { story: 'body', paraId, offset: 0 };
    await editWithFrame(client, intent(1, 0, 'select', { selection: { anchor: start, focus: start } }));
    const typed = await client.edit(intent(2, 0, 'insertText', { text: 'x' }));
    expect(typed.outcomes[0]?.status).toBe('applied');
    // A press on the picture as drawn before the edit: its version is stale, so nothing is selected and no text can be.
    const stale = await client.edit(intent(3, 1, 'select', { object: { handle: first!, expectVersion: staleVersion } }));
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-version' });
    // The handle follows the picture to offset 3 after the insert before it.
    const selected = await editWithFrame(client, intent(4, 1, 'select', { object: { handle: first!, expectVersion: typed.frame!.payload.documentVersion! } }));
    expect(images(selected).map(object => object.selected)).toEqual([true, undefined]);
    const alone = await editWithFrame(client, intent(5, 1, 'select', { object: { handle: second!, expectVersion: typed.frame!.payload.documentVersion! } }));
    expect(images(alone).map(object => object.selected)).toEqual([undefined, true]);
    const removed = await client.edit(intent(6, 1, 'deleteForward', {}));
    expect(removed.outcomes[0]?.status).toBe('applied');
    // Delete removes the selected picture only; the text before the first picture is intact.
    expect(engine.storySegments('body').filter(segment => segment.kind === 'embed')).toHaveLength(1);
    expect(engine.storySegments('body').filter(segment => segment.kind === 'text').map(segment => segment.text).join('')).toBe('xab');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker moves a Top and Bottom picture by floating it, since layout keeps that wrap in the text flow', async () => {
  const image = { src: 'data:image/png;base64,', width: 40, height: 20 };
  const { client, engine, initial } = await openWorker({ prepare: engine => {
    engine.applyRawOps('body', [
      { op: 'insertEmbed', index: 1, kind: 'image', payload: { ...image, id: 'top-bottom-image', wrapType: 'topAndBottom', displayMode: 'block',
        position: { behindDoc: false, horizontal: { relativeTo: 'column', posOffset: 0 }, vertical: { relativeTo: 'paragraph', posOffset: 0 } } } },
    ]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
  } });
  try {
    const image = initial.frame.payload.geometry!.objects.find(object => object.kind === 'image')!;
    const moved = await client.edit(intent(1, 0, 'moveObject', {
      handle: image.handle, hOffsetEmu: 95250, vOffsetEmu: 190500, expectVersion: initial.frame.payload.documentVersion!,
    }));
    expect(moved.outcomes[0]?.status).toBe('applied');
    const segment = engine.storySegments('body').find(part => part.kind === 'embed' && part.payload.id === 'top-bottom-image')!;
    expect(segment).toMatchObject({ payload: { wrapType: 'square', displayMode: 'float', position: {
      horizontal: { relativeTo: 'page', posOffset: 95250 }, vertical: { relativeTo: 'page', posOffset: 190500 },
    } } });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker cell and text handles survive input and resize the issued cell column', async () => {
  const { client, engine, initial, point } = await openWorker({ bodyXml:
    '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  });
  try {
    const objects = initial.frame.payload.geometry!.objects;
    const cells = objects.filter(object => object.kind === 'cell');
    expect(cells).toHaveLength(2);
    expect(cells.every(cell => cell.rects.length > 0)).toBe(true);
    const text = objects.find(object => object.kind === 'text' && object.story === 'body')!;
    const typed = await client.edit(intent(1, 0, 'insertText', { text: '+', selection: {
      anchor: { ...point, offset: 0 }, focus: { ...point, offset: 0 },
    } }));
    expect(typed.frame!.payload.geometry!.objects.find(object => object.kind === 'text' && object.story === 'body')!.handle).toBe(text.handle);
    expect(typed.frame!.payload.geometry!.objects.filter(object => object.kind === 'cell').map(object => object.handle))
      .toEqual(cells.map(object => object.handle));
    const resized = await client.edit(intent(2, 1, 'resizeObject', {
      handle: cells[1]!.handle, widthTwips: 2160, expectVersion: typed.frame!.payload.documentVersion!,
    }));
    expect(resized.outcomes[0]?.status).toBe('applied');
    expect(engine.tablePayload('body', 0)).toHaveProperty('grid', [1440, 2160]);
    expect(resized.frame!.payload.geometry!.objects.find(object => object.handle === cells[1]!.handle)!.rects[0]!.width)
      .toBeGreaterThan(cells[1]!.rects[0]!.width);
    const denied = await client.edit(intent(3, 2, 'moveObject', {
      handle: text.handle, hOffsetEmu: 0, vOffsetEmu: 0, expectVersion: resized.frame!.payload.documentVersion!,
    }));
    expect(denied.outcomes[0]).toMatchObject({ status: 'refused', reason: 'unsupported-object' });
    const cellRects = (response: typeof resized) => response.frame!.payload.geometry!.objects
      .find(object => object.handle === cells[1]!.handle)!.rects;
    const undone = await client.edit(intent(4, 2, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect(cellRects(undone)).toEqual(cellRects(typed));
    const redone = await client.edit(intent(5, 3, 'redo', {}));
    expect(redone.outcomes[0]?.status).toBe('applied');
    expect(cellRects(redone)).toEqual(cellRects(resized));
  } finally { client.destroy(); }
}, TIMEOUT);

for (const kind of ['insertPageBreak', 'insertSectionBreak'] as const) {
  test(`worker safety refuses ${kind} in a cell before mutation and keeps commands usable`, async () => {
    const { client, engine, point } = await openWorker({ bodyXml: '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>', prepare(engine) {
      const story = 'body:t0:r0c0';
      return { story, paraId: engine.geometryReader.paragraphSpans(story)[0]!.paraId, offset: 4 };
    } });
    try {
      const version = engine.proposalEngine.version();
      const vector = engine.encodeStateVector();
      const at = { ...point, offset: kind === 'insertPageBreak' ? 4 : 0 };
      const refused = await client.edit(intent(1, 0, kind, {
        ...(kind === 'insertSectionBreak' ? { type: 'nextPage' } : {}),
        selection: { anchor: at, focus: at },
      }));
      expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-selection' });
      expect(refused.version).toBe(0);
      expect(engine.proposalEngine.version()).toBe(version);
      expect(engine.encodeStateVector()).toEqual(vector);
      const selected = await client.edit(intent(2, 0, kind, {
        ...(kind === 'insertSectionBreak' ? { type: 'continuous' } : {}),
        selection: { anchor: { ...point, offset: 0 }, focus: point },
      }));
      expect(selected.outcomes[0]?.status).toBe('refused');
      expect(engine.paragraphs(point.story)[0]!.text).toBe('Cell');
      expect((await client.edit(intent(3, 0, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('worker safety resizes the issued cell rather than the table story suffix ordinal', async () => {
  const { client, engine, point } = await openWorker({ bodyXml:
    '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  });
  try {
    const inserted = await client.edit(intent(1, 0, 'insertTable', { rows: 1, columns: 1, selection: { anchor: point, focus: point } }));
    const cell = inserted.frame!.payload.geometry!.objects.find(object => object.kind === 'cell' && object.story === inserted.selection!.focus.story)!;
    expect(cell.story).toMatch(/^body:t1:r0c0~\d+$/);
    const resized = await client.edit(intent(2, 1, 'resizeObject', { handle: cell.handle, widthTwips: 1234, expectVersion: inserted.frame!.payload.documentVersion! }));
    expect(resized.outcomes[0]?.status).toBe('applied');
    expect(engine.tablePayload('body', 0)).toHaveProperty('grid', [1234]);
    expect(engine.tablePayload('body', 1)).toHaveProperty('grid', [1440, 1440]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker safety refuses removed cell handles when a new table reuses the story', async () => {
  const { client, engine } = await openWorker();
  try {
    const inserted = await client.edit(intent(1, 0, 'insertTable', { rows: 1, columns: 1 }));
    const old = inserted.frame!.payload.geometry!.objects.find(object => object.kind === 'cell')!;
    await client.edit(intent(2, 1, 'undo', {}));
    const replacement = await client.edit(intent(3, 2, 'insertTable', { rows: 1, columns: 1 }));
    const current = replacement.frame!.payload.geometry!.objects.find(object => object.kind === 'cell')!;
    expect(current.story).toBe(old.story);
    const before = engine.tablePayload('body', 0);
    const refused = await client.edit(intent(4, 3, 'resizeObject', { handle: old.handle, widthTwips: 1234, expectVersion: replacement.frame!.payload.documentVersion! }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-object' });
    expect(current.handle).not.toBe(old.handle);
    expect(engine.tablePayload('body', 0)).toEqual(before);
    expect((await client.edit(intent(5, 3, 'resizeObject', { handle: current.handle, widthTwips: 1234, expectVersion: replacement.frame!.payload.documentVersion! }))).outcomes[0]?.status).toBe('applied');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker safety refuses malformed proposal anchors atomically without poisoning commands', async () => {
  const { client, engine } = await openWorker();
  try {
    const identities = await client.documentRead({ kind: 'paragraphIdentities' });
    const paragraph = identities.value.paragraphs[0]!.session!;
    const suggest = { author: 'Ada', date: '2026-01-01' };
    const refused = await client.edit(intent(1, 0, 'propose', { expectVersion: identities.version, proposals: [
      { id: 'valid', op: 'insertText', at: 'end', paragraph, text: '!', suggest },
      { id: 'invalid', op: 'insertText', at: 'end', paragraph: {} as typeof paragraph, text: '!', suggest },
    ] }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    expect(engine.proposalEngine.version()).toBe(identities.version);
    expect((await client.documentRead({ kind: 'reviewSnapshot' })).value!.proposals.proposals).toEqual([]);
    expect((await client.edit(intent(2, 0, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker safety never labels an old frame with a queued proposal preview version', async () => {
  const { client, point } = await openWorker();
  try {
    const identities = await client.documentRead({ kind: 'paragraphIdentities' });
    await client.edit(intent(1, 0, 'propose', { expectVersion: identities.version, proposals: [{
      id: 'one', op: 'replaceText', paragraph: identities.value.paragraphs[0]!.session!,
      search: 'Seed', replaceWith: 'Bud', suggest: { author: 'Ada', date: '2026-01-01' },
    }] }));
    const before = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    const decided = client.edit(intent(2, 1, 'setProposalStates', { expectVersion: before.version, expectPreviewVersion: before.previewVersion, changes: [{ id: 'one', state: 'accepted' }] }));
    const read = client.documentRead({ kind: 'reviewSnapshot', expectPreviewVersion: before.previewVersion + 1 });
    const selected = editWithFrame(client, { ...intent(3, 2, 'select', { selection: { anchor: point, focus: point } }), afterSeq: 2 });
    const [decision, snapshot, selection] = await Promise.all([decided, read, selected]);
    expect(decision.outcomes[0]?.status).toBe('applied');
    expect(snapshot.value === null || snapshot.value.frameEpoch > before.frameEpoch).toBe(true);
    const current = (await client.documentRead({ kind: 'reviewSnapshot', expectPreviewVersion: before.previewVersion + 1 })).value!;
    expect(current.frameEpoch).toBe(selection.frame!.revision);
    expect(current.previewVersion).toBe(before.previewVersion + 1);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker safety repairs selection after rejecting a suggested table before ordered typing', async () => {
  const { client, engine } = await openWorker();
  try {
    const inserted = await client.edit(intent(1, 0, 'insertTable', { rows: 1, columns: 1, suggesting: { name: 'Ada', date: '2026-01-01' } }));
    expect(inserted.selection!.focus.story).toMatch(/^body:t0:r0c0~\d+$/);
    const revisionId = engine.geometryReader.listRevisions()[0]!.revisionId;
    const rejected = await client.edit(intent(2, 1, 'rejectRevision', { revisionId, expectVersion: engine.proposalEngine.version() }));
    expect(rejected.outcomes[0]?.status).toBe('applied');
    expect(rejected.selection).toMatchObject({ focus: { story: 'body' } });
    expect((await client.edit({ ...intent(3, 2, 'insertText', { text: '!' }), afterSeq: 2 })).outcomes[0]?.status).toBe('applied');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker withdrawal ignores missing targets and retries after settlement as a no-op', async () => {
  const { client, engine } = await openWorker();
  try {
    const identities = await client.documentRead({ kind: 'paragraphIdentities' });
    await client.edit(intent(1, 0, 'propose', { expectVersion: identities.version, proposals: [{
      id: 'one', op: 'insertText', at: 'end', paragraph: identities.value.paragraphs[0]!.session!,
      text: '!', suggest: { author: 'Ada', date: '2026-01-01' },
    }] }));
    const before = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    const withdrawn = await client.edit(intent(2, 1, 'withdrawProposals', { ids: ['one', 'missing'], expectVersion: before.version }));
    expect(withdrawn.outcomes[0]).toMatchObject({ status: 'applied' });
    const after = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    expect(after.proposals.proposals).toEqual([]);
    expect(after.revisions).toEqual([]);
    const settled = engine.encodeState();
    expect((await client.edit(intent(3, 2, 'withdrawProposals', { ids: ['one'], expectVersion: after.version }))).outcomes[0]?.status).toBe('noop');
    expect(engine.encodeState()).toEqual(settled);
  } finally { client.destroy(); }
}, TIMEOUT);

const OUTER_TABLE = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
const syntheticParagraphs = (count: number) => Array.from({ length: count }, (_, index) =>
  '<w:p w14:paraId="' + (index + 1).toString(16).padStart(8, '0') + '"><w:r><w:t>Seed paragraph ' + index + '</w:t></w:r></w:p>').join('');

test('round two page release retains a matching review preview frame', async () => {
  const { client } = await openWorker({ bodyXml: syntheticParagraphs(120), displayWindow: [0, 2] });
  try {
    const full = await client.requestFullFrame();
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(full.frame!.payload.frame));
    expect(frame.pages[1]!.page.unbuilt).not.toBe(true);
    const before = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    const released = await client.releasePages([{ index: 1, pageId: frame.pages[1]!.pageId.toString() }], frame.frameEpoch);
    expect('superseded' in released).toBe(false);
    const review = (await client.documentRead({ kind: 'reviewSnapshot', expectPreviewVersion: before.previewVersion })).value;
    expect(review).not.toBeNull();
    expect(review!.previewVersion).toBe(before.previewVersion);
    expect(review!.frameEpoch).toBeGreaterThan(before.frameEpoch);
  } finally { client.destroy(); }
}, TIMEOUT);

test('a proposal decision computes its review once for all the reads of its frame', async () => {
  const { client, engine } = await openWorker();
  try {
    const ids = await client.documentRead({ kind: 'paragraphIdentities' });
    await client.edit(intent(1, 0, 'propose', { expectVersion: ids.version, proposals: [{
      id: 'one', op: 'replaceText', paragraph: ids.value.paragraphs[0]!.session!,
      search: 'Seed', replaceWith: 'Bud', suggest: { author: 'Ada', date: '2026-01-01' },
    }] }));
    const before = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    const computed = spyOn(engine, 'commentProjection');
    try {
      for (const [seq, state] of [[2, 'rejected'], [3, 'proposed']] as const) {
        computed.mockClear();
        const decided = await client.edit(intent(seq, seq - 1, 'setProposalStates', {
          expectVersion: before.version, expectPreviewVersion: before.previewVersion + seq - 2, changes: [{ id: 'one', state }],
        }));
        expect(decided.outcomes[0]?.status).toBe('applied');
        const reads = await Promise.all([
          client.documentRead({ kind: 'reviewSnapshot' }),
          client.documentRead({ kind: 'reviewSnapshot', expectPreviewVersion: before.previewVersion + seq - 1 }),
          client.documentRead({ kind: 'reviewSnapshot' }),
        ]);
        expect(reads.map(read => read.value?.proposals.proposals[0]?.state)).toEqual([state, state, state]);
        expect(computed.mock.calls.length).toBeLessThanOrEqual(new Set(reads.map(read => read.value!.frameEpoch)).size);
      }
    } finally { computed.mockRestore(); }
  } finally { client.destroy(); }
}, TIMEOUT);

test('repeated review reads follow proposal decisions, comments, edits and undo', async () => {
  const { client, engine } = await openWorker({ bodyXml: syntheticParagraphs(2) });
  try {
    const review = async () => {
      const [first, second] = await Promise.all([
        client.documentRead({ kind: 'reviewSnapshot' }), client.documentRead({ kind: 'reviewSnapshot' }),
      ]);
      expect(second.value).toEqual(first.value);
      expect(first.value!.version).toBe(engine.proposalEngine.version());
      expect(first.value!.comments).toEqual(engine.listComments());
      expect(first.value!.revisions).toEqual(engine.geometryReader.listRevisions());
      return first.value!;
    };
    const ids = await client.documentRead({ kind: 'paragraphIdentities' });
    await client.edit(intent(1, 0, 'propose', { expectVersion: ids.version, proposals: [{
      id: 'one', op: 'replaceText', paragraph: ids.value.paragraphs[0]!.session!,
      search: 'Seed', replaceWith: 'Bud', suggest: { author: 'Ada', date: '2026-01-01' },
    }] }));
    const proposed = await review();
    expect(proposed.proposals.proposals[0]?.state).toBe('proposed');
    await client.edit(intent(2, 1, 'setProposalStates', {
      expectVersion: proposed.version, expectPreviewVersion: proposed.previewVersion, changes: [{ id: 'one', state: 'rejected' }],
    }));
    const rejected = await review();
    expect(rejected.version).toBe(proposed.version);
    expect(rejected.previewVersion).toBe(proposed.previewVersion + 1);
    expect(rejected.proposals.proposals[0]?.state).toBe('rejected');
    const paraId = engine.geometryReader.paragraphSpans('body')[1]!.paraId;
    const range = { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 4 } };
    await client.edit(intent(3, 2, 'addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Note', expectVersion: rejected.version }));
    const commented = await review();
    expect(commented.comments).toHaveLength(1);
    expect(commented.uiComments).toHaveLength(1);
    const id = commented.comments[0]!.id;
    expect(commented.commentGeometry[id]!.length).toBeGreaterThan(0);
    await client.edit(intent(4, 3, 'updateComment', { id, patch: { done: true }, expectVersion: commented.version }));
    const resolved = await review();
    expect(resolved.comments[0]).toMatchObject({ id, done: true });
    expect(resolved.uiComments?.[0]?.done).toBe(true);
    const at = { story: 'body', paraId, offset: 0 };
    await client.edit(intent(5, 4, 'insertText', {
      text: 'New ', selection: { anchor: at, focus: at }, suggesting: { name: 'Ada', date: '2026-01-01' },
    }));
    const edited = await review();
    expect(edited.version).not.toBe(resolved.version);
    expect(edited.revisions.length).toBeGreaterThan(resolved.revisions.length);
    expect((await client.edit(intent(6, 5, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    const undone = await review();
    expect(undone.version).not.toBe(edited.version);
    expect(undone.revisions.length).toBeLessThan(edited.revisions.length);
  } finally { client.destroy(); }
}, TIMEOUT);

test('round two rejected inner table repairs to the nearest live cell story', async () => {
  const { client, engine } = await openWorker({ bodyXml: OUTER_TABLE, prepare(engine) {
    const story = 'body:t0:r0c0';
    return { story, paraId: engine.geometryReader.paragraphSpans(story)[0]!.paraId, offset: 4 };
  } });
  try {
    await client.edit(intent(1, 0, 'insertTable', { rows: 1, columns: 1, suggesting: { name: 'Ada', date: '2026-01-01' } }));
    const revisionId = engine.geometryReader.listRevisions()[0]!.revisionId;
    const rejected = await client.edit(intent(2, 1, 'rejectRevision', { revisionId, expectVersion: engine.proposalEngine.version() }));
    expect(rejected.selection!.focus.story).toBe('body:t0:r0c0');
    expect((await client.edit({ ...intent(3, 2, 'insertText', { text: '!' }), afterSeq: 2 })).outcomes[0]?.status).toBe('applied');
  } finally { client.destroy(); }
}, TIMEOUT);

for (const kind of ['header', 'footer', 'footnote', 'nested'] as const) {
  test('round two accepted table deletion repairs text and cell selection in ' + kind, async () => {
    const parent = kind === 'nested' ? 'body:t0:r0c0' : kind === 'footnote' ? 'fn:1' : kind === 'header' ? 'hf:rIdHeader' : 'hf:rIdFooter';
    const { client, engine } = await openWorker({ headersFooters: kind === 'header' || kind === 'footer', notes: kind === 'footnote',
      ...(kind === 'nested' ? { bodyXml: OUTER_TABLE } : {}), prepare(engine) {
        const paraId = engine.geometryReader.paragraphSpans(parent)[0]!.paraId;
        const inserted = engine.insertTable({ story: parent, paraId, offset: 0 }, 1, 1);
        const spans = engine.geometryReader.paragraphSpans(parent);
        const owner = spans.find((_, index) => engine.paragraphSegments(parent, [index])[0]!.some(segment => segment.kind === 'embed' && segment.embedKind === 'table'))!;
        engine.deleteRange({ story: parent, start: { paraId: owner.paraId, offset: 0 }, end: { paraId: owner.paraId, offset: 1 } }, { name: 'Ada', date: '2026-01-01' });
        const cell = { ...inserted.table, row: 0, column: 0 };
        engine.setCellSelection({ anchor: cell, head: cell });
        const story = inserted.createdStoryIds[0]!;
        return { story, paraId: engine.geometryReader.paragraphSpans(story)[0]!.paraId, offset: 0 };
      } });
    try {
      const revisionId = engine.geometryReader.listRevisions()[0]!.revisionId;
      const accepted = await client.edit(intent(1, 0, 'acceptRevision', { revisionId, expectVersion: engine.proposalEngine.version() }));
      expect(accepted.outcomes[0]?.status).toBe('applied');
      expect(accepted.selection!.focus.story).toBe(parent);
      expect(() => engine.cellSelection()).not.toThrow();
      expect((await client.edit({ ...intent(2, 1, 'insertText', { text: '!' }), afterSeq: 1 })).outcomes[0]?.status).toBe('applied');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('round two repair helper works after an external story removal', async () => {
  const actions = await import('./residentCommandActions');
  const repair = (actions as unknown as { repairCommandSelection: (engine: unknown, state: unknown) => boolean }).repairCommandSelection;
  expect(typeof repair).toBe('function');
  const { client, engine, initial } = await openWorker({ bodyXml: OUTER_TABLE, prepare(engine) {
    return { story: 'body:t0:r0c0', paraId: engine.geometryReader.paragraphSpans('body:t0:r0c0')[0]!.paraId, offset: 4 };
  } });
  try {
    const inserted = await client.edit(intent(1, 0, 'insertTable', { rows: 1, columns: 1, suggesting: { name: 'Ada', date: '2026-01-01' } }));
    const state = { readOnly: false, selection: inserted.selection, storedMarks: null, stickyPositions: new Map() };
    engine.rejectChange({ revisionId: engine.geometryReader.listRevisions()[0]!.revisionId });
    expect(repair(engine, state)).toBe(true);
    expect(state.selection!.focus.story).toBe('body:t0:r0c0');
    expect(() => engine.cellSelection()).not.toThrow();
    expect(initial.selection).not.toBeNull();
  } finally { client.destroy(); }
}, TIMEOUT);

test('round two object geometry reads only paragraphs on resident pages', async () => {
  const { readResidentFrameGeometry } = await import('./residentObjectGeometry');
  const { client, engine, initial } = await openWorker({ bodyXml: syntheticParagraphs(1000), displayWindow: [0, 1] });
  try {
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const read = spyOn(engine, 'paragraphSegments');
    try {
      await readResidentFrameGeometry(engine, frame, initial.selection ?? null, 1);
      expect(read.mock.calls.reduce((total, [, indices]) => total + indices.length, 0)).toBeLessThan(100);
    } finally { read.mockRestore(); }
  } finally { client.destroy(); }
}, TIMEOUT);

test('frame geometry reads the resident layout without a second display list or segment reads for text', async () => {
  const { readResidentFrameGeometry } = await import('./residentObjectGeometry');
  const { loadRustDisplayListQueryEngine } = await import('../layout/render/rustDisplayList');
  const { client, engine, initial } = await openWorker({ bodyXml: syntheticParagraphs(200), displayWindow: [0, 1] });
  try {
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const opened = spyOn(await loadRustDisplayListQueryEngine(), 'openDisplayList');
    const segments = spyOn(engine, 'paragraphSegments');
    try {
      const geometry = await readResidentFrameGeometry(engine, frame, initial.selection ?? null, 1);
      expect(geometry.objects.filter(object => object.kind === 'text').length).toBeGreaterThan(10);
      expect(opened.mock.calls).toHaveLength(0);
      expect(segments.mock.calls).toHaveLength(0);
    } finally { opened.mockRestore(); segments.mockRestore(); }
  } finally { client.destroy(); }
}, TIMEOUT);

for (const offset of [0, 4]) {
  test('round two break page insertion persists and undoes at offset ' + offset, async () => {
    const { client, engine, point } = await openWorker();
    try {
      const at = { ...point, offset };
      const inserted = await client.edit(intent(1, 0, 'insertPageBreak', { selection: { anchor: at, focus: at } }));
      expect(inserted.outcomes[0]?.status).toBe('applied');
      expect(inserted.version).toBe(1);
      expect(inserted.selection!.focus.offset).toBe(1);
      expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(offset ? ['Seed', '', ''] : ['Seed']);
      const { parseDocx } = await import('../docx/parser');
      const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
      expect(JSON.stringify(reopened.package.document)).toContain('"breakType":"page"');
      expect((await client.edit(intent(2, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed']);
      expect((await client.edit(intent(3, 2, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('a page break at the end of the document keeps two pages after saving and reopening', async () => {
  const text = 'Seed';
  const { client, engine, init } = await openWorker({
    bodyXml: `<w:p w14:paraId="00000001">${textRun(text)}</w:p>`,
  });
  const reopened = await createResidentEngineSession();
  try {
    const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: text.length };
    const inserted = await client.edit(intent(1, 0, 'insertPageBreak', { selection: { anchor: at, focus: at } }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(JSON.parse(engine.retainedLayoutJson()).layout.pages).toHaveLength(2);
    const bytes = (await client.save({ comments: [] })).bytes;
    const paragraphs = savedBody(bytes).match(/<w:p\b[^>]*(?:\/>|>[\s\S]*?<\/w:p>)/g)!;
    expect(paragraphs.at(-2)).toContain('<w:br w:type="page"/>');
    expect(paragraphs.at(-1)).not.toMatch(/<w:(?:r|hyperlink|sdt)\b/);
    reopened.openDocx(new Uint8Array(bytes));
    const fontBytes = init.fonts[0]!;
    if (!(fontBytes instanceof Uint8Array)) throw new Error('Font bytes unavailable');
    const font = reopened.registerFont(fontBytes);
    const request = JSON.parse(init.layoutInput);
    request.measurement.fontChains = { 'liberation sans|0|0': [font] };
    reopened.layoutDocumentWithRegionsRetained(JSON.stringify(request));
    expect(JSON.parse(reopened.retainedLayoutJson()).layout.pages).toHaveLength(2);
  } finally { reopened.destroy(); client.destroy(); }
}, TIMEOUT);

test('round two break commands refuse a cell before mutation', async () => {
  const { client, engine } = await openWorker({ bodyXml: OUTER_TABLE, prepare(engine) {
    const story = 'body:t0:r0c0';
    return { story, paraId: engine.geometryReader.paragraphSpans(story)[0]!.paraId, offset: 4 };
  } });
  try {
    const before = engine.proposalEngine.version();
    for (const [index, kind] of ['insertPageBreak', 'insertSectionBreak'].entries()) {
      const result = await client.edit(intent(index + 1, 0, kind!, kind === 'insertSectionBreak' ? { type: 'nextPage' } : {}));
      expect(result.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-selection' });
      expect(engine.proposalEngine.version()).toBe(before);
    }
  } finally { client.destroy(); }
}, TIMEOUT);

for (const type of ['nextPage', 'continuous'] as const) {
  test('round two break section insertion persists and undoes for ' + type, async () => {
    const { client, engine } = await openWorker();
    try {
      const inserted = await client.edit(intent(1, 0, 'insertSectionBreak', { type }));
      expect(inserted.outcomes[0]?.status).toBe('applied');
      expect(inserted.version).toBe(1);
      expect(inserted.selection!.focus.offset).toBe(0);
      expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed', '']);
      const sections = (await client.documentRead({ kind: 'sectionProperties' })).value.sections;
      expect(sections).toHaveLength(2);
      expect(sections.map(section => section.properties.sectionStart ?? null)).toEqual([null, type]);
      const { parseDocx } = await import('../docx/parser');
      const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
      expect(reopened.package.document.sections).toHaveLength(2);
      expect(reopened.package.document.sections!.map(section => section.properties.sectionStart ?? null)).toEqual([null, type]);
      const undo = await client.edit(intent(2, 1, 'undo', {}));
      expect(undo.outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed']);
      expect(engine.sectionProperties().sections).toHaveLength(1);
      expect((await client.edit(intent(3, 2, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('round two break sections preserve current and existing closing properties', async () => {
  const { client, engine } = await openWorker({ bodyXml: '<w:p w14:paraId="00000001"><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgMar w:left="1800"/></w:sectPr></w:pPr><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>', prepare(engine) {
    const paraId = engine.geometryReader.paragraphSpans('body')[0]!.paraId;
    engine.setSectionMargins(paraId, { marginLeft: 1800 });
    return { story: 'body', paraId, offset: 4 };
  } });
  try {
    const inserted = await client.edit(intent(1, 0, 'insertSectionBreak', { type: 'nextPage' }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    const sections = engine.sectionProperties().sections;
    expect(sections).toHaveLength(3);
    expect(sections[0]!.properties).toMatchObject({ sectionStart: 'continuous', marginLeft: 1800 });
    expect(sections[1]!.properties).toMatchObject({ sectionStart: 'nextPage', marginLeft: 1800 });
    const { parseDocx } = await import('../docx/parser');
    expect((await parseDocx((await client.save({ comments: [] })).bytes)).package.document.sections).toHaveLength(3);
    expect((await client.edit(intent(2, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed', 'Tail']);
    expect(engine.sectionProperties().sections).toHaveLength(2);
  } finally { client.destroy(); }
}, TIMEOUT);

test('a legacy section break lowers on the worker as a section split', async () => {
  const { client, engine } = await openWorker({ bodyXml: '<w:p w14:paraId="00000001"><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgMar w:left="1800"/></w:sectPr></w:pPr><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>' });
  const remote = await createYrsSession({ clientId: 8311 });
  try {
    remote.loadState(engine.encodeState());
    remote.insertSectionBreak({ story: 'body', paraId: remote.paragraphs('body')[0]!.paraId, offset: 2 }, 'nextPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Se', 'ed', 'Tail']);
    expect(remote.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Se', 'ed', 'Tail']);
    const sections = engine.sectionProperties().sections;
    expect(sections).toHaveLength(3);
    expect(sections[0]!.properties).toMatchObject({ sectionStart: 'continuous', marginLeft: 1800 });
    expect(sections[1]!.properties).toMatchObject({ sectionStart: 'nextPage', marginLeft: 1800 });
    const { parseDocx } = await import('../docx/parser');
    expect((await parseDocx((await client.save({ comments: [] })).bytes)).package.document.sections).toHaveLength(3);
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a legacy section break in the final section copies its properties and keeps its start', async () => {
  const bodyXml = '<w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>';
  const sectPr = '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:type w:val="evenPage"/><w:pgMar w:left="2000"/></w:sectPr>';
  const { client, engine } = await openWorker({ bodyXml, sectPr, headersFooters: true });
  const remote = await createYrsSession({ clientId: 8312 });
  try {
    remote.openDocx(syntheticDocument(true, false, false, bodyXml, sectPr), false);
    remote.loadState(engine.encodeState());
    remote.insertSectionBreak({ story: 'body', paraId: remote.paragraphs('body')[1]!.paraId, offset: 2 }, 'nextPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed', 'Ta', 'il']);
    const sections = engine.sectionProperties().sections;
    expect(sections).toHaveLength(2);
    expect(sections[0]!.properties).toMatchObject({ sectionStart: 'evenPage', marginLeft: 2000 });
    expect(sections[1]!.properties).toMatchObject({ sectionStart: 'nextPage', marginLeft: 2000 });
    const { parseDocx } = await import('../docx/parser');
    expect((await parseDocx((await client.save({ comments: [] })).bytes)).package.headers?.size).toBe(1);
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a resident section break undoes in one step', async () => {
  const { client, engine } = await openWorker({ sectPr: '<w:sectPr><w:pgMar w:left="1800"/></w:sectPr>' });
  try {
    const paraId = engine.paragraphs('body')[0]!.paraId;
    engine.insertSectionBreak({ story: 'body', paraId, offset: 2 }, 'nextPage');
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Se', 'ed']);
    expect(engine.undo()).toBe(true);
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed']);
  } finally { client.destroy(); }
}, TIMEOUT);

test('a legacy section break keeps the closing section margins on both halves', async () => {
  const { client, engine } = await openWorker({ bodyXml: '<w:p><w:pPr><w:sectPr><w:pgMar w:left="1800"/></w:sectPr></w:pPr><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>' });
  const remote = await createYrsSession({ clientId: 8314 });
  try {
    engine.setSectionMargins(engine.sectionProperties().sections[0]!.section, { marginLeft: 2400 });
    remote.loadState(engine.encodeState());
    const paraId = remote.paragraphs('body')[0]!.paraId;
    remote.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, { other: { sectPr: { marginLeft: 1800 } } });
    remote.insertSectionBreak({ story: 'body', paraId, offset: 2 }, 'nextPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    const sections = engine.sectionProperties().sections;
    expect(sections).toHaveLength(3);
    expect(sections[0]!.properties).toMatchObject({ marginLeft: 2400 });
    expect(sections[1]!.properties).toMatchObject({ sectionStart: 'nextPage', marginLeft: 2400 });
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a legacy section break refuses table cells and state-only final sections', async () => {
  const bodyXml = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>';
  const { client, engine } = await openWorker({ bodyXml });
  const remote = await createYrsSession({ clientId: 8315 });
  try {
    remote.loadState(engine.encodeState());
    const cell = remote.paragraphs('body:t0:r0c0');
    expect(cell.map(paragraph => paragraph.text)).toEqual(['Cell']);
    expect(() => remote.insertSectionBreak({ story: 'body:t0:r0c0', paraId: cell[0]!.paraId, offset: 2 }, 'nextPage')).toThrow('body paragraphs');
    expect(remote.paragraphs('body:t0:r0c0')).toHaveLength(1);
    expect(remote.paragraphs('body').map(paragraph => paragraph.text)).toEqual(engine.paragraphs('body').map(paragraph => paragraph.text));
  } finally { remote.destroy(); client.destroy(); }
  const origin = await createYrsSession({ clientId: 8316 });
  const stateOnly = await createYrsSession({ clientId: 8317 });
  try {
    origin.openDocx(syntheticDocument(false, false, false, bodyXml), true);
    stateOnly.loadState(origin.encodeState());
    const tail = stateOnly.paragraphs('body').at(-1)!;
    expect(() => stateOnly.insertSectionBreak({ story: 'body', paraId: tail.paraId, offset: 2 }, 'nextPage')).toThrow('section properties');
    expect(stateOnly.paragraphs('body').map(paragraph => paragraph.text)).toEqual(origin.paragraphs('body').map(paragraph => paragraph.text));
  } finally { stateOnly.destroy(); origin.destroy(); }
}, TIMEOUT);

test('a legacy section break in a final section with an empty sectPr lowers on the worker', async () => {
  const { client, engine } = await openWorker();
  const remote = await createYrsSession({ clientId: 8318 });
  try {
    remote.openDocx(syntheticDocument(), false);
    remote.loadState(engine.encodeState());
    remote.insertSectionBreak({ story: 'body', paraId: remote.paragraphs('body')[0]!.paraId, offset: 2 }, 'nextPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Se', 'ed']);
    const sections = engine.sectionProperties().sections;
    expect(sections).toHaveLength(2);
    expect(sections[1]!.properties).toMatchObject({ sectionStart: 'nextPage' });
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a legacy final-section break gives the closed section the headers it inherits', async () => {
  const bodyXml = '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Seed</w:t></w:r></w:p>';
  const { client, engine } = await openWorker({ bodyXml, sectPr: '<w:sectPr/>', headersFooters: true });
  const remote = await createYrsSession({ clientId: 8319 });
  try {
    remote.openDocx(syntheticDocument(true, false, false, bodyXml, '<w:sectPr/>'), false);
    remote.loadState(engine.encodeState());
    const sectPr = (text: string) => {
      const properties = remote.paragraphs('body').find(paragraph => paragraph.text === text)!.properties;
      return (properties.sectPr ?? properties._sectionProperties) as SectionProperties;
    };
    remote.insertSectionBreak({ story: 'body', paraId: remote.paragraphs('body')[1]!.paraId, offset: 2 }, 'nextPage');
    expect(sectPr('Se').headerReferences).toEqual(sectPr('First').headerReferences);
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.sectionProperties().sections).toHaveLength(3);
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a legacy final-section break after an earlier split keeps each section on the worker', async () => {
  const bodyXml = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Seed</w:t></w:r></w:p>';
  const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
  const { client, engine } = await openWorker({ bodyXml, sectPr });
  const remote = await createYrsSession({ clientId: 8320 });
  try {
    remote.openDocx(syntheticDocument(false, false, false, bodyXml, sectPr), false);
    remote.loadState(engine.encodeState());
    remote.insertSectionBreak({ story: 'body', paraId: remote.paragraphs('body')[0]!.paraId, offset: 2 }, 'continuous');
    const seed = remote.paragraphs('body').find(paragraph => paragraph.text === 'Seed')!;
    remote.insertSectionBreak({ story: 'body', paraId: seed.paraId, offset: 2 }, 'nextPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Fi', 'rst', 'Se', 'ed']);
    const sections = engine.sectionProperties().sections;
    expect(sections.map(section => section.properties.pageWidth)).toEqual([12240, 12240, 11906, 11906]);
    expect(sections.at(-1)!.properties).toMatchObject({ sectionStart: 'nextPage' });
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('page breaks land past leading blocks on both session kinds', async () => {
  const bodyXml = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p><w:p><w:r><w:t>Seed</w:t></w:r></w:p>';
  const { client, engine } = await openWorker({ bodyXml });
  const remote = await createYrsSession({ clientId: 8316 });
  try {
    remote.loadState(engine.encodeState());
    const spans = engine.geometryReader.paragraphSpans('body').length;
    const tail = remote.paragraphs('body').findIndex(paragraph => paragraph.text === 'Tail');
    remote.insertPageBreak({ story: 'body', paraId: remote.paragraphs('body')[tail]!.paraId, offset: 0 });
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.geometryReader.paragraphSpans('body')).toHaveLength(spans);
    expect(engine.paragraphSegments('body', [tail])[0]!.slice(0, 2).map(segment => segment.kind === 'embed' ? segment.embedKind : segment.kind)).toEqual(['table', 'pageBreak']);
    const seed = engine.paragraphs('body').findIndex(paragraph => paragraph.text === 'Seed');
    const target = engine.insertPageBreak({ story: 'body', paraId: engine.paragraphs('body')[seed]!.paraId, offset: 2 });
    expect(engine.geometryReader.paragraphSpans('body')).toHaveLength(spans + 1);
    expect(target).toMatchObject({ paraId: engine.paragraphs('body')[seed + 1]!.paraId, offset: 0 });
    expect(engine.paragraphSegments('body', [seed + 1])[0]![0]).toMatchObject({ kind: 'embed', embedKind: 'pageBreak' });
    expect((await client.save({ comments: [] })).bytes.byteLength).toBeGreaterThan(0);
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a legacy page break inside a paragraph splits it before the break', async () => {
  const { client, engine } = await openWorker({ bodyXml: '<w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>' });
  const remote = await createYrsSession({ clientId: 8313 });
  try {
    remote.loadState(engine.encodeState());
    remote.insertPageBreak({ story: 'body', paraId: remote.paragraphs('body')[0]!.paraId, offset: 2 });
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.geometryReader.paragraphSpans('body')).toHaveLength(3);
    const segments = engine.paragraphSegments('body', [1])[0]!;
    expect(segments[0]).toMatchObject({ kind: 'embed', embedKind: 'pageBreak' });
    expect(segments.flatMap(segment => segment.kind === 'text' ? [segment.text] : []).join('')).toBe('ed');
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('round two break sections remain editable through section property commands', async () => {
  const { client, engine } = await openWorker();
  try {
    expect((await client.edit(intent(1, 0, 'insertSectionBreak', { type: 'nextPage' }))).outcomes[0]?.status).toBe('applied');
    const snapshot = (await client.documentRead({ kind: 'sectionProperties' })).value;
    const properties = await client.edit(intent(2, 1, 'setSectionProperties', {
      section: snapshot.sections[0]!.section, patch: { marginLeft: 1800, sectionStart: 'continuous' }, expectVersion: snapshot.version,
    }));
    if (properties.outcomes[0]?.status !== 'applied') throw new Error(JSON.stringify(properties.outcomes[0]));
    expect(properties.outcomes[0]?.status).toBe('applied');
    expect(engine.sectionProperties().sections[0]!.properties).toMatchObject({ marginLeft: 1800, sectionStart: 'continuous' });
    const { parseDocx } = await import('../docx/parser');
    const saved = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(saved.package.document.sections![0]!.properties).toMatchObject({ marginLeft: 1800, sectionStart: 'continuous' });
    expect((await client.edit(intent(3, 2, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect(engine.sectionProperties().sections[0]!.properties.sectionStart).toBe(snapshot.sections[0]!.properties.sectionStart);
    expect((await client.edit(intent(4, 3, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed']);
  } finally { client.destroy(); }
}, TIMEOUT);

type BreakStart = 'continuous' | 'nextPage' | 'oddPage' | 'evenPage';
const otherStart = (type: BreakStart) => type === 'continuous' ? 'nextPage' : 'continuous';
const sectPrXml = (type: string) => `<w:sectPr><w:type w:val="${type}"/></w:sectPr>`;
const sectionParagraph = (id: number, text: string, sectPr = '') =>
  `<w:p w14:paraId="${id.toString(16).padStart(8, '0')}">${sectPr ? `<w:pPr>${sectPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const SECTION_BREAK_CASES: {
  name: string; paragraph: number; body: (type: BreakStart) => string; sectPr: string;
  pages: Record<BreakStart, string[]>; sections: (type: BreakStart) => [string, string | null][];
}[] = [{
  name: 'the final section', paragraph: 0,
  body: () => sectionParagraph(1, 'BeforeAfter'),
  sectPr: '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
  pages: { continuous: ['BeforeAfter'], nextPage: ['Before', 'After'], oddPage: ['Before', '', 'After'], evenPage: ['Before', 'After'] },
  sections: type => [['Before', null], ['', type]],
}, {
  name: 'a paragraph ending a section', paragraph: 0,
  body: type => sectionParagraph(1, 'BeforeAfter', sectPrXml(otherStart(type))) + sectionParagraph(2, 'Tail'),
  sectPr: sectPrXml('continuous'),
  pages: { continuous: ['BeforeAfterTail'], nextPage: ['Before', 'AfterTail'], oddPage: ['Before', '', 'AfterTail'], evenPage: ['Before', 'AfterTail'] },
  sections: type => [['Before', otherStart(type)], ['After', type], ['', 'continuous']],
}, {
  name: 'the middle of three sections', paragraph: 1,
  body: type => sectionParagraph(1, 'One', sectPrXml('continuous')) + sectionParagraph(2, 'BeforeAfter') +
    sectionParagraph(3, 'Mid', sectPrXml(otherStart(type))) + sectionParagraph(4, 'Three'),
  sectPr: sectPrXml('continuous'),
  pages: { continuous: ['One', 'BeforeAfterMidThree'], nextPage: ['OneBefore', 'AfterMidThree'], oddPage: ['OneBefore', '', 'AfterMidThree'], evenPage: ['OneBefore', 'AfterMidThree'] },
  sections: type => [['One', 'continuous'], ['Before', otherStart(type)], ['Mid', type], ['', 'continuous']],
}];

function pageTexts(frame: RetainedFrame): string[] {
  return frame.displayList.pages.map(({ primitives }) => primitives
    .map(primitive => primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : '').join(''));
}

function savedDocumentXml(bytes: ArrayBuffer): string {
  return new TextDecoder().decode(unzipContainer(new Uint8Array(bytes))['word/document.xml']!);
}

function savedBody(bytes: ArrayBuffer): string {
  const xml = savedDocumentXml(bytes);
  return xml.slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'));
}

function savedSectionStarts(bytes: ArrayBuffer): [string, string | null][] {
  const sectPr = /<w:sectPr\b(?:[^>]*\/>|[\s\S]*?<\/w:sectPr>)/;
  const start = (xml: string) => /<w:type\b[^>]*\bw:val="([^"]*)"/.exec(xml)?.[1] ?? null;
  const sections: [string, string | null][] = [];
  const rest = savedDocumentXml(bytes).replace(/<w:p(?:\s[^>]*[^/>])?>[\s\S]*?<\/w:p>/g, paragraph => {
    const found = sectPr.exec(paragraph)?.[0];
    if (found) sections.push([[...paragraph.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join(''), start(found)]);
    return '';
  });
  const final = sectPr.exec(rest)?.[0];
  return final ? [...sections, ['', start(final)]] : sections;
}

for (const scenario of SECTION_BREAK_CASES) {
  for (const type of ['continuous', 'nextPage', 'oddPage', 'evenPage'] as const) {
    test(`section break ${type} starts the section after the caret in ${scenario.name}`, async () => {
      const open = () => openWorker({ bodyXml: scenario.body(type), sectPr: scenario.sectPr, prepare(engine) {
        return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[scenario.paragraph]!.paraId, offset: 6 };
      } });
      const neutralXml = await open().then(async ({ client, initial }) => {
        try {
          const margins = { section: 'final', margins: { marginLeft: 2880 }, expectVersion: initial.frame.payload.documentVersion! };
          expect((await client.edit(intent(1, 0, 'setSectionMargins', margins))).outcomes[0]?.status).toBe('applied');
          await client.save({ comments: [] });
          expect((await client.edit(intent(2, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
          return savedDocumentXml((await client.save({ comments: [] })).bytes);
        } finally { client.destroy(); }
      });
      const { client, engine, initial } = await open();
      try {
        const original = savedBody((await client.save({ comments: [] })).bytes);
        const inserted = await client.edit(intent(1, 0, 'insertSectionBreak', { type }));
        expect(inserted).toMatchObject({ version: 1, outcomes: [{ status: 'applied' }] });
        const acked = [...inserted.frame!.payload.pageFrames ?? [], inserted.frame!.payload.frame]
          .reduce((frame, delta) => applyFrameDeltaOwned(frame, decodeFrameDelta(delta)), applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame)));
        const frame = applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame));
        const pages = scenario.pages[type];
        expect(pageTexts(acked)).toEqual(pages);
        expect(pageTexts(frame)).toEqual(pages);
        const layout = JSON.parse(engine.retainedLayoutJson()) as { layout: { pages: { parityFiller?: boolean }[] } };
        expect(layout.layout.pages.map(page => page.parityFiller === true)).toEqual(pages.map(text => text === ''));
        const expected = scenario.sections(type);
        const snapshot = (await client.documentRead({ kind: 'sectionProperties' })).value;
        expect(snapshot.sections.map((section): string | null => section.properties.sectionStart ?? null)).toEqual(expected.map(([, start]) => start));
        for (const { properties } of engine.paragraphs('body')) {
          const sectPr = (properties.sectPr ?? properties._sectionProperties) as { sectionStart?: string } | undefined;
          expect(properties.sectionBreakType ?? null).toBe(sectPr?.sectionStart ?? null);
        }
        expect(savedSectionStarts((await client.save({ comments: [] })).bytes)).toEqual(expected);
        const undone = await client.edit(intent(2, 1, 'undo', {}));
        expect(undone).toMatchObject({ version: 2, outcomes: [{ status: 'applied' }], editState: { canUndo: false } });
        const restored = (await client.save({ comments: [] })).bytes;
        expect(savedBody(restored)).toBe(original);
        expect(savedDocumentXml(restored)).toBe(neutralXml);
      } finally { client.destroy(); }
    }, TIMEOUT);
  }
}

const UNTYPED_FOLLOWER_PAGES: Record<BreakStart, string[]> = {
  continuous: ['BeforeAfter', 'Tail'], nextPage: ['Before', 'After', 'Tail'],
  oddPage: ['Before', '', 'After', 'Tail'], evenPage: ['Before', 'After', 'Tail'],
};

for (const type of ['continuous', 'nextPage', 'oddPage', 'evenPage'] as const) {
  test(`section break ${type} leaves a following untyped section on a new page`, async () => {
    const { client } = await openWorker({ bodyXml: sectionParagraph(1, 'BeforeAfter', '<w:sectPr/>') + sectionParagraph(2, 'Tail'),
      sectPr: '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>', prepare(engine) {
        return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 6 };
      } });
    const fullFrame = async () => applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame));
    try {
      expect(pageTexts(await fullFrame())).toEqual(['BeforeAfter', 'Tail']);
      expect((await client.edit(intent(1, 0, 'insertSectionBreak', { type }))).outcomes[0]?.status).toBe('applied');
      expect(pageTexts(await fullFrame())).toEqual(UNTYPED_FOLLOWER_PAGES[type]);
      expect(savedSectionStarts((await client.save({ comments: [] })).bytes)).toEqual([['Before', null], ['After', type], ['', null]]);
    } finally { client.destroy(); }
  }, TIMEOUT);
}

const commentXml = (id: number) => `<w:comment w:id="${id}" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`;
const commented = (id: number, content: string) =>
  `<w:commentRangeStart w:id="${id}"/>${content}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
const textRun = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const savedMarks = (xml: string) => ({
  page: xml.match(/<w:br w:type="page"\/>/g)?.length ?? 0,
  column: xml.match(/<w:br w:type="column"\/>/g)?.length ?? 0,
  ranges: xml.match(/<w:commentRangeStart\b/g)?.length ?? 0,
});

async function sourceComments(bodyXml: string, comments: string) {
  const { parseDocx } = await import('../docx/parser');
  return (await parseDocx(syntheticDocument(false, false, false, bodyXml, undefined, comments).buffer as ArrayBuffer)).package.document.comments ?? [];
}

for (const comment of [false, true]) {
  test(`remote page and column breaks survive worker saves without a local break command${comment ? ' beside a saved comment' : ''}`, async () => {
    const bodyXml = sectionParagraph(1, 'One') + `<w:p w14:paraId="00000002">${comment ? commented(1, textRun('Two')) : textRun('Two')}</w:p>` +
      sectionParagraph(3, 'Three');
    const comments = comment ? commentXml(1) : undefined;
    const { client, engine } = await openWorker({ bodyXml, comments });
    const remote = await createResidentEngineSession();
    try {
      remote.loadState(engine.encodeState());
      const two = remote.geometryReader.paragraphSpans('body')[1]!;
      remote.insertPageBreak({ story: 'body', paraId: two.paraId, offset: 0 });
      const three = remote.geometryReader.paragraphSpans('body').slice(0, 2).reduce((index, span) => index + span.length + 1, 0);
      remote.applyRawOps('body', [{ op: 'insertEmbed', index: three, kind: 'columnBreak', payload: {} }]);
      await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
      const host = comment ? await sourceComments(bodyXml, comments!) : [];
      for (let save = 0; save < 2; save += 1) {
        const xml = savedDocumentXml((await client.save({ comments: host })).bytes);
        expect(savedMarks(xml)).toEqual({ page: 1, column: 1, ranges: comment ? 1 : 0 });
        expect(xml.indexOf('w:type="page"')).toBeLessThan(xml.indexOf('>Two<'));
        expect(xml.indexOf('w:type="column"')).toBeGreaterThan(xml.indexOf('>Two<'));
        expect(xml.indexOf('w:type="column"')).toBeLessThan(xml.indexOf('>Three<'));
      }
    } finally { remote.destroy(); client.destroy(); }
  }, TIMEOUT);
}

for (const kind of ['pageBreak', 'columnBreak'] as const) {
  test(`a remote ${kind} after the final paragraph survives worker saves and a reopen`, async () => {
    const { client, engine } = await openWorker({ bodyXml: sectionParagraph(1, 'Tail') });
    const remote = await createResidentEngineSession();
    const [own, other] = kind === 'pageBreak' ? ['page', 'column'] as const : ['column', 'page'] as const;
    try {
      remote.loadState(engine.encodeState());
      const end = remote.storySegments('body').reduce((index, segment) => index + (segment.kind === 'text' ? segment.text.length : 1), 0);
      remote.applyRawOps('body', [{ op: 'insertEmbed', index: end, kind, payload: {} }]);
      await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
      for (let save = 0; save < 2; save += 1) {
        const bytes = (await client.save({ comments: [] })).bytes;
        const xml = savedDocumentXml(bytes);
        const marks = savedMarks(xml);
        expect([marks[own], marks[other]]).toEqual([1, 0]);
        expect(xml.indexOf(`<w:br w:type="${own}"/>`)).toBeGreaterThan(xml.indexOf('>Tail<'));
        const reopened = await createResidentEngineSession();
        try {
          reopened.openDocx(new Uint8Array(bytes));
          expect(reopened.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === kind)).toHaveLength(1);
        } finally { reopened.destroy(); }
      }
    } finally { remote.destroy(); client.destroy(); }
  }, TIMEOUT);
}

test('a commanded page break and an authored column break survive comment saves', async () => {
  const bodyXml = `<w:p w14:paraId="00000001">${commented(1, textRun('First'))}</w:p>` +
    `<w:p w14:paraId="00000002">${commented(2, textRun('Col') + '<w:r><w:br w:type="column"/></w:r>' + textRun('umn'))}</w:p>` +
    sectionParagraph(3, 'Tail');
  const comments = commentXml(1) + commentXml(2);
  const { client } = await openWorker({ bodyXml, comments, prepare(engine) {
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 2 };
  } });
  try {
    const host = await sourceComments(bodyXml, comments);
    const inserted = await client.edit(intent(1, 0, 'insertPageBreak', {}));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    let version = inserted.version;
    const snapshot = await client.documentRead({ kind: 'reviewSnapshot' });
    const ids = snapshot.value!.comments.map(({ id }) => id);
    expect(ids).toHaveLength(2);
    for (const [index, id] of ids.entries()) {
      const read = await client.documentRead({ kind: 'reviewSnapshot' });
      const updated = await client.edit(intent(index + 2, version, 'updateComment', { id, patch: { body: `Updated ${id}` }, expectVersion: read.version }));
      expect(updated.outcomes[0]?.status).toBe('applied');
      version = updated.version;
      expect(savedMarks(savedDocumentXml((await client.save({ comments: host })).bytes))).toEqual({ page: 1, column: 1, ranges: 2 });
    }
  } finally { client.destroy(); }
}, TIMEOUT);

for (const comment of [true, false]) {
  test(`an authored page break stays in its edited paragraph through worker saves${comment ? ' beside a saved comment' : ''}`, async () => {
    const content = textRun('Left') + '<w:r><w:br w:type="page"/></w:r>' + textRun('Right');
    const bodyXml = `<w:p w14:paraId="00000001">${comment ? commented(1, content) : content}</w:p>` + sectionParagraph(2, 'Tail');
    const comments = comment ? commentXml(1) : undefined;
    const { client, engine } = await openWorker({ bodyXml, comments });
    const texts = (xml: string) => [...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join('');
    try {
      const host = comment ? await sourceComments(bodyXml, comments!) : [];
      const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 1 };
      expect((await client.edit(intent(1, 0, 'insertText', { text: '+', selection: { anchor: at, focus: at } }))).outcomes[0]?.status).toBe('applied');
      for (let save = 0; save < 2; save += 1) {
        const xml = savedDocumentXml((await client.save({ comments: host })).bytes);
        expect(savedMarks(xml)).toEqual({ page: 1, column: 0, ranges: comment ? 1 : 0 });
        const paragraphs = xml.match(/<w:p(?:\s[^>]*[^/>])?>[\s\S]*?<\/w:p>/g) ?? [];
        const owner = paragraphs.find(paragraph => paragraph.includes('w14:paraId="00000001"')) ?? '';
        expect(owner.split('<w:br w:type="page"/>').map(texts)).toEqual(['L+eft', 'Right']);
        expect(paragraphs.find(paragraph => paragraph.includes('>Tail<'))).not.toContain('<w:br');
      }
    } finally { client.destroy(); }
  }, TIMEOUT);
}

async function openAuthoredBreakRegression(comment: boolean, type: 'page' | 'column' = 'page') {
  const content = textRun('Left') + `<w:r><w:br w:type="${type}"/></w:r>` + textRun('Right');
  const bodyXml = `<w:p w14:paraId="00000001">${comment ? commented(1, content) : content}</w:p>` + sectionParagraph(2, 'Tail');
  const comments = comment ? commentXml(1) : undefined;
  const host = comment ? await sourceComments(bodyXml, comments!) : [];
  return { ...await openWorker({ bodyXml, comments }), host };
}

function nativeBreakCounts(engine: Awaited<ReturnType<typeof createResidentEngineSession>>): [number, number] {
  const segments = engine.storySegments('body');
  return ['pageBreak', 'columnBreak'].map(kind => segments.filter(segment =>
    segment.kind === 'embed' && segment.embedKind === kind).length) as [number, number];
}

async function expectAuthoredBreakSaves(
  opened: Awaited<ReturnType<typeof openAuthoredBreakRegression>>,
  expected: [number, number]
) {
  for (let save = 0; save < 2; save += 1) {
    expect(nativeBreakCounts(opened.engine)).toEqual(expected);
    const marks = savedMarks(savedDocumentXml((await opened.client.save({ comments: opened.host })).bytes));
    expect([marks.page, marks.column]).toEqual(expected);
    expect(nativeBreakCounts(opened.engine)).toEqual(expected);
  }
}

for (const comment of [true, false]) {
  const suffix = comment ? ' beside a saved comment' : '';
  test(`a locally deleted authored page break stays deleted after editing its paragraph${suffix}`, async () => {
    const opened = await openAuthoredBreakRegression(comment);
    const { client, engine } = opened;
    try {
      expect(nativeBreakCounts(engine)).toEqual([1, 0]);
      const tail = { story: 'body', paraId: engine.paragraphs('body')[1]!.paraId, offset: 0 };
      const deleted = await client.edit(intent(1, 0, 'deleteForward', { selection: { anchor: tail, focus: tail } }));
      expect(deleted.outcomes[0]?.status).toBe('applied');
      expect(nativeBreakCounts(engine)).toEqual([0, 0]);
      const first = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 1 };
      expect((await client.edit(intent(2, deleted.version, 'insertText', {
        text: '+', selection: { anchor: first, focus: first },
      }))).outcomes[0]?.status).toBe('applied');
      await expectAuthoredBreakSaves(opened, [0, 0]);
    } finally { client.destroy(); }
  });

  test(`an authored page break is written once after Enter splits its paragraph${suffix}`, async () => {
    const opened = await openAuthoredBreakRegression(comment);
    const { client, engine } = opened;
    try {
      expect(nativeBreakCounts(engine)).toEqual([1, 0]);
      const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 2 };
      expect((await client.edit(intent(1, 0, 'paragraphBreak', {
        selection: { anchor: at, focus: at },
      }))).outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body')).toHaveLength(3);
      await expectAuthoredBreakSaves(opened, [1, 0]);
    } finally { client.destroy(); }
  });

  test(`a deleted authored column break is not replayed when its paragraph is edited${suffix}`, async () => {
    const opened = await openAuthoredBreakRegression(comment, 'column');
    const { client, engine } = opened;
    try {
      expect(nativeBreakCounts(engine)).toEqual([0, 1]);
      const tail = { story: 'body', paraId: engine.paragraphs('body')[1]!.paraId, offset: 0 };
      const deleted = await client.edit(intent(1, 0, 'deleteForward', { selection: { anchor: tail, focus: tail } }));
      expect(deleted.outcomes[0]?.status).toBe('applied');
      expect(nativeBreakCounts(engine)).toEqual([0, 0]);
      const first = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 1 };
      expect((await client.edit(intent(2, deleted.version, 'insertText', {
        text: '+', selection: { anchor: first, focus: first },
      }))).outcomes[0]?.status).toBe('applied');
      await expectAuthoredBreakSaves(opened, [0, 0]);
    } finally { client.destroy(); }
  });

  test(`an authored page break is written once through split undo and redo${suffix}`, async () => {
    const opened = await openAuthoredBreakRegression(comment);
    const { client, engine } = opened;
    try {
      expect(nativeBreakCounts(engine)).toEqual([1, 0]);
      const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 2 };
      expect((await client.edit(intent(1, 0, 'paragraphBreak', {
        selection: { anchor: at, focus: at },
      }))).outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body')).toHaveLength(3);
      await expectAuthoredBreakSaves(opened, [1, 0]);
      const beforeUndo = await client.requestFullFrame();
      expect((await client.edit(intent(2, beforeUndo.version, 'undo', {}))).outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body')).toHaveLength(2);
      await expectAuthoredBreakSaves(opened, [1, 0]);
      const beforeRedo = await client.requestFullFrame();
      expect((await client.edit(intent(3, beforeRedo.version, 'redo', {}))).outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body')).toHaveLength(3);
      await expectAuthoredBreakSaves(opened, [1, 0]);
    } finally { client.destroy(); }
  });

  test(`a peer-deleted authored page break stays deleted after a local paragraph edit${suffix}`, async () => {
    const opened = await openAuthoredBreakRegression(comment);
    const { client, engine } = opened;
    const remote = await createResidentEngineSession();
    try {
      remote.loadState(engine.encodeState());
      expect(nativeBreakCounts(remote)).toEqual([1, 0]);
      const segments = remote.storySegments('body');
      const embed = segments.findIndex(segment => segment.kind === 'embed' && segment.embedKind === 'pageBreak');
      expect(embed).toBeGreaterThanOrEqual(0);
      const index = segments.slice(0, embed).reduce((offset, segment) =>
        offset + (segment.kind === 'text' ? segment.text.length : 1), 0);
      remote.applyRawOps('body', [{ op: 'delete', index, len: 1 }]);
      expect(nativeBreakCounts(remote)).toEqual([0, 0]);
      await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
      expect(nativeBreakCounts(engine)).toEqual([0, 0]);
      const basis = await client.requestFullFrame();
      const first = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 1 };
      expect((await client.edit(intent(1, basis.version, 'insertText', {
        text: '+', selection: { anchor: first, focus: first },
      }))).outcomes[0]?.status).toBe('applied');
      await expectAuthoredBreakSaves(opened, [0, 0]);
    } finally { remote.destroy(); client.destroy(); }
  });
}

test('a source-recorded page break survives deletion undo and a later comment-paragraph edit', async () => {
  const opened = await openAuthoredBreakRegression(true);
  const { client, engine } = opened;
  try {
    expect(nativeBreakCounts(engine)).toEqual([1, 0]);
    const tail = { story: 'body', paraId: engine.paragraphs('body')[1]!.paraId, offset: 0 };
    const deleted = await client.edit(intent(1, 0, 'deleteForward', { selection: { anchor: tail, focus: tail } }));
    expect(deleted.outcomes[0]?.status).toBe('applied');
    expect(nativeBreakCounts(engine)).toEqual([0, 0]);
    engine.addUndoBoundary();
    const first = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 1 };
    expect((await client.edit(intent(2, deleted.version, 'insertText', {
      text: '+', selection: { anchor: first, focus: first },
    }))).outcomes[0]?.status).toBe('applied');
    await expectAuthoredBreakSaves(opened, [0, 0]);
    const beforeUndoText = await client.requestFullFrame();
    expect((await client.edit(intent(3, beforeUndoText.version, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    await expectAuthoredBreakSaves(opened, [0, 0]);
    const beforeUndoBreak = await client.requestFullFrame();
    const restored = await client.edit(intent(4, beforeUndoBreak.version, 'undo', {}));
    expect(restored.outcomes[0]?.status).toBe('applied');
    expect(nativeBreakCounts(engine)).toEqual([1, 0]);
    expect((await client.edit(intent(5, restored.version, 'insertText', {
      text: '+', selection: { anchor: first, focus: first },
    }))).outcomes[0]?.status).toBe('applied');
    await expectAuthoredBreakSaves(opened, [1, 0]);
  } finally { client.destroy(); }
});

const sectionXml = (type: string, left: number) => `<w:sectPr><w:type w:val="${type}"/><w:pgMar w:left="${left}"/></w:sectPr>`;

function savedSectionMargins(bytes: ArrayBuffer): [string, string | null, number | null][] {
  const sectPr = /<w:sectPr\b(?:[^>]*\/>|[\s\S]*?<\/w:sectPr>)/;
  const read = (xml: string): [string | null, number | null] => {
    const left = /<w:pgMar\b[^>]*\bw:left="(\d+)"/.exec(xml)?.[1];
    return [/<w:type\b[^>]*\bw:val="([^"]*)"/.exec(xml)?.[1] ?? null, left === undefined ? null : Number(left)];
  };
  const sections: [string, string | null, number | null][] = [];
  const rest = savedDocumentXml(bytes).replace(/<w:p(?:\s[^>]*[^/>])?>[\s\S]*?<\/w:p>/g, paragraph => {
    const found = sectPr.exec(paragraph)?.[0];
    if (found) sections.push([[...paragraph.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join(''), ...read(found)]);
    return '';
  });
  const final = sectPr.exec(rest)?.[0];
  return final ? [...sections, ['', ...read(final)]] : sections;
}

const sectionStarts = (engine: Awaited<ReturnType<typeof openWorker>>['engine']) => engine.sectionProperties().sections
  .map(({ properties }): [string | null, number | null] => [properties.sectionStart ?? null, properties.marginLeft ?? null]);

test('a remote section boundary shift keeps every section its own properties through a local break', async () => {
  const bodyXml = sectionParagraph(1, 'FirstMore', sectionXml('continuous', 1800)) + sectionParagraph(2, 'Second', sectionXml('nextPage', 2800)) +
    sectionParagraph(3, 'Tail');
  const { client, engine } = await openWorker({ bodyXml, sectPr: sectionXml('evenPage', 3800) });
  const remote = await createResidentEngineSession();
  try {
    remote.loadState(engine.encodeState());
    const first = remote.paragraphs('body')[0]!;
    remote.splitSection({ story: 'body', paraId: first.paraId, offset: 5 }, first.properties.sectPr as SectionProperties, 'oddPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['First', 'More', 'Second', 'Tail']);
    expect(sectionStarts(engine)).toEqual([['continuous', 1800], ['oddPage', 1800], ['nextPage', 2800], ['evenPage', 3800]]);
    const at = { story: 'body', paraId: engine.paragraphs('body')[1]!.paraId, offset: 2 };
    const inserted = await client.edit(intent(1, 0, 'insertSectionBreak', { type: 'continuous', selection: { anchor: at, focus: at } }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(sectionStarts(engine)).toEqual([['continuous', 1800], ['oddPage', 1800], ['continuous', 1800], ['nextPage', 2800], ['evenPage', 3800]]);
    expect(savedSectionMargins((await client.save({ comments: [] })).bytes)).toEqual([
      ['First', 'continuous', 1800], ['Mo', 'oddPage', 1800], ['re', 'continuous', 1800], ['Second', 'nextPage', 2800], ['', 'evenPage', 3800],
    ]);
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a remote start type change on an existing closing paragraph survives a local break', async () => {
  const bodyXml = sectionParagraph(1, 'ZeroMore') + sectionParagraph(2, 'Closing', sectionXml('continuous', 1800)) + sectionParagraph(3, 'Tail');
  const { client, engine } = await openWorker({ bodyXml, sectPr: sectionXml('nextPage', 2800) });
  const remote = await createResidentEngineSession();
  try {
    remote.loadState(engine.encodeState());
    const [zero, closing] = remote.paragraphs('body');
    remote.splitSection({ story: 'body', paraId: zero!.paraId, offset: 4 }, closing!.properties.sectPr as SectionProperties, 'oddPage');
    await client.applyRemoteUpdate(remote.encodeStateAsUpdate(engine.encodeStateVector()));
    expect(sectionStarts(engine)).toEqual([['continuous', 1800], ['oddPage', 1800], ['nextPage', 2800]]);
    const at = { story: 'body', paraId: engine.paragraphs('body')[1]!.paraId, offset: 2 };
    const inserted = await client.edit(intent(1, 0, 'insertSectionBreak', { type: 'continuous', selection: { anchor: at, focus: at } }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(sectionStarts(engine)).toEqual([['continuous', 1800], ['oddPage', 1800], ['continuous', 1800], ['nextPage', 2800]]);
    expect(savedSectionMargins((await client.save({ comments: [] })).bytes)).toEqual([
      ['Zero', 'continuous', 1800], ['Mo', 'oddPage', 1800], ['Closing', 'continuous', 1800], ['', 'nextPage', 2800],
    ]);
  } finally { remote.destroy(); client.destroy(); }
}, TIMEOUT);

test('a section break over a selection across a section boundary splits the section the deletion leaves', async () => {
  const bodyXml = sectionParagraph(1, 'Zero') + sectionParagraph(2, 'First', sectionXml('continuous', 1800)) + sectionParagraph(3, 'Second') +
    sectionParagraph(4, 'Third', sectionXml('nextPage', 2800)) + sectionParagraph(5, 'Tail');
  const { client, engine } = await openWorker({ bodyXml, sectPr: sectionXml('evenPage', 3800) });
  try {
    const [zero, , second] = engine.geometryReader.paragraphSpans('body');
    const selection = { anchor: { story: 'body', paraId: zero!.paraId, offset: 2 }, focus: { story: 'body', paraId: second!.paraId, offset: 3 } };
    const deleted = await client.edit(intent(1, 0, 'deleteBackward', { selection }));
    expect(deleted.outcomes[0]?.status).toBe('applied');
    const merged = engine.sectionProperties();
    const body = merged.paragraphs.filter(({ story }) => story === 'body').map(({ paraId }) => paraId);
    const ordinal = body.indexOf(deleted.selection!.focus.paraId);
    const index = merged.sections.findIndex(({ section }) => section === 'final' || body.indexOf(section) >= ordinal);
    const containing = merged.sections[index]!.properties;
    const left = sectionStarts(engine);
    expect((await client.edit(intent(2, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    const inserted = await client.edit(intent(3, 2, 'insertSectionBreak', { type: 'oddPage', selection }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(sectionStarts(engine)).toEqual([...left.slice(0, index),
      [containing.sectionStart ?? null, containing.marginLeft ?? null], ['oddPage', containing.marginLeft ?? null], ...left.slice(index + 1)]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('an edited paragraph keeps an empty section mark through save and reopen', async () => {
  const mark = '<w:pPr><w:sectPr/></w:pPr>';
  const marked = `<w:p w14:paraId="00000001">${mark}<w:r><w:t>Seed</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Tail</w:t></w:r></w:p>`;
  const save = async (bodyXml: string) => {
    const { client } = await openWorker({ bodyXml });
    try {
      expect((await client.edit(intent(1, 0, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
      return (await client.save({ comments: [] })).bytes;
    } finally { client.destroy(); }
  };
  const saved = await save(marked);
  expect(savedDocumentXml(await save(marked.replace(mark, '')))).toBe(savedDocumentXml(saved).replace(mark, ''));
  expect(savedSectionStarts(saved)).toEqual([['Seed!', null], ['', null]]);
  const { parseDocx } = await import('../docx/parser');
  expect((await parseDocx(saved)).package.document.sections).toHaveLength(2);
}, TIMEOUT);

test('section break typing reuses unchanged section projections', async () => {
  const { client, engine } = await openWorker({ bodyXml: syntheticParagraphs(1000), displayWindow: [0, 1] });
  try {
    expect((await client.edit(intent(1, 0, 'insertSectionBreak', { type: 'continuous' }))).outcomes[0]?.status).toBe('applied');
    expect((await client.edit(intent(2, 1, 'insertText', { text: '+' }))).outcomes[0]?.status).toBe('applied');
    const paragraphs = spyOn(engine, 'paragraphs');
    try {
      expect((await client.edit(intent(3, 2, 'insertText', { text: '+' }))).outcomes[0]?.status).toBe('applied');
      expect(paragraphs.mock.calls).toHaveLength(0);
    } finally { paragraphs.mockRestore(); }
  } finally { client.destroy(); }
}, TIMEOUT);

for (const part of ['header', 'footer', 'footnote'] as const) {
  test('round two break commands refuse non-body ' + part + ' stories', async () => {
    const story = part === 'footnote' ? 'fn:1' : part === 'header' ? 'hf:rIdHeader' : 'hf:rIdFooter';
    const { client, engine } = await openWorker({ headersFooters: part !== 'footnote', notes: part === 'footnote', prepare(engine) {
      return { story, paraId: engine.geometryReader.paragraphSpans(story)[0]!.paraId, offset: 0 };
    } });
    try {
      const before = engine.proposalEngine.version();
      for (const [index, kind] of ['insertPageBreak', 'insertSectionBreak'].entries()) {
        const refused = await client.edit(intent(index + 1, 0, kind!, kind === 'insertSectionBreak' ? { type: 'continuous' } : {}));
        expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-selection' });
        expect(engine.proposalEngine.version()).toBe(before);
      }
      expect((await client.edit(intent(3, 0, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test.skipIf(process.env.NATIVE_BREAK_PROBE !== '1')('round two native section insertion probe mirrors the hybrid call', async () => {
  const { client, engine, point, init } = await openWorker();
  try {
    engine.insertSectionBreak(point, 'nextPage');
    const { parseDocx } = await import('../docx/parser');
    try {
      engine.layoutDocumentWithRegionsRetained(init.layoutInput);
      const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
      console.log('NATIVE_SECTION_PROBE', JSON.stringify({ lowers: true, sections: reopened.package.document.sections?.length }));
    } catch (error) {
      console.log('NATIVE_SECTION_PROBE', JSON.stringify({ lowers: false, error: String(error) }));
    }
  } finally { client.destroy(); }
}, TIMEOUT);

test('round two highlight formatting persists through worker save and reopen', async () => {
  const { client, point } = await openWorker();
  try {
    const selected = { anchor: { ...point, offset: 0 }, focus: point };
    expect((await client.edit(intent(1, 0, 'formatSelection', { selection: selected, delta: { highlight: 'yellow' } }))).outcomes[0]?.status).toBe('applied');
    const { parseDocx } = await import('../docx/parser');
    const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(JSON.stringify(reopened.package.document)).toContain('"highlight":"yellow"');
  } finally { client.destroy(); }
}, TIMEOUT);

test('round two highlight save projection retains existing string attributes', async () => {
  const { client } = await openWorker({ prepare(engine) {
    engine.applyRawOps('body', [{ op: 'format', index: 0, len: 4, attrs: { highlight: 'yellow' } }]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 4 };
  } });
  try {
    const { parseDocx } = await import('../docx/parser');
    const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
    expect(JSON.stringify(reopened.package.document)).toContain('"highlight":"yellow"');
  } finally { client.destroy(); }
}, TIMEOUT);

test('round two resident geometry matches the reference walk on built pages', async () => {
  const { readResidentFrameGeometry } = await import('./residentObjectGeometry');
  const { readResidentFrameGeometry: reference } = await import('./residentObjectGeometry.reference');
  const inner = OUTER_TABLE.slice(0, -6);
  const table = OUTER_TABLE.replace('<w:tc>', '<w:tc>' + inner);
  const bodyXml = syntheticParagraphs(100).replace('</w:p>', '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>' + table);
  const { client, engine, initial } = await openWorker({ headersFooters: true, notes: true, objects: true,
    bodyXml, displayWindow: [0, 2] });
  try {
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const before = await reference(engine, frame, initial.selection ?? null, 1);
    const after = await readResidentFrameGeometry(engine, frame, initial.selection ?? null, 1);
    const normalize = (value: typeof before) => ({ ...value,
      objects: value.objects.map(({ handle, ...object }) => object),
    });
    expect(normalize(after)).toEqual(normalize(before));
    expect(after.objects.some(object => object.kind === 'image')).toBe(true);
    expect(after.objects.some(object => object.kind === 'cell')).toBe(true);
    expect(after.objects.some(object => object.story === 'fn:1')).toBe(true);
    const at = initial.selection!.focus;
    const range = { anchor: { ...at, offset: 0 }, focus: at };
    expect(normalize(await readResidentFrameGeometry(engine, frame, range, 1))).toEqual(normalize(await reference(engine, frame, range, 1)));
    const unresolved = { ...frame, displayList: { ...frame.displayList,
      pages: frame.displayList.pages.map(page => ({ ...page,
        header: page.header ? { ...page.header, rId: 'missing' } : undefined,
      })),
    } };
    expect(normalize(await readResidentFrameGeometry(engine, unresolved, range, 1)))
      .toEqual(normalize(await reference(engine, unresolved, range, 1)));
  } finally { client.destroy(); }
}, TIMEOUT);

const WRAPPED = 'Wrapped' + ' echo words'.repeat(40);

test('typing in a wrapped paragraph under an active display window carries echo context', async () => {
  const bodyXml = Array.from({ length: 120 }, (_, index) => '<w:p w14:paraId="' + (index + 1).toString(16).padStart(8, '0') +
    '"><w:r><w:t>' + (index === 2 ? WRAPPED : 'Seed paragraph ' + index) + '</w:t></w:r></w:p>').join('');
  const { client, initial } = await openWorker({ bodyXml, displayWindow: [0, 2],
    prepare: () => ({ story: 'body', paraId: '00000003', offset: 200 }) });
  try {
    await client.requestViewport({ start: 0, end: 3, frameEpoch: initial.frame!.revision });
    const typed = (await client.edit(intent(1, 0, 'insertText', { text: 'x' }))).frame!.payload.echoContext;
    expect(typed).toBeDefined();
    const text = WRAPPED.slice(0, 200) + 'x' + WRAPPED.slice(200);
    expect(text).toContain(typed!.text);
    expect(typed!.text.length).toBeLessThan(text.length);
    expect(typed!.boundaries[typed!.selection.focusGrapheme]!.utf16).toBe(typed!.text.indexOf('x') + 1);
    const removed = (await client.edit(intent(2, 1, 'deleteBackward', {}))).frame!.payload.echoContext;
    expect(removed).toBeDefined();
    expect(WRAPPED).toContain(removed!.text);
  } finally { client.destroy(); }
}, TIMEOUT);

test('resident geometry, object rects and selection after inline controls match the layout', async () => {
  const { createDisplayListQueries } = await import('../layout/render/displayListQueries');
  const { loadRustDisplayListQueryEngine } = await import('../layout/render/rustDisplayList');
  const r = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const p = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;
  const sdt = (properties: string, content: string) => `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const checkbox = (id: number) => sdt(`<w:id w:val="${id}"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>`, r('☐'));
  const bodyXml = p('00000001', checkbox(1) + r('A')) + p('00000002', checkbox(2) + r('B')) +
    sdt('<w:id w:val="3"/><w:tag w:val="block"/>', p('00000003', r('Block'))) +
    p('00000004', r('Text') + sdt('<w:id w:val="4"/><w:tag w:val="name"/><w:text/>', r('Name')) + r('End')) +
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p('00000005', checkbox(5) + r('C'))}</w:tc></w:tr></w:tbl>` +
    p('00000006', r('Tail'));
  const { client, engine, initial } = await openWorker({ bodyXml, objects: true });
  const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
  const queries = createDisplayListQueries(frame.displayList, await loadRustDisplayListQueryEngine());
  try {
    interface Block {
      kind: string; paraId?: string; pmStart?: number; pmEnd?: number; sdtGroups?: { id: string }[];
      runs?: { kind: string; text?: string; pmStart?: number; pmEnd?: number; inlineSdtWidget?: { pos: number } }[];
      rows?: { cells: { blocks: Block[] }[] }[];
    }
    const layout: Block[] = [];
    const walk = (blocks: Block[]): void => {
      for (const block of blocks) {
        if (block.kind === 'paragraph') layout.push(block);
        for (const row of block.rows ?? []) for (const cell of row.cells) walk(cell.blocks);
      }
    };
    walk(engine.yrsBlocksForStory('body') as Block[]);
    expect(layout).toHaveLength(6);
    const located = <T extends { pageIndex: number }>(rects: T[]) => rects.map(rect => ({ ...rect, pageId: frame.pages[rect.pageIndex]!.pageId.toString() }));
    const sorted = (values: unknown[]) => values.map(value => JSON.stringify(value)).sort();

    const geometry = engine.residentGeometry('body', [[0, 10_000]]);
    expect(sorted(geometry.paragraphs.map(({ paraId, from, to }) => ({ paraId, from, to }))))
      .toEqual(sorted(layout.map(block => ({ paraId: block.paraId, from: block.pmStart! + 1, to: block.pmEnd! - 1 }))));
    const storyOf = (paraId: string | undefined) => geometry.paragraphs.find(paragraph => paragraph.paraId === paraId)!.story;
    const control = layout.find(block => block.paraId === '00000003')!;
    expect(geometry.controls.map(({ story, contentStory, from }) => ({ story, contentStory, group: `sdt@${from}` })))
      .toEqual([{ story: 'body', contentStory: storyOf('00000003'), group: control.sdtGroups![0]!.id }]);
    const widgets = layout.flatMap(block => (block.runs ?? []).flatMap(run => run.inlineSdtWidget ? [`${block.paraId}@${run.inlineSdtWidget.pos}`] : []));
    const inline = geometry.paragraphs.flatMap(paragraph => (paragraph.controls ?? []).map(({ from }) => `${paragraph.paraId}@${from}`));
    expect(widgets).toHaveLength(3);
    expect(inline).toHaveLength(4);
    expect(inline).toEqual(expect.arrayContaining(widgets));

    const objects = initial.frame.payload.geometry!.objects;
    expect(sorted(objects.filter(object => object.kind === 'text').map(({ story, rects }) => ({ story, rects }))))
      .toEqual(sorted(layout.map(block => ({ story: storyOf(block.paraId), rects: located(queries.rangeRects(block.pmStart! + 1, block.pmEnd! - 1)) }))
        .filter(object => object.rects.length > 0)));
    const images = layout.flatMap(block => (block.runs ?? []).filter(run => run.kind === 'image').map(run => run.pmStart!));
    expect(images).toHaveLength(2);
    expect(sorted(objects.filter(object => object.kind === 'image').map(object => object.rects)))
      .toEqual(sorted(images.map(position => located([queries.imageByPos(position, 'body')!.rect]))));

    const end = layout.find(block => block.paraId === '00000004')!.runs!.find(run => run.text === 'End')!;
    const selected = await editWithFrame(client, intent(1, 0, 'select', { selection: {
      anchor: { story: 'body', paraId: '00000004', offset: 6 }, focus: { story: 'body', paraId: '00000004', offset: 9 },
    } }));
    const expected = located(queries.rangeRects(end.pmStart!, end.pmEnd!));
    expect(expected.length).toBeGreaterThan(0);
    expect(selected.frame!.payload.geometry!.selection).toEqual(expected);
  } finally { queries.dispose(); client.destroy(); }
}, TIMEOUT);

const TABLE_BODY =
  '<w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
  '<w:p w14:paraId="00000002"><w:r><w:t>After</w:t></w:r></w:p>';

function cellPoint(engine: { paragraphs(story: string): { paraId: string; text: string }[]; geometryReader: { storyIds(): string[] } }, text: string) {
  const story = engine.geometryReader.storyIds()
    .find((story) => story.includes(':t') && engine.paragraphs(story)[0]?.text === text)!;
  const paraId = engine.paragraphs(story)[0]!.paraId;
  return { story, selection: { anchor: { story, paraId, offset: 0 }, focus: { story, paraId, offset: 0 } } };
}

function cellObjects(response: { frame?: DocxCommandFrame | null }) {
  return response.frame!.payload.geometry!.objects.filter((object) => object.kind === 'cell');
}

test('worker orders table structure commands with single version steps duplicates and refusals', async () => {
  const { client, engine, initial } = await openWorker({ bodyXml: TABLE_BODY });
  try {
    expect(cellObjects(initial)).toHaveLength(4);
    const cellA = cellPoint(engine, 'A');
    const selected = await client.edit(intent(1, 0, 'select', { selection: cellA.selection }));
    expect(selected.outcomes[0]?.status).toBe('noop');
    expect(selected.version).toBe(0);
    const inserted = await client.edit(intent(2, 0, 'insertTableColumn', { side: 'right' }));
    expect(inserted).toMatchObject({ version: 1, outcomes: [{ status: 'applied' }] });
    expect(engine.selection()?.head.story).toBe(cellA.story);
    expect(cellObjects(inserted)).toHaveLength(6);
    expect(cellObjects(inserted).map((cell) => cell.rects[0]))
      .not.toEqual(cellObjects(initial).map((cell) => cell.rects[0]));
    const duplicate = await client.edit(intent(2, 0, 'insertTableColumn', { side: 'right' }));
    expect(duplicate.outcomes).toEqual(inserted.outcomes);
    expect(duplicate.version).toBe(1);
    expect(cellObjects(await client.requestFullFrame())).toHaveLength(6);
    const refused = await client.edit(intent(3, 1, 'mergeTableCells', {}));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-selection' });
    expect(refused.version).toBe(1);
    const stale = await client.edit({ ...intent(4, 1, 'insertTableRow', { side: 'below' }), basis: { epoch: 2, version: 1 } });
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-epoch' });
    expect(stale.version).toBe(1);
    expect(engine.proposalEngine.version()).toBe(inserted.frame!.payload.documentVersion!);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker refreshes deleted cell handles and resizes merged cells', async () => {
  const { client, engine, initial } = await openWorker({ bodyXml: TABLE_BODY });
  try {
    const before = cellObjects(initial);
    expect(before).toHaveLength(4);
    const cellB = cellPoint(engine, 'B');
    const deleted = await client.edit(intent(1, 0, 'deleteTableColumn', { selection: cellB.selection }));
    expect(deleted.outcomes[0]?.status).toBe('applied');
    const staleHandle = before.find((cell) => cell.story === cellB.story)!.handle;
    const refused = await client.edit(intent(2, 1, 'resizeObject', {
      handle: staleHandle, widthTwips: 2160, expectVersion: deleted.frame!.payload.documentVersion!,
    }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-object' });
    expect(refused.version).toBe(1);
    const undone = await client.edit(intent(3, 1, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    expect((await client.edit(intent(4, 2, 'select', { selection: cellPoint(engine, 'A').selection }))).outcomes[0]?.status).toBe('noop');
    engine.setCellSelection({
      anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 }, head: { story: 'body', tableIndex: 0, row: 0, column: 1 },
    });
    const merged = await client.edit(intent(5, 2, 'mergeTableCells', {}));
    expect(merged.outcomes[0]?.status).toBe('applied');
    const mergedCell = cellObjects(merged).find((cell) => cell.story === cellPoint(engine, 'A').story)!;
    const resized = await client.edit(intent(6, 3, 'resizeObject', {
      handle: mergedCell.handle, widthTwips: 2160, expectVersion: merged.frame!.payload.documentVersion!,
    }));
    expect(resized.outcomes[0]?.status).toBe('applied');
    expect(engine.tablePayload('body', 0)).toHaveProperty('grid', [2160, 1440]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker saves table cell shading and reopens it from the bytes', async () => {
  const { client, engine } = await openWorker({ bodyXml: TABLE_BODY });
  try {
    const cellA = cellPoint(engine, 'A');
    const shaded = await client.edit(intent(1, 0, 'setTableCellShading', { selection: cellA.selection, color: 'AABBCC' }));
    expect(shaded.outcomes[0]?.status).toBe('applied');
    expect(shaded.editState?.tableContext).toMatchObject({ cellBackgroundColor: 'AABBCC' });
    const saved = await client.save({ comments: [] });
    const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(saved.bytes))['word/document.xml']!);
    expect(xml).toContain('AABBCC');
    const { parseDocx } = await import('../docx');
    const reopened = await parseDocx(saved.bytes, { preloadFonts: false });
    const tables = reopened.package.document.content.filter((block): block is Extract<typeof block, { type: 'table' }> => block.type === 'table');
    expect(tables).toHaveLength(1);
    expect(JSON.stringify(tables[0]!.rows[0]!.cells[0]!.formatting?.shading)).toContain('AABBCC');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker_table_formatting_saves_and_reopens', async () => {
  const { client, engine } = await openWorker({ bodyXml: TABLE_BODY });
  try {
    const cellA = cellPoint(engine, 'A');
    await client.edit(intent(1, 0, 'select', { selection: cellA.selection }));
    const line = { style: 'single', size: 4, color: { rgb: '000000' } };
    const bordered = await client.edit(intent(2, 0, 'setTableBorders', {
      borders: { top: line, bottom: line, left: line, right: line, insideH: line, insideV: line },
    }));
    expect(bordered.outcomes[0]?.status).toBe('applied');
    const styled = await client.edit(intent(3, 1, 'setTableProperties', {
      op: 'patch', patch: { width: 2500, widthType: 'pct', justification: 'center' },
    }));
    expect(styled.outcomes[0]?.status).toBe('applied');
    const duplicate = await client.edit(intent(3, 1, 'setTableProperties', {
      op: 'patch', patch: { width: 2500, widthType: 'pct', justification: 'center' },
    }));
    expect(duplicate.outcomes).toEqual(styled.outcomes);
    const saved = await client.save({ comments: [] });
    const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(saved.bytes))['word/document.xml']!);
    expect(xml).toContain('w:tcBorders');
    expect(xml).toContain('<w:tblW w:w="2500" w:type="pct"/>');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).not.toContain('w:tblStyle');
    const { parseDocx } = await import('../docx');
    const reopened = await parseDocx(saved.bytes, { preloadFonts: false });
    const tables = reopened.package.document.content.filter((block): block is Extract<typeof block, { type: 'table' }> => block.type === 'table');
    expect(tables).toHaveLength(1);
    expect(JSON.stringify(tables[0]!.formatting)).toContain('2500');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker refuses table structure commands in header stories', async () => {
  const { client, engine } = await openWorker({ headersFooters: true });
  try {
    const paragraph = engine.geometryReader.paragraphSpans('hf:rIdHeader')[0]!;
    const at = { story: 'hf:rIdHeader', paraId: paragraph.paraId, offset: 0 };
    const inserted = await client.edit(intent(1, 0, 'insertTable', {
      selection: { anchor: at, focus: at }, rows: 1, columns: 1,
    }));
    expect(inserted.outcomes[0]?.status).toBe('applied');
    expect(inserted.selection?.focus.story.startsWith('hf:rIdHeader')).toBe(true);
    const refused = await client.edit(intent(2, 1, 'insertTableRow', { side: 'below' }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-selection' });
    expect(refused.version).toBe(1);
  } finally { client.destroy(); }
}, TIMEOUT);

const TRACKED_PARAGRAPH = '<w:p w14:paraId="00000003"><w:r><w:t>Old </w:t></w:r>' +
  '<w:ins w:id="9" w:author="A" w:date="2026-10-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>';

test('review target reads resolve a comment by its public id and a revision by its id', async () => {
  const { client, engine } = await openWorker({ bodyXml: SOURCE_COMMENT.body + TRACKED_PARAGRAPH, commentsXml: SOURCE_COMMENT.comments });
  try {
    const [comment] = engine.listComments();
    const [revision] = engine.geometryReader.listRevisions();
    if (!comment || !revision) throw new Error('Source review items unavailable');
    const publicId = engine.commentProjection()!.ids[comment.id]!;
    const first = engine.paragraphs('body')[0]!.paraId;
    const read = async (target: ResidentReviewTarget) => (await client.documentRead({ kind: 'reviewTarget', target })).value;
    expect(await read({ kind: 'comment', id: publicId })).toEqual({ id: comment.id,
      ranges: [{ story: 'body', start: { paraId: first, offset: 0 }, end: { paraId: first, offset: 4 } }] });
    expect(await read({ kind: 'revision', revisionId: revision.revisionId })).toEqual({ id: revision.revisionId,
      ranges: [{ story: revision.story, start: revision.range.start, end: revision.range.end }] });
    expect(await read({ kind: 'comment', id: publicId + 40 })).toBeNull();
    expect(await read({ kind: 'revision', revisionId: 'missing' })).toBeNull();
  } finally { client.destroy(); }
}, TIMEOUT);

test('a reply comment links its parent in one undo step and comment commands report every comment', async () => {
  const { client, engine } = await openWorker({ bodyXml: SOURCE_COMMENT.body, commentsXml: SOURCE_COMMENT.comments });
  try {
    const [parent] = engine.listComments();
    if (!parent) throw new Error('Source comment unavailable');
    const parentId = engine.commentProjection()!.ids[parent.id]!;
    const { ranges } = (await client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: parentId } })).value!;
    const reply = (seq: number, version: number, parent: string) => client.edit(intent(seq, version, 'addComment', {
      ranges, author: 'B', date: '2026-01-02', body: 'Reply', parentId: parent, expectVersion: engine.proposalEngine.version() }));
    const unknown = await reply(1, 0, 'missing');
    expect(unknown.outcomes[0]).toMatchObject({ status: 'refused', reason: 'unknown-comment' });
    const replied = await reply(2, unknown.version, parent.id);
    expect(replied.outcomes[0]?.status).toBe('applied');
    const added = engine.listComments().find(comment => comment.id !== parent.id)!;
    expect(added.parentId).toBe(parent.id);
    const replyId = replied.editState!.comment!.id;
    expect(replied.editState!.comment!.comments.find(comment => comment.id === replyId)).toMatchObject({ parentId, author: 'B' });
    const done = await client.edit(intent(3, replied.version, 'updateComment', { id: parent.id, patch: { done: true }, expectVersion: engine.proposalEngine.version() }));
    expect(done.editState!.commentChange).toMatchObject({ seq: 3, comments: expect.arrayContaining([expect.objectContaining({ id: parentId, done: true })]) });
    expect(done.editState!.comment).toBeUndefined();
    const undone = await client.edit(intent(4, done.version, 'undo', {}));
    expect(undone.outcomes[0]?.status).toBe('applied');
    const undoneReply = await client.edit(intent(5, undone.version, 'undo', {}));
    expect(undoneReply.outcomes[0]?.status).toBe('applied');
    expect(engine.listComments().map(comment => comment.id)).toEqual([parent.id]);
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker saves a reply added through comment commands with its parent and the parent done state', async () => {
  const { parseDocx } = await import('../docx/parser');
  const { client, engine } = await openWorker({ bodyXml: SOURCE_COMMENT.body, commentsXml: SOURCE_COMMENT.comments });
  try {
    const [parent] = engine.listComments();
    if (!parent) throw new Error('Source comment unavailable');
    const parentId = engine.commentProjection()!.ids[parent.id]!;
    const { ranges } = (await client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: parentId } })).value!;
    const replied = await client.edit(intent(1, 0, 'addComment', { ranges, author: 'B', date: '2026-01-02T00:00:00Z', body: 'Reply',
      parentId: parent.id, expectVersion: engine.proposalEngine.version() }));
    const done = await client.edit(intent(2, replied.version, 'updateComment', { id: parent.id, patch: { done: true }, expectVersion: engine.proposalEngine.version() }));
    expect(done.outcomes[0]?.status).toBe('applied');
    const saved = (await parseDocx((await client.save({ comments: [] })).bytes)).package.document.comments ?? [];
    expect(saved.find(comment => comment.id === parentId)).toMatchObject({ done: true });
    expect(saved.find(comment => comment.id === replied.editState!.comment!.id)).toMatchObject({ parentId, author: 'B' });
  } finally { client.destroy(); }
}, TIMEOUT);


test('reopening an imported resolved comment updates the projection and saved extended metadata', async () => {
  const { parseDocx } = await import('../docx/parser');
  const { client, engine } = await openWorker({ bodyXml: RESOLVED_THREAD.bodies[0], commentsXml: RESOLVED_THREAD.comments, commentsExtendedXml: RESOLVED_THREAD.extended,
    prepare: engine => {
      expect(engine.commentProjection()?.comments.find(comment => comment.id === 0)?.done).toBe(true);
      return { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 };
    } });
  try {
    const result = await client.edit(intent(1, 0, 'updateComment', { id: '0', patch: { done: false }, expectVersion: engine.proposalEngine.version() }));
    expect(result.outcomes[0]?.status).toBe('applied');
    expect(result.editState?.commentChange?.comments.find(comment => comment.id === 0)?.done).toBe(false);
    const bytes = (await client.save({ comments: [] })).bytes;
    const comments = (await parseDocx(bytes)).package.document.comments!;
    expect(comments.find(comment => comment.id === 0)?.done ?? false).toBe(false);
    const parts = unzipContainer(new Uint8Array(bytes));
    expect(new TextDecoder().decode(parts['word/commentsExtended.xml'])).toMatch(/w15:paraId="0C000001"[^>]*w15:done="0"/);
    expect((await client.edit(intent(2, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect(engine.commentProjection()?.comments.find(comment => comment.id === 0)?.done).toBe(true);
    expect((await parseDocx((await client.save({ comments: [] })).bytes)).package.document.comments?.find(comment => comment.id === 0)?.done).toBe(true);
  } finally { client.destroy(); }
}, TIMEOUT);

test('an imported reply without range markers inherits its ancestor anchor and supports reply resolve and removal', async () => {
  const { parseDocx } = await import('../docx/parser');
  const { client, engine } = await openWorker({ bodyXml: RESOLVED_THREAD.bodies[0], commentsXml: RESOLVED_THREAD.comments, commentsExtendedXml: RESOLVED_THREAD.extended });
  try {
    const parent = (await client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: 0 } })).value!;
    const reply = (await client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: 1 } })).value;
    expect(reply).toEqual({ id: '1', ranges: parent.ranges });
    const added = await client.edit(intent(1, 0, 'addComment', { ranges: reply!.ranges, parentId: reply!.id, body: 'Nested reply', author: 'C', date: '2026-01-03', expectVersion: engine.proposalEngine.version() }));
    expect(added.outcomes[0]?.status).toBe('applied');
    expect(added.editState?.comment?.comments.find(comment => comment.id === added.editState?.comment?.id)?.parentId).toBe(1);
    const reopened = await client.edit(intent(2, 1, 'updateComment', { id: reply!.id, patch: { done: false }, expectVersion: engine.proposalEngine.version() }));
    expect(reopened.outcomes[0]?.status).toBe('applied');
    expect(reopened.editState?.commentChange?.comments.find(comment => comment.id === 1)?.done).toBe(false);
    const removed = await client.edit(intent(3, 2, 'removeComment', { id: reply!.id, expectVersion: engine.proposalEngine.version() }));
    expect(removed.outcomes[0]?.status).toBe('applied');
    expect(removed.editState?.commentChange?.comments.some(comment => comment.id === 1)).toBe(false);
    expect((await parseDocx((await client.save({ comments: [] })).bytes)).package.document.comments?.some(comment => comment.id === 1)).toBe(false);
    expect((await client.edit(intent(4, 3, 'undo', {}))).outcomes[0]?.status).toBe('applied');
    expect((await client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: 1 } })).value).toEqual(reply);
    const restored = (await parseDocx((await client.save({ comments: [] })).bytes)).package.document.comments?.find(comment => comment.id === 1);
    expect(restored).toMatchObject({ parentId: 0 });
    expect(restored?.done ?? false).toBe(false);
  } finally { client.destroy(); }
}, TIMEOUT);

const note = (id: number, text: string) => ({ id, author: 'Host', date: '2026-02-01T00:00:00Z',
  content: [{ type: 'paragraph' as const, formatting: {}, content: [{ type: 'run' as const, formatting: {}, content: [{ type: 'text' as const, text }] }] }] });

test('a host comment list wins over the worker comment projection on save and keeps native anchors', async () => {
  const { client, engine } = await openWorker({ bodyXml: SOURCE_COMMENT.body, commentsXml: SOURCE_COMMENT.comments });
  try {
    const [source] = engine.listComments();
    const id = engine.commentProjection()!.ids[source!.id]!;
    const part = async (request: Parameters<typeof client.save>[0], name: string) =>
      new TextDecoder().decode(unzipContainer(new Uint8Array((await client.save(request)).bytes))[name] ?? new Uint8Array());
    const hosted = { comments: [note(id, 'Host note'), note(id + 40, 'Host only')], commentsSource: 'host' as const };
    const comments = await part(hosted, 'word/comments.xml');
    expect(comments).toContain('Host note');
    expect(comments).toContain('Host only');
    expect(comments).not.toContain('Parent');
    const anchored = comments.match(/<w:comment [^>]*w:id="(\d+)"(?:(?!<\/w:comment>)[\s\S])*Host note/)?.[1];
    expect(await part(hosted, 'word/document.xml')).toContain(`<w:commentRangeStart w:id="${anchored}"`);
    expect(await part({ comments: [], commentsSource: 'host' }, 'word/comments.xml')).not.toContain('Parent');
    expect(await part({ comments: [note(id, 'Ignored')] }, 'word/comments.xml')).toContain('Parent');
  } finally { client.destroy(); }
}, TIMEOUT);

test('a host comment list saves on the comment ids the worker handed out after a removal', async () => {
  const { client, engine } = await openWorker({ bodyXml: SOURCE_COMMENT.body, commentsXml: SOURCE_COMMENT.comments });
  try {
    const paraId = engine.paragraphs('body')[1]!.paraId;
    const range = { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 5 } };
    const ids: number[] = [];
    for (const [index, body] of ['First', 'Second'].entries()) {
      const added = await client.edit(intent(index + 1, index, 'addComment', { ranges: [range], author: 'A', date: '2026-01-01', body,
        expectVersion: engine.proposalEngine.version() }));
      ids.push(added.editState!.comment!.id);
    }
    const first = (await client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: ids[0]! } })).value!;
    expect((await client.edit(intent(3, 2, 'removeComment', { id: first.id, expectVersion: engine.proposalEngine.version() }))).outcomes[0]?.status).toBe('applied');
    const parts = unzipContainer(new Uint8Array((await client.save({ comments: [note(0, 'Host parent'), note(ids[1]!, 'Host second')], commentsSource: 'host' })).bytes));
    const part = (name: string) => new TextDecoder().decode(parts[name] ?? new Uint8Array());
    const saved = part('word/comments.xml').match(/<w:comment [^>]*w:id="(\d+)"(?:(?!<\/w:comment>)[\s\S])*Host second/)?.[1];
    expect(saved).toBeDefined();
    expect(part('word/document.xml')).toContain(`<w:commentRangeStart w:id="${saved}"`);
  } finally { client.destroy(); }
}, TIMEOUT);

test('a comment selection range follows its issued sticky positions across later edits', async () => {
  const { client, engine, point } = await openWorker();
  try {
    const locs = [{ ...point, offset: 1 }, { ...point, offset: 3 }];
    const [anchor, focus] = (await client.documentRead({ kind: 'stickyAnchors', locs, version: engine.proposalEngine.version() })).value;
    const atStart = { ...point, offset: 0 };
    expect((await client.edit(intent(1, 0, 'insertText', { text: '--', selection: { anchor: atStart, focus: atStart } }))).outcomes[0]?.status).toBe('applied');
    const read = await client.documentRead({ kind: 'commentSelectionRange', selection: {
      anchor: { ...locs[0]!, sticky: anchor!.encoded }, focus: { ...locs[1]!, sticky: focus!.encoded } } });
    expect(read.value).toEqual({ story: point.story, start: { paraId: point.paraId, offset: 3 }, end: { paraId: point.paraId, offset: 5 } });
    expect(read.version).toBe(engine.proposalEngine.version());
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker picture geometry carries renderer paint order across document order and relativeHeight', async () => {
  const image = { src: 'data:image/png;base64,', width: 40, height: 20, displayMode: 'float', cssFloat: 'none' };
  const { client, initial } = await openWorker({ prepare: engine => {
    engine.applyRawOps('body', [
      { op: 'insertEmbed', index: 1, kind: 'image', payload: { ...image, id: 'front', wrapType: 'inFront',
        position: { relativeHeight: 1, horizontal: { relativeTo: 'page', posOffset: 952500 }, vertical: { relativeTo: 'page', posOffset: 476250 } } } },
      { op: 'insertEmbed', index: 2, kind: 'image', payload: { ...image, id: 'high-behind', wrapType: 'behind',
        position: { behindDoc: true, relativeHeight: 9, horizontal: { relativeTo: 'page', posOffset: 952500 }, vertical: { relativeTo: 'page', posOffset: 476250 } } } },
      { op: 'insertEmbed', index: 3, kind: 'image', payload: { ...image, id: 'low-behind', wrapType: 'behind',
        position: { behindDoc: true, relativeHeight: 2, horizontal: { relativeTo: 'page', posOffset: 952500 }, vertical: { relativeTo: 'page', posOffset: 476250 } } } },
    ]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
  } });
  try {
    const images = initial.frame.payload.geometry!.objects.filter(object => object.kind === 'image');
    expect(images).toHaveLength(3);
    expect(images.map(object => object.paintOrder)).toEqual([2, 1, 0]);
    expect(images.map(object => object.rects[0])).toEqual(Array(3).fill(images[0]!.rects[0]));
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker picture geometry marks cell anchors as nested', async () => {
  const { client, initial } = await openWorker({ bodyXml: OUTER_TABLE, prepare: engine => {
    const story = 'body:t0:r0c0';
    engine.applyRawOps(story, [{ op: 'insertEmbed', index: 1, kind: 'image', payload: {
      src: 'data:image/png;base64,', width: 40, height: 20, displayMode: 'float', wrapType: 'inFront',
      position: { horizontal: { relativeTo: 'page', posOffset: 0 }, vertical: { relativeTo: 'page', posOffset: 0 } },
    } }]);
    return { story, paraId: engine.geometryReader.paragraphSpans(story)[0]!.paraId, offset: 0 };
  } });
  try {
    const images = initial.frame.payload.geometry!.objects.filter(object => object.kind === 'image');
    expect(images).toHaveLength(1);
    expect(images[0]).toMatchObject({ floating: true, nested: true });
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker commits floating resize position and dimensions in one undo step', async () => {
  const { client, engine, initial, init } = await openWorker({ prepare: engine => {
    engine.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'image', payload: {
      id: 'resize-position', src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==', width: 40, height: 20, displayMode: 'float', wrapType: 'behind',
      position: { behindDoc: true, horizontal: { relativeTo: 'page', posOffset: 952500 }, vertical: { relativeTo: 'page', posOffset: 476250 } },
    } }]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
  } });
  const reopened = new ResidentEngineWorkerClient(startWorker(8399));
  try {
    const image = (response: { frame?: typeof initial.frame }) => response.frame!.payload.geometry!.objects.find(object => object.kind === 'image')!;
    const original = image(initial).rects[0]!;
    expect(original).toMatchObject({ x: 100, y: 50, width: 40, height: 20 });
    let seq = 0;
    for (const position of [{ hOffsetEmu: 0 }, { vOffsetEmu: 0 }, { hOffsetEmu: 0, vOffsetEmu: Infinity }]) {
      const refused = await client.edit(intent(++seq, 0, 'resizeObject', {
        handle: image(initial).handle, expectVersion: initial.frame.payload.documentVersion!,
        widthEmu: 285750, heightEmu: 142875, ...position,
      }));
      expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
      expect(engine.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === 'image'))
        .toMatchObject({ payload: { width: 40, height: 20 } });
    }
    const resized = await client.edit(intent(++seq, 0, 'resizeObject', {
      handle: image(initial).handle, expectVersion: initial.frame.payload.documentVersion!,
      widthEmu: 285750, heightEmu: 142875, hOffsetEmu: 1047750, vOffsetEmu: 523875,
    }));
    expect(resized.outcomes[0]?.status).toBe('applied');
    expect(image(resized).rects[0]).toMatchObject({ x: 110, y: 55, width: 30, height: 15 });
    const saved = await client.save({ comments: [] });
    await reopened.open(new Uint8Array(saved.bytes), { generation: 'd1-command-worker' });
    const restored = await reopened.commandInit(init);
    expect(image(restored).rects[0]).toMatchObject({ x: 110, y: 55, width: 30, height: 15 });
    const undone = await client.edit(intent(++seq, 1, 'undo', {}));
    expect(image(undone).rects[0]).toEqual(original);
    const redone = await client.edit(intent(++seq, 2, 'redo', {}));
    expect(image(redone).rects[0]).toEqual(image(resized).rects[0]);
  } finally { reopened.destroy(); client.destroy(); }
}, TIMEOUT);

test('worker geometry exposes picture rotation for resize admission', async () => {
  const { client, initial } = await openWorker({ prepare: engine => {
    engine.applyRawOps('body', [{ op: 'insertEmbed', index: 1, kind: 'image', payload: {
      id: 'rotated', src: 'data:image/png;base64,', width: 100, height: 50, transform: 'rotate(90deg)',
    } }]);
    return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[0]!.paraId, offset: 0 };
  } });
  try {
    expect(initial.frame.payload.geometry!.objects.find(object => object.kind === 'image')).toMatchObject({ rotation: 90 });
  } finally { client.destroy(); }
}, TIMEOUT);
