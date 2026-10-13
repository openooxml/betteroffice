import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { applyFrameDeltaOwned, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import { preloadEditWasm } from './wasm/index';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createResidentEngineSession } from './residentEngineSession';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import type { DocxCommandEnvelope, ResidentCommandInitRequest } from './residentCommandProtocol';
import type { ResidentEngineWorkerResponse } from './residentEngineWorkerProtocol';

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

function syntheticDocument(headersFooters = false, notes = false, controls = false, bodyXml?: string, sectPr?: string): Uint8Array {
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
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export async function openWorker(options: { headersFooters?: boolean; layoutWithRegions?: boolean; notes?: boolean; controls?: boolean; propertyImages?: boolean; objects?: boolean; bodyXml?: string; sectPr?: string; displayWindow?: [number, number]; factory?: typeof startWorker; prepare?: (engine: Awaited<ReturnType<typeof createResidentEngineSession>>) => { story: string; paraId: string; offset: number } } = {}) {
  const worker = (options.factory ?? startWorker)(8301);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.open(syntheticDocument(options.headersFooters, options.notes, options.controls, options.bodyXml, options.sectPr), { generation: 'd1-command-worker' });
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

function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages.flatMap(({ primitives }) => primitives)
    .map((primitive) => primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : '').join('');
}

test('worker frames include highlights for keyboard and point range selections', async () => {
  const { client, point } = await openWorker();
  try {
    const selected = await client.edit(intent(1, 0, 'select', {
      selection: { anchor: { ...point, offset: 0 }, focus: point },
    }));
    expect(selected.frame!.payload.selectionRects?.length).toBeGreaterThan(0);
    const rect = selected.frame!.payload.selectionRects![0]!;
    const collapsed = await client.edit(intent(2, 0, 'navigate', { unit: 'character', direction: 'forward', extend: false }));
    expect(collapsed.frame!.payload.selectionRects).toEqual([]);
    const keyboard = await client.edit(intent(3, 0, 'navigate', { unit: 'character', direction: 'backward', extend: true }));
    expect(keyboard.frame!.payload.selectionRects?.length).toBeGreaterThan(0);
    const pointer = await client.edit(intent(4, 0, 'extendSelection', {
      page: { pageId: '1', pageIndex: rect.pageIndex }, x: rect.x, y: rect.y + rect.height / 2,
      layoutRevision: keyboard.frame!.payload.layoutRevision, frameEpoch: keyboard.frame!.revision, clickCount: 1,
      modifiers: { shift: true, alt: false, ctrl: false, meta: false },
    }));
    expect(pointer.frame!.payload.selectionRects?.length).toBeGreaterThan(0);
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
    expect(table.selection?.focus.story).toMatch(/:t\d+:r0c0$/);
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
    expect(geometry).toMatchObject({ version: initial.frame.payload.documentVersion, frameEpoch: initial.frame!.revision });
    const images = geometry.objects.filter(object => object.kind === 'image');
    expect(images).toHaveLength(2);
    expect(new Set(images.map(object => object.handle)).size).toBe(2);
    expect(images.every(object => object.rects.length > 0)).toBe(true);
    const selected = await client.edit(intent(1, 0, 'select', { selection: {
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
    expect(cell.story).toBe('body:t1:r0c0');
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
    const selected = client.edit({ ...intent(3, 2, 'select', { selection: { anchor: point, focus: point } }), afterSeq: 2 });
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
    expect(inserted.selection!.focus.story).toBe('body:t0:r0c0');
    const revisionId = engine.geometryReader.listRevisions()[0]!.revisionId;
    const rejected = await client.edit(intent(2, 1, 'rejectRevision', { revisionId, expectVersion: engine.proposalEngine.version() }));
    expect(rejected.outcomes[0]?.status).toBe('applied');
    expect(rejected.selection).toMatchObject({ focus: { story: 'body' } });
    expect((await client.edit({ ...intent(3, 2, 'insertText', { text: '!' }), afterSeq: 2 })).outcomes[0]?.status).toBe('applied');
  } finally { client.destroy(); }
}, TIMEOUT);

test('worker safety refuses unknown withdrawal targets before settling any proposal', async () => {
  const { client, engine } = await openWorker();
  try {
    const identities = await client.documentRead({ kind: 'paragraphIdentities' });
    await client.edit(intent(1, 0, 'propose', { expectVersion: identities.version, proposals: [{
      id: 'one', op: 'insertText', at: 'end', paragraph: identities.value.paragraphs[0]!.session!,
      text: '!', suggest: { author: 'Ada', date: '2026-01-01' },
    }] }));
    const before = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    const refused = await client.edit(intent(2, 1, 'withdrawProposals', { ids: ['one', 'missing'], expectVersion: before.version }));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'unknown-proposal' });
    expect(engine.proposalEngine.version()).toBe(before.version);
    const after = (await client.documentRead({ kind: 'reviewSnapshot' })).value!;
    expect(after.proposals).toEqual(before.proposals);
    expect(after.revisions).toEqual(before.revisions);
    expect((await client.edit(intent(3, 1, 'withdrawProposals', { ids: ['one'], expectVersion: before.version }))).outcomes[0]?.status).toBe('applied');
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

for (const offset of [0, 4]) {
  test('round two break page insertion persists and undoes at offset ' + offset, async () => {
    const { client, engine, point } = await openWorker();
    try {
      const at = { ...point, offset };
      const inserted = await client.edit(intent(1, 0, 'insertPageBreak', { selection: { anchor: at, focus: at } }));
      expect(inserted.outcomes[0]?.status).toBe('applied');
      expect(inserted.version).toBe(1);
      expect(inserted.selection!.focus.offset).toBe(1);
      expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(offset ? ['Seed', ''] : ['Seed']);
      const { parseDocx } = await import('../docx/parser');
      const reopened = await parseDocx((await client.save({ comments: [] })).bytes);
      expect(JSON.stringify(reopened.package.document)).toContain('"breakType":"page"');
      expect((await client.edit(intent(2, 1, 'undo', {}))).outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed']);
      expect((await client.edit(intent(3, 2, 'insertText', { text: '!' }))).outcomes[0]?.status).toBe('applied');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

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
      const { client, engine } = await openWorker({ bodyXml: scenario.body(type), sectPr: scenario.sectPr, prepare(engine) {
        return { story: 'body', paraId: engine.geometryReader.paragraphSpans('body')[scenario.paragraph]!.paraId, offset: 6 };
      } });
      try {
        const original = savedBody((await client.save({ comments: [] })).bytes);
        const inserted = await client.edit(intent(1, 0, 'insertSectionBreak', { type }));
        expect(inserted).toMatchObject({ version: 1, outcomes: [{ status: 'applied' }] });
        const frame = applyFrameDeltaOwned(null, decodeFrameDelta((await client.requestFullFrame()).frame!.payload.frame));
        const pages = scenario.pages[type];
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
        expect(savedBody((await client.save({ comments: [] })).bytes)).toBe(original);
      } finally { client.destroy(); }
    }, TIMEOUT);
  }
}

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
