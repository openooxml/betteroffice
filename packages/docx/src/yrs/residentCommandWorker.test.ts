import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { applyFrameDeltaOwned, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import { preloadEditWasm } from './wasm/index';
import { preloadOpcWasm } from '../wasm/opc';
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

function syntheticDocument(headersFooters = false, notes = false, controls = false, bodyXml?: string): Uint8Array {
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
    (headersFooters ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>' : '<w:sectPr/>') +
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
    parts.set('word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
    parts.set('word/footnotes.xml', toBytes('<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>'));
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function openWorker(options: { headersFooters?: boolean; layoutWithRegions?: boolean; notes?: boolean; controls?: boolean; propertyImages?: boolean; bodyXml?: string } = {}) {
  const worker = startWorker(8301);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.open(syntheticDocument(options.headersFooters, options.notes, options.controls, options.bodyXml), { generation: 'd1-command-worker' });
    const engine = worker.sessions[0]!;
    if (options.propertyImages) engine.applyRawOps('body', [1, 2].map(index => ({
      op: 'insertEmbed', index, kind: 'image', payload: {
        id: 'shared-image', src: 'data:image/png;base64,', widthEmu: 381000, heightEmu: 190500,
      },
    })));
    const read = engine.proposalEngine.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    const point = { story: 'body', paraId: read.paragraphs[0]!.paraId, offset: 4 };
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
    };
    const initial = await client.commandInit(init);
    return { worker, client, engine, init, initial, point };
  } catch (error) {
    client.destroy();
    throw error;
  }
}

function intent(seq: number, version: number, kind: string, payload: DocxCommandEnvelope['payload']): DocxCommandEnvelope {
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
    const layouts = spyOn(engine, 'layoutDocumentWithRegionsRetained');
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
  const layouts = spyOn(engine, 'layoutDocumentWithRegionsRetained');
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
    expect(geometry).toMatchObject({ version, layoutRevision: 1, frameEpoch: initial.frame.revision });
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
