import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as wasm from '../wasm/loader';
import type { SlideDisplayList } from '../types';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods } from './commandProtocol';

const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function xml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function deck(texts = ['alpha', 'beta', 'gamma']): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${texts.map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${texts.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${texts.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${rel}/slide" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
  for (const [index, text] of texts.entries()) {
    const paragraphs = text.split('\n').map((paragraph) => `<a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${xml(paragraph)}</a:t></a:r></a:p>`).join('');
    zip.file(`ppt/slides/slide${index + 1}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

const identity = { session: 'insert-session', incarnation: 'insert-worker' };
const input = { inputTime: 1, undoGroup: 'insert-session' };

test('command_host_does_not_import_react_source', async () => {
  const source = await readFile(new URL('./commandHost.ts', import.meta.url), 'utf8');
  expect(source).not.toMatch(/from ['"][^'"]*pptx-react/);
});

test('comment payload validation refuses without poisoning the session', async () => {
  const session = await open();
  try {
    let state: PresentationCommandAck = session.state;
    let seq = 0;
    const draft = { author: 'Writer', text: 'Discuss', created: '2026-01-01T00:00:00Z' };
    for (const fields of [{ text: '' }, { text: '\u0000' }, { author: '\u0001' }, { initials: '\ud800' }]) {
      state = await session.command(state, ++seq, { kind: 'invoke', payload: {
        ...input, method: 'addComment', args: [session.story.slideId, { ...draft, ...fields }] } });
      expect(state.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload', committed: false });
      expect(state.version).toBe(session.state.version);
    }
    state = await session.command(state, ++seq, { kind: 'invoke', payload: {
      ...input, method: 'setCommentFlavor', args: ['modern'] } });
    state = await session.command(state, ++seq, { kind: 'invoke', payload: {
      ...input, method: 'addComment', args: [session.story.slideId, draft] } });
    expect(state.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const commentId = (state.mutationResult as { commentId: string }).commentId;
    state = await session.command(state, ++seq, { kind: 'invoke', payload: {
      ...input, method: 'replyToComment', args: [commentId, { ...draft, text: '' }] } });
    expect(state.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-payload', committed: false });
    state = await session.command(state, ++seq, { kind: 'invoke', payload: {
      ...input, method: 'replyToComment', args: [commentId, draft] } });
    expect(state.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
  } finally { await session.dispose(); }
});

async function open() {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await wasm.initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  let handle: wasm.PresentationHandle;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const rpc = createSessionClient<PresentationCommandMethods, PresentationCommandEvents>(pair.client,
    { methods: PRESENTATION_COMMAND_METHODS });
  try {
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const state = await bounded(rpc.call.openCommandEditor(new Uint8Array(await bounded(deck(['alpha']))).buffer,
      { ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] }));
    const result = (await bounded(rpc.call.readEditor({ ...identity, seq: 0 }))).result;
    if (!result.ok) throw new Error(result.failure.message);
    return { state, rpc, story: result.stories[0]!, handle: () => handle!,
      command(state: PresentationCommandAck, seq: number, intent: unknown, afterSeq?: number) {
        return bounded(rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(afterSeq === undefined ? {} : { afterSeq }), ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope));
      },
      async dispose() { try { await bounded(rpc.dispose()); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } },
    };
  } catch (error) {
    opened.mockRestore();
    registration.dispose();
    pair.client.close();
    throw error;
  }
}

function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([promise, new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('Command timed out')), 5000);
  })]).finally(() => clearTimeout(timer));
}

const rect = { x: 200000, y: 200000, width: 1500000, height: 900000 };
const png = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a7N8AAAAASUVORK5CYII=', 'base64'));

test('image_roundtrip_and_creation_receipt', async () => {
  const session = await open();
  try {
    const ack = await session.command(session.state, 1, { kind: 'insertPicture', payload: {
      ...input, slideId: session.story.slideId, name: 'picture', rect, contentType: 'image/png', bytes: png.buffer } });
    expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(ack.creation).toMatchObject({ slideId: session.story.slideId, shapeId: ack.selection?.shapeId });
    expect(ack.selection?.kind).toBe('object');
    const saved = session.handle().save();
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const reopened = wasm.openPresentation(saved, { fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }] });
    try { expect(reopened.snapshot().slides[0]!.shapes).toHaveLength(2);
      expect(reopened.layoutSlide(0).primitives.some((primitive) => primitive.kind === 'image')).toBe(true);
    } finally { reopened.dispose(); }
  } finally { await session.dispose(); }
});

test('insert_picture_reply_delivers_media_after_discarded_geometry_frames', async () => {
  const session = await open();
  try {
    const ack = await session.command(session.state, 1, { kind: 'insertPicture', payload: {
      ...input, slideId: session.story.slideId, name: 'picture', rect, contentType: 'image/png', bytes: png.slice().buffer } });
    expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(ack.selection?.kind).toBe('object');
    expect(ack.frame).not.toBeNull();
    const list = JSON.parse(new TextDecoder().decode(ack.frame!.displayList)) as SlideDisplayList;
    const picture = list.primitives.find((primitive) => primitive.kind === 'image');
    expect(picture?.kind).toBe('image');
    if (picture?.kind !== 'image' || !picture.assetId) throw new Error('Expected an inserted picture');
    expect(ack.frame!.media).toEqual([{ assetId: picture.assetId, bytes: png.buffer }]);
    const flushed = await session.rpc.call.flushCommandEditor({ ...identity, seq: 1 });
    expect(flushed.frame!.media).toEqual([]);
    const thumbnail = await session.rpc.call.thumbnailFrame({ ...identity, slideId: session.story.slideId });
    expect(thumbnail.status).toBe('frame');
    if (thumbnail.status !== 'frame') throw new Error('Expected a thumbnail frame');
    expect(thumbnail.frame.media).toEqual([]);
    const undone = await session.command(ack, 2, { kind: 'undo', payload: input });
    const redone = await session.command(undone, 3, { kind: 'redo', payload: input });
    expect(redone.frame!.media).toEqual([]);
    const restored = JSON.parse(new TextDecoder().decode(redone.frame!.displayList)) as SlideDisplayList;
    expect(restored.primitives.find((primitive) => primitive.kind === 'image')).toMatchObject({ assetId: picture.assetId });
  } finally { await session.dispose(); }
});

test('insertion_preflight_refuses_without_mutation', async () => {
  const session = await open();
  try {
    let state: PresentationCommandAck = session.state;
    let seq = 0;
    for (const overrides of [{ slideId: 'deleted' }, { contentType: 'text/plain' }, { bytes: new ArrayBuffer(0) },
      { bytes: new ArrayBuffer(8 * 1024 * 1024 + 1) }, { rect: { ...rect, width: 0 } }]) {
      state = await session.command(state, ++seq, { kind: 'insertPicture', payload: {
        ...input, slideId: session.story.slideId, name: 'picture', rect, contentType: 'image/png', bytes: png.buffer,
        ...overrides } });
      expect(state.outcomes[0]).toMatchObject({ status: 'refused', committed: false });
      expect(state.version).toBe(session.state.version);
    }
    state = await session.command(state, ++seq, { kind: 'setReadOnly', payload: { ...input, readOnly: true } });
    state = await session.command(state, ++seq, { kind: 'insertTextBox', payload: { ...input, slideId: session.story.slideId,
      draft: { name: 'box', rect, text: '', style: {} } } });
    expect(state.outcomes[0]).toMatchObject({ reason: 'read-only', committed: false });
  } finally { await session.dispose(); }
});

test('picture_conversion_stays_in_worker', async () => {
  const session = await open();
  try {
    const zip = await JSZip.loadAsync(await readFile(new URL('../render/fixtures/tiff-image.pptx', import.meta.url)));
    const path = Object.keys(zip.files).find((name) => /\.tiff?$/i.test(name));
    if (!path) throw new Error('TIFF fixture required');
    const bytes = await zip.file(path)!.async('uint8array');
    const result = await session.rpc.call.prepareCommandPicture({ ...identity, seq: 0 }, bytes.slice().buffer);
    expect([...new Uint8Array(result).slice(0, 4)]).toEqual([137, 80, 78, 71]);
    expect(session.handle().version()).toBe(session.state.version);
  } finally { await session.dispose(); }
});
