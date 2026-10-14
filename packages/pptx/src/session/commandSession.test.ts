import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient, type SessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure } from '../../../../shared/office-session/types';
import type { PptxReadResult, PptxStoryText } from '../edits';
import type { PresentationCommandFailure } from '../index';
import type { SlideDisplayList } from '../types';
import { DEFAULT_EDITOR_TEXT_STYLE, effectiveStyleFromSelection } from '../textFormatting';
import type { PptxDocument } from '../wasm/generated/pptx_wasm.js';
import * as wasm from '../wasm/loader';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { registerPresentationCommandHandlers } from './commandHost';
import { hitGeometry, indexGeometry, indexHitGeometry, navigateGeometry } from './commandGeometry';
import {
  CommandOutcomeUnknown,
  PRESENTATION_COMMAND_METHODS,
  type PresentationCommandAck,
  type PresentationCommandBasis,
  type PresentationCommandClient,
  type PresentationCommandEnvelope,
  type PresentationCommandEvents,
  type PresentationCommandFrame,
  type PresentationCommandIntent,
  type PresentationCommandMethods,
  type PresentationCommandMotion,
  type PresentationCommandFrameTag,
  type PresentationCommandOpenResult,
  type PresentationCommandOutcome,
  type PresentationCommandSelection,
} from './commandProtocol';

type Content = Extract<PptxReadResult, { ok: true }>;
type Rpc = SessionClient<PresentationCommandMethods, PresentationCommandEvents>;

const identity = { session: 'p1-session', incarnation: 'p1-worker' };
const input = { inputTime: 100, undoGroup: 'p1-intent' };
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

async function fieldDeck(before: string, value: string, after: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await deck([before + value + after]));
  const path = 'ppt/slides/slide1.xml';
  const source = await zip.file(path)!.async('string');
  const properties = '<a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr>';
  zip.file(path, source.replace(/<a:p>.*?<\/a:p>/,
    `<a:p><a:r>${properties}<a:t>${xml(before)}</a:t></a:r><a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum">${properties}<a:t>${xml(value)}</a:t></a:fld><a:r>${properties}<a:t>${xml(after)}</a:t></a:r></a:p>`));
  return zip.generateAsync({ type: 'uint8array' });
}

async function styledDeck(fontSizePt = 42): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await deck(['alpha']));
  const path = 'ppt/slides/slide1.xml';
  const source = await zip.file(path)!.async('string');
  zip.file(path, source.replace(/<a:p>.*?<\/a:p>/,
    `<a:p><a:pPr><a:defRPr b="1" sz="1800"><a:solidFill><a:srgbClr val="111827"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr b="0" sz="${fontSizePt * 100}"><a:solidFill><a:srgbClr val="3478CC"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr><a:t>alpha</a:t></a:r></a:p>`));
  return zip.generateAsync({ type: 'uint8array' });
}

function content(result: PptxReadResult): Content {
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

function selection(story: PptxStoryText, anchor = 0, focus = anchor): PresentationCommandSelection {
  return { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
    anchor, focus, affinity: 'downstream' };
}

function insert(at: PresentationCommandSelection, text: string): PresentationCommandIntent {
  return { kind: 'insertText', payload: { ...input, at: { source: 'confirmed', selection: at }, text } };
}

function envelope(
  state: PresentationCommandAck, seq: number, intent: PresentationCommandIntent,
  basis: PresentationCommandBasis = { epoch: state.epoch, version: state.version }, afterSeq?: number
): PresentationCommandEnvelope {
  return { ...identity, seq, basis, ...(afterSeq === undefined ? {} : { afterSeq }), ...intent };
}

function decode(frame: PresentationCommandFrame): SlideDisplayList {
  return JSON.parse(new TextDecoder().decode(frame.displayList)) as SlideDisplayList;
}

function covering(ack: PresentationCommandAck, slideId: string): PresentationCommandFrame {
  expect(ack.frame).toBeDefined();
  const frame = ack.frame!;
  expect(frame).toMatchObject({ slideId, documentEpoch: ack.epoch,
    incarnation: ack.incarnation, coversThroughSeq: ack.throughSeq, version: ack.version });
  expect(frame.displayList).toBeInstanceOf(ArrayBuffer);
  expect(decode(frame).primitives.some((primitive) => primitive.kind === 'textBox')).toBe(true);
  return frame;
}

async function engineHost() {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async (source) => {
    const { initWasm } = await import('../wasm/loader');
    await initWasm(source ?? await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  const { PptxDocument: Document } = await import('../wasm/generated/pptx_wasm.js');
  const documents: PptxDocument[] = [];
  const handles: wasm.PresentationHandle[] = [];
  const openPresentation = wasm.openPresentation;
  const handled = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    const handle = openPresentation(...args);
    handles.push(handle);
    return handle;
  });
  const original = Document.openCollaborative;
  const opened = spyOn(Document, 'openCollaborative').mockImplementation((...args) => {
    const document = original(...args);
    documents.push(document);
    return document;
  });
  const sentBuffers: ArrayBuffer[] = [];
  const transport: SessionTransport = { ...pair.host, post(message, transfer) {
    if (isHostMessage(message) && message.kind === 'reply' && message.ok &&
      typeof message.value === 'object' && message.value !== null && 'frame' in message.value) {
      const frame = (message.value as PresentationCommandAck).frame;
      if (frame) sentBuffers.push(frame.displayList);
    }
    pair.host.post(message, transfer);
  } };
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(transport, {
    handlers: registration.handlers, policies: registration.policies,
    context: null, onDispose: registration.dispose,
  });
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  return {
    transport: pair.client, Document, documents, sentBuffers, registration,
    options: { ...identity, clientId: 7811,
      fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] },
    document(): PptxDocument {
      expect(documents).toHaveLength(1);
      return documents[0]!;
    },
    handle(): wasm.PresentationHandle {
      expect(handles).toHaveLength(1);
      return handles[0]!;
    },
    dispose() { handled.mockRestore(); opened.mockRestore(); registration.dispose(); pair.client.close(); },
  };
}

async function open(texts?: string[], readOnly = false, bytes?: Uint8Array) {
  const host = await engineHost();
  const rpc: Rpc = createSessionClient<PresentationCommandMethods, PresentationCommandEvents>(
    host.transport, { methods: PRESENTATION_COMMAND_METHODS }
  );
  try {
    const state = await rpc.call.openCommandEditor(new Uint8Array(bytes ?? await deck(texts)).buffer,
      { ...host.options, readOnly });
    if (!state.frame || state.activeSlideId === null) throw new Error('Expected an active slide');
    const fence = (seq: number) => ({ ...identity, seq });
    return {
      ...host, rpc, state: { ...state, frame: state.frame, activeSlideId: state.activeSlideId }, fence,
      read: async (seq = 0) => content((await rpc.call.readEditor(fence(seq))).result),
      dispose: async () => { try { await rpc.dispose(); } finally { host.dispose(); } },
    };
  } catch (error) {
    try { await rpc.dispose(); } finally { host.dispose(); }
    throw error;
  }
}

test('activateSlide orders edits and navigation with covering frames', async () => {
  const host = await open();
  try {
    const story = (await host.read()).stories[0]!;
    const edit = host.rpc.call.command(envelope(host.state, 1, insert(selection(story), 'typed ')));
    const nextId = host.state.slides[1]!.id;
    const next = host.rpc.call.command(envelope(host.state, 2, { kind: 'activateSlide', payload: { slideId: nextId } }));
    const lastId = host.state.slides[2]!.id;
    const last = host.rpc.call.command(envelope(host.state, 3, { kind: 'activateSlide', payload: { slideId: lastId } }));
    expect((await edit).dirtySlideIds).toEqual([story.slideId]);
    expect(await next).toMatchObject({ activeSlideId: nextId, throughSeq: 2, selection: null,
      outcomes: [{ seq: 2, status: 'applied', committed: false }] });
    covering(await next, nextId);
    covering(await last, lastId);
    expect((await host.read(3)).stories[0]!.text).toBe('typed alpha');
    expect(host.handle().version()).toBe((await edit).version);
  } finally { await host.dispose(); }
});

for (const [mode, createClient] of [['command', createPresentationCommandClient],
  ['shared', createSharedPresentationCommandClient]] as const) {
  test(`stale_activation_ack_publishes_media_without_restoring_slide_${mode}`, async () => {
    const zip = await JSZip.loadAsync(await deck());
    const path = 'ppt/slides/slide2.xml';
    const source = await zip.file(path)!.async('string');
    zip.file(path, source.replace('</p:spTree>', '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="200000" y="200000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>'));
    zip.file('ppt/slides/_rels/slide2.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/image" Target="../media/picture.png"/></Relationships>`);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', types.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
    zip.file('ppt/media/picture.png', Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aX1cAAAAASUVORK5CYII='), (value) => value.charCodeAt(0)));
    const host = await engineHost();
    const commandIds = new Set<number>();
    const held: unknown[] = [];
    let hold = true;
    let deliver: (message: unknown) => void = () => {};
    const transport: SessionTransport = { ...host.transport,
      post(message, transfer) {
        if (isClientMessage(message) && message.kind === 'call' &&
          (message.method === 'command' || message.method === 'edit')) commandIds.add(message.id);
        host.transport.post(message, transfer);
      },
      listen(listener) {
        deliver = listener;
        return host.transport.listen((message) => {
          if (hold && isHostMessage(message) && message.kind === 'reply' && commandIds.has(message.id)) held.push(message);
          else listener(message);
        });
      },
    };
    let client: PresentationCommandClient | undefined;
    try {
      client = await createClient(await zip.generateAsync({ type: 'uint8array' }), host.options, transport);
      const nextId = client.state.slides[1]!.id;
      const lastId = client.state.slides[2]!.id;
      const frames: PresentationCommandFrame[] = [];
      client.subscribeFrame((frame) => frames.push(frame));
      const media: PresentationCommandFrame['media'] = [];
      client.subscribeMedia((assets) => media.push(...assets));
      let first: PresentationCommandAck | undefined;
      let second: PresentationCommandAck | undefined;
      void client.dispatch({ kind: 'activateSlide', payload: { slideId: nextId } }).then((ack) => { first = ack; }, () => {});
      void client.dispatch({ kind: 'activateSlide', payload: { slideId: lastId } }).then((ack) => { second = ack; }, () => {});
      for (let index = 0; index < 60 && held.length < 2; index += 1) await new Promise((resolve) => setTimeout(resolve, 0));
      expect(held).toHaveLength(2);
      deliver(held.pop());
      for (let index = 0; index < 60; index += 1) await Promise.resolve();
      expect(second?.throughSeq).toBe(2);
      const confirmed = client.state.frame;
      deliver(held.pop());
      for (let index = 0; index < 60; index += 1) await Promise.resolve();
      expect(first?.frame?.media).toHaveLength(1);
      expect(client.state.frame).toBe(confirmed);
      expect(client.state.activeSlideId).toBe(lastId);
      expect(frames.map((frame) => frame.slideId)).toEqual([lastId]);
      expect(media).toEqual(first!.frame!.media);
      const thumbnail = await client.thumbnail(nextId);
      expect(thumbnail.status).toBe('frame');
      if (thumbnail.status !== 'frame') throw new Error('Expected thumbnail frame');
      expect(thumbnail.frame.media).toEqual([]);
      hold = false;
      const revisited = await client.dispatch({ kind: 'activateSlide', payload: { slideId: nextId } });
      expect(covering(revisited, nextId).media).toEqual([]);
      expect(client.state.frame?.media).toEqual([]);
      const replayed: PresentationCommandFrame['media'] = [];
      const offMedia = client.subscribeMedia((assets) => replayed.push(...assets));
      expect(replayed).toEqual(media);
      offMedia();
    } finally {
      hold = false;
      for (const reply of held.splice(0)) deliver(reply);
      try { await client?.dispose(); } finally { host.dispose(); }
    }
  });
}

test('activateSlide invalidates an old composition', async () => {
  const host = await open();
  try {
    const story = (await host.read()).stories[0]!;
    const payload = { ...input, at: { source: 'confirmed' as const, selection: selection(story) }, compositionId: 'activation-ime' };
    const started = await host.rpc.call.command(envelope(host.state, 1, { kind: 'compositionStart', payload }));
    expect(started.outcomes).toEqual([{ seq: 1, status: 'applied', committed: false }]);
    const nextId = host.state.slides[1]!.id;
    const activated = await host.rpc.call.command(envelope(started, 2, { kind: 'activateSlide', payload: { slideId: nextId } }));
    expect(activated.selection).toBeNull();
    const committed = await host.rpc.call.command(envelope(activated, 3,
      { kind: 'compositionCommit', payload: { ...payload, text: 'preedit' } }));
    expect(committed.outcomes).toEqual([{ seq: 3, status: 'refused', reason: 'invalid-composition', committed: false }]);
    expect(committed.selection).toBeNull();
    covering(committed, nextId);
    expect((await host.read(3)).stories[0]!.text).toBe('alpha');
  } finally { await host.dispose(); }
});

test('activateSlide refuses invalid and deleted identities without changing selection', async () => {
  const host = await open();
  try {
    const story = (await host.read()).stories.find((entry) => entry.slideId === host.state.slides[1]!.id)!;
    await host.rpc.call.command(envelope(host.state, 1, { kind: 'activateSlide', payload: { slideId: story.slideId } }));
    const selected = await host.rpc.call.command(envelope(host.state, 2, { kind: 'setSelection',
      payload: { ...input, selection: selection(story) } }));
    const removedId = host.state.slides[0]!.id;
    host.handle().deleteSlide(removedId);
    for (const [index, slideId] of ['missing', removedId].entries()) {
      const reply = await host.rpc.call.command(envelope(selected, index + 3,
        { kind: 'activateSlide', payload: { slideId } }));
      expect(reply).toMatchObject({ activeSlideId: story.slideId, selection: selected.selection,
        version: selected.version, dirtySlideIds: [],
        outcomes: [{ seq: index + 3, status: 'refused', reason: 'bad-target', committed: false }] });
      expect(covering(reply, story.slideId).slideIndex).toBe(0);
    }
  } finally { await host.dispose(); }
});

test('activateSlide refuses a deleted active slide without a frame', async () => {
  const host = await open();
  try {
    host.handle().deleteSlide(host.state.activeSlideId);
    const reply = await host.rpc.call.command(envelope(host.state, 1,
      { kind: 'activateSlide', payload: { slideId: host.state.activeSlideId } }));
    expect(reply).toMatchObject({ throughSeq: 1, activeSlideId: host.state.activeSlideId, frame: null,
      outcomes: [{ seq: 1, status: 'refused', reason: 'bad-target', committed: false }] });
    const liveId = host.state.slides[1]!.id;
    const activated = await host.rpc.call.command(envelope(reply, 2,
      { kind: 'activateSlide', payload: { slideId: liveId } }));
    expect(activated.outcomes).toEqual([{ seq: 2, status: 'applied', committed: false }]);
    expect(covering(activated, liveId).slideIndex).toBe(0);
  } finally { await host.dispose(); }
});

test('command frames reuse metadata until activation or topology changes', async () => {
  const host = await open();
  const metadata = spyOn(wasm, 'presentationMetadata');
  try {
    const story = (await host.read()).stories[0]!;
    const selected = await host.rpc.call.command(envelope(host.state, 1,
      { kind: 'setSelection', payload: { ...input, selection: selection(story) } }));
    const edited = await host.rpc.call.command(envelope(selected, 2, insert(selection(story), 'typed ')));
    await host.rpc.call.flushCommandEditor(host.fence(2));
    await host.rpc.call.thumbnailFrame({ ...identity, slideId: host.state.slides[1]!.id });
    expect(metadata).not.toHaveBeenCalled();
    const activeId = host.state.slides[2]!.id;
    const activated = await host.rpc.call.command(envelope(edited, 3,
      { kind: 'activateSlide', payload: { slideId: activeId } }));
    expect(metadata).toHaveBeenCalledTimes(1);
    host.handle().deleteSlide(host.state.activeSlideId);
    const refreshed = await host.rpc.call.flushCommandEditor(host.fence(3));
    expect(metadata).toHaveBeenCalledTimes(2);
    expect(covering(refreshed, activeId).slideIndex).toBe(1);
    await host.rpc.call.flushCommandEditor(host.fence(3));
    expect(metadata).toHaveBeenCalledTimes(2);
    expect(activated.outcomes[0]?.status).toBe('applied');
  } finally { metadata.mockRestore(); await host.dispose(); }
});

test('deleted_thumbnail_targets_are_refused_before_revision_short_circuit', async () => {
  const host = await open();
  try {
    const slideId = host.state.slides[1]!.id;
    host.handle().deleteSlide(slideId);
    for (const knownRevision of [undefined, 0]) {
      expect(await host.rpc.call.thumbnailFrame({ ...identity, slideId, knownRevision })).toMatchObject({
        ...identity, epoch: host.state.epoch, slideId, status: 'refused', reason: 'bad-target',
      });
    }
    const activated = await host.rpc.call.command(envelope(host.state, 1,
      { kind: 'activateSlide', payload: { slideId: host.state.slides[2]!.id } }));
    expect(activated.outcomes[0]?.status).toBe('applied');
    covering(activated, host.state.slides[2]!.id);
  } finally { await host.dispose(); }
});

test('s1_initial_bounds_clamp', async () => {
  for (const [activeSlideIndex, slideIndex] of [[-3, 0], [99, 2]]) {
    const host = await engineHost();
    let client: PresentationCommandClient | undefined;
    try {
      client = await createPresentationCommandClient(await deck(), { ...host.options, activeSlideIndex }, host.transport);
      expect(client.state.activeSlideId).toBe(client.state.slides[slideIndex]!.id);
      expect(client.state.frame?.slideIndex).toBe(slideIndex);
      expect(await client.thumbnail('missing')).toMatchObject({ status: 'refused', reason: 'bad-target' });
      expect((await client.read({ slideIds: ['missing'] })).result.ok).toBe(false);
    } finally { try { await client?.dispose(); } finally { host.dispose(); } }
  }
});

test('s1_initial_bounds_reject_noninteger', async () => {
  const host = await engineHost();
  try {
    await expect(createPresentationCommandClient(await deck(), { ...host.options, activeSlideIndex: 0.5 }, host.transport))
      .rejects.toThrow('Active slide index is out of range');
  } finally { host.dispose(); }
});

test('s1_empty_open_read_save_without_frame', async () => {
  const host = await engineHost();
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck([]), { ...host.options, activeSlideIndex: 99 }, host.transport);
    expect(client.state).toMatchObject({ slides: [], activeSlideId: null, frame: null, selection: null });
    expect(host.sentBuffers).toHaveLength(0);
    const frames: PresentationCommandFrame[] = [];
    client.subscribeFrame((frame) => frames.push(frame));
    expect((await client.read()).result).toMatchObject({ ok: true, slides: [], stories: [] });
    expect((await client.read({ slideIds: ['missing'] })).result.ok).toBe(false);
    await client.dispatch({ kind: 'setReadOnly', payload: { ...input, readOnly: true } });
    await client.flush();
    expect(client.state.frame).toBeNull();
    expect(client.state.activeSlideId).toBeNull();
    expect(frames).toHaveLength(0);
    expect((await client.save()).byteLength).toBeGreaterThan(0);
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_open_returns_first_frame_without_peer', async () => {
  const host = await engineHost();
  const replay = spyOn(host.Document, 'openReplayBaseline');
  const peer = spyOn(host.Document, 'openPeerDeckJson');
  const hydrate = spyOn(host.Document.prototype, 'peerHydrationJson');
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck(), host.options, host.transport);
    const initial: PresentationCommandOpenResult = client.state;
    expect(initial).toMatchObject({ ...identity, throughSeq: 0, outcomes: [], dirtySlideIds: [],
      editState: { readOnly: false, canUndo: false, canRedo: false } });
    expect(initial.slides).toHaveLength(3);
    expect(initial.activeSlideId).toBe(initial.slides[0]!.id);
    expect(initial.selection === null || initial.selection.slideId === initial.activeSlideId).toBe(true);
    const frame = covering(initial, initial.activeSlideId!);
    expect(frame.slideRevision).toBeGreaterThanOrEqual(0);
    expect(decode(frame).width).toBeGreaterThan(0);
    expect(decode(frame).primitives.filter((primitive) => primitive.kind === 'textBox')
      .flatMap((box) => box.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text))))
      .toEqual(['alpha']);
    expect(host.documents).toHaveLength(1);
    expect(host.sentBuffers.every((buffer) => buffer.byteLength === 0)).toBe(true);
    expect(replay).not.toHaveBeenCalled();
    expect(peer).not.toHaveBeenCalled();
    expect(hydrate).not.toHaveBeenCalled();
    expect(host.sentBuffers).toHaveLength(1);
  } finally {
    replay.mockRestore(); peer.mockRestore(); hydrate.mockRestore();
    try { await client?.dispose(); } finally { host.dispose(); }
  }
});

test('p1_open_reply_then_transport_failure_rejects_creation', async () => {
  const host = await engineHost();
  const failure = new SessionFailure('crash', 'worker crashed after open reply');
  const delivered: string[] = [];
  let openId: number | undefined;
  let crash: ((error: unknown) => void) | undefined;
  const transport: SessionTransport = { ...host.transport,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'openCommandEditor') {
        openId = message.id;
      }
      host.transport.post(message, transfer);
    },
    listen(listener) {
      return host.transport.listen((message) => {
        listener(message);
        if (isHostMessage(message) && message.kind === 'reply' && message.ok && message.id === openId) {
          delivered.push('open', 'failure');
          crash!(failure);
        }
      });
    },
    onError(listener) { crash = listener; return host.transport.onError(listener); },
  };
  let client: PresentationCommandClient | undefined;
  try {
    const error = await createPresentationCommandClient(await deck(['alpha']), host.options, transport)
      .then((value) => { client = value; }, (error: unknown) => error);
    expect(delivered).toEqual(['open', 'failure']);
    expect(error).toBe(failure);
    expect(error).toMatchObject({ name: 'SessionFailure', code: 'crash' });
    expect(client).toBeUndefined();
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(host.document().canUndo()).toBe(false);
    expect(host.sentBuffers).toHaveLength(1);
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_initial_read_content_refusal_fails_open_without_empty_story_cache', async () => {
  const host = await engineHost();
  const originalRead = host.Document.prototype.readContentJson;
  const originalFree = host.Document.prototype.free;
  let snapshot: string | undefined;
  let released: { text: string; snapshot: string; canUndo: boolean; canRedo: boolean } | undefined;
  const read = spyOn(host.Document.prototype, 'readCommandContentJson').mockImplementation(function (this: PptxDocument, request: string) {
    const before = content(JSON.parse(originalRead.call(this, request)) as PptxReadResult);
    snapshot = this.snapshotJson();
    return JSON.stringify({ ok: false, version: before.version,
      failure: { code: 'limit-exceeded', message: 'Read response exceeds 64 MiB' } } satisfies PptxReadResult);
  });
  const free = spyOn(host.Document.prototype, 'free').mockImplementation(function (this: PptxDocument) {
    released = { text: content(JSON.parse(originalRead.call(this, '{}')) as PptxReadResult).stories[0]!.text,
      snapshot: this.snapshotJson(), canUndo: this.canUndo(), canRedo: this.canRedo() };
    originalFree.call(this);
  });
  let client: PresentationCommandClient | undefined;
  try {
    const error = await createPresentationCommandClient(await deck(['alpha']), host.options, host.transport)
      .then((value) => { client = value; }, (error: unknown) => error);
    expect(error).toBeInstanceOf(SessionFailure);
    expect(error).toMatchObject({ name: 'SessionFailure', code: 'message', diagnostics: 'limit-exceeded',
      message: 'Presentation command open refused: Read response exceeds 64 MiB' });
    expect(client).toBeUndefined();
    expect(read).toHaveBeenCalledTimes(1);
    expect(free).toHaveBeenCalledTimes(1);
    expect(released).toEqual({ text: 'alpha', snapshot: snapshot!, canUndo: false, canRedo: false });
    expect(host.sentBuffers).toEqual([]);
  } finally {
    try { await client?.dispose(); } finally {
      host.dispose(); read.mockRestore(); free.mockRestore();
    }
  }
});

test('p1_text_reply_contains_ack_selection_state_and_frame', async () => {
  const session = await open();
  try {
    const story = (await session.read()).stories[0]!;
    const ack = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story), 'Hi ')));
    expect(ack).toMatchObject({ ...identity, epoch: session.state.epoch, throughSeq: 1,
      outcomes: [{ seq: 1, status: 'applied', committed: true }],
      selection: { ...selection(story), anchor: 3, focus: 3 },
      editState: { readOnly: false, canUndo: true, canRedo: false,
        toolbar: { target: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId }, canEditText: true } },
      dirtySlideIds: [story.slideId] });
    expect(ack.version).not.toBe(session.state.version);
    const frame = covering(ack, story.slideId);
    expect(frame.slideRevision).toBeGreaterThan(session.state.frame.slideRevision);
    expect(decode(frame).primitives.some((primitive) => primitive.kind === 'textBox' &&
      primitive.storyId === story.storyId && primitive.lines.some((line) =>
        line.caretStops.some((stop) => stop.position === ack.selection!.focus)))).toBe(true);
    expect(decode(frame).primitives.filter((primitive) => primitive.kind === 'textBox')
      .flatMap((box) => box.lines.flatMap((line) => line.runs.map((run) => run.text))).join(''))
      .toBe('Hi alpha');
    expect((await session.read(1)).stories[0]!.text).toBe('Hi alpha');
    const saved = await session.rpc.call.saveCommandEditor(session.fence(1));
    expect(saved).toMatchObject({ ...identity, epoch: ack.epoch, throughSeq: 1, version: ack.version });
    expect(Array.from(new Uint8Array(saved.bytes))).toEqual(Array.from(session.document().saveBytes()));
    expect(session.sentBuffers.every((buffer) => buffer.byteLength === 0)).toBe(true);
    if (!ack.selection || ack.selection.kind === 'object') throw new Error('Expected text selection');
    const noop = await session.rpc.call.command(envelope(ack, 2, insert(ack.selection, '')));
    expect(noop).toMatchObject({ throughSeq: 2, version: ack.version, dirtySlideIds: [],
      outcomes: [{ seq: 2, status: 'noop', committed: false }],
      editState: { canUndo: true, canRedo: false } });
    expect(covering(noop, story.slideId).slideRevision).toBe(frame.slideRevision);
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

test('p1_pipelines_confirmed_basis_afterSeq', async () => {
  const host = await engineHost();
  const posted: PresentationCommandEnvelope[] = [];
  const commandIds = new Set<number>();
  const held: unknown[] = [];
  let hold = false;
  let deliver: ((message: unknown) => void) | undefined;
  let ready!: () => void;
  const repliesHeld = new Promise<void>((resolve) => { ready = resolve; });
  const transport: SessionTransport = { ...host.transport,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'command') {
        posted.push(message.args[0] as PresentationCommandEnvelope);
        commandIds.add(message.id);
      }
      host.transport.post(message, transfer);
    },
    listen(listener) {
      deliver = listener;
      return host.transport.listen((message) => {
        if (hold && isHostMessage(message) && message.kind === 'reply' && commandIds.has(message.id)) {
          held.push(message);
          if (held.length === 2) ready();
        } else listener(message);
      });
    },
  };
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck(), host.options, transport);
    const story = content((await client.read()).result).stories[0]!;
    const selected = await client.dispatch({ kind: 'setSelection', payload: { ...input, selection: selection(story) } });
    const basis = { epoch: selected.epoch, version: selected.version };
    const acknowledgements: PresentationCommandAck[] = [];
    const frames: PresentationCommandFrame[] = [];
    const unsubscribe = client.subscribe((ack) => acknowledgements.push(ack));
    const unsubscribeFrame = client.subscribeFrame((frame) => frames.push(frame));
    hold = true;
    const first = client.dispatch(insert(selection(story), 'A'));
    const second = client.dispatch({ kind: 'insertText', payload: { ...input,
      at: { source: 'afterSeq', seq: selected.throughSeq + 1 }, text: 'B' } });
    await repliesHeld;
    expect(posted.slice(-2).map((command) => command.basis)).toEqual([basis, basis]);
    expect(posted.slice(-2).map((command) => command.seq)).toEqual([selected.throughSeq + 1, selected.throughSeq + 2]);
    expect(posted[posted.length - 1]!.afterSeq).toBe(selected.throughSeq + 1);
    expect(acknowledgements).toEqual([]);
    hold = false;
    for (const reply of held) deliver!(reply);
    const [one, two] = await Promise.all([first, second]);
    expect(one.outcomes).toContainEqual({ seq: selected.throughSeq + 1, status: 'applied', committed: true });
    expect(two.outcomes).toContainEqual({ seq: selected.throughSeq + 2, status: 'applied', committed: true });
    covering(two, story.slideId);
    expect(acknowledgements.map((ack) => ack.throughSeq)).toEqual([one.throughSeq, two.throughSeq]);
    expect(frames.map((frame) => frame.coversThroughSeq)).toEqual([one.throughSeq, two.throughSeq]);
    expect((await client.flush()).throughSeq).toBe(two.throughSeq);
    expect(content((await client.read()).result).stories[0]!.text).toBe('ABalpha');
    expect(await client.save()).toEqual(host.document().saveBytes());
    unsubscribe(); unsubscribeFrame();
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_refusal_preserves_document_and_allows_next_command', async () => {
  const session = await open(['A😀B']);
  try {
    const before = await session.read();
    const story = before.stories[0]!;
    const at = selection(story, 0, 1);
    const saved = session.document().saveBytes();
    const invalid: Array<{ selection: PresentationCommandSelection; basis?: PresentationCommandBasis; reason: string }> = [
      { selection: { ...at, slideId: 'missing-slide' }, reason: 'bad-target' },
      { selection: { ...at, shapeId: 'missing-shape' }, reason: 'bad-target' },
      { selection: { ...at, storyId: 'missing-story' }, reason: 'bad-target' },
      { selection: { ...at, focus: 999 }, reason: 'invalid-offset' },
      { selection: { ...at, anchor: -1 }, reason: 'invalid-offset' },
      { selection: { ...at, focus: 2 }, reason: 'invalid-offset' },
      { selection: at, basis: { epoch: session.state.epoch + 1, version: session.state.version }, reason: 'stale-basis' },
      { selection: at, basis: { epoch: session.state.epoch, version: 'stale' }, reason: 'stale-basis' },
    ];
    let seq = 0;
    for (const bad of invalid) {
      seq += 1;
      const ack = await session.rpc.call.command(envelope(session.state, seq,
        { kind: 'replaceText', payload: { ...input, at: { source: 'confirmed', selection: bad.selection }, text: 'bad' } }, bad.basis));
      expect(ack).toMatchObject({ throughSeq: seq, version: before.version, selection: session.state.selection,
        outcomes: [{ seq, status: 'refused', committed: false, reason: bad.reason }],
        dirtySlideIds: [], editState: { canUndo: false, canRedo: false } });
      expect(await session.read(seq)).toEqual(before);
      expect(session.document().saveBytes()).toEqual(saved);
      expect(session.document().canUndo()).toBe(false);
      expect(session.document().canRedo()).toBe(false);
    }
    const valid = await session.rpc.call.command(envelope(session.state, seq + 1, insert(selection(story), 'ok')));
    expect(valid.outcomes).toEqual([{ seq: seq + 1, status: 'applied', committed: true }]);
    expect((await session.read(seq + 1)).stories[0]!.text).toBe('okA😀B');
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories).toEqual(before.stories);
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

test('p1_field_split_refuses_before_mutation_and_save_keeps_field', async () => {
  const zip = await JSZip.loadAsync(await deck(['tail']));
  const path = 'ppt/slides/slide1.xml';
  const source = await zip.file(path)!.async('string');
  zip.file(path, source.replace('<a:p><a:r>',
    '<a:p><a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:rPr lang="en-US" sz="2400"/><a:t>1</a:t></a:fld><a:r>'));
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  for (const kind of ['paragraphBreak', 'replaceText'] as const) {
    const session = await open(undefined, false, bytes);
    try {
      const before = await session.read();
      const story = before.stories[0]!;
      expect(story.text).toBe('1tail');
      expect(story.paragraphs[0]!.fields).toEqual([{ start: 0, end: 1, fieldType: 'slidenum' }]);
      const saved = session.document().saveBytes();
      const snapshot = session.document().snapshotJson();
      const at = { source: 'confirmed' as const, selection: selection(story, 1, kind === 'replaceText' ? 2 : 1) };
      const intent: PresentationCommandIntent = kind === 'paragraphBreak'
        ? { kind, payload: { ...input, at } } : { kind, payload: { ...input, at, text: 'X\nY' } };
      const ack = await session.rpc.call.command(envelope(session.state, 1, intent));
      expect(ack).toMatchObject({ throughSeq: 1, version: before.version, dirtySlideIds: [],
        outcomes: [{ seq: 1, status: 'refused', reason: 'unsupported', committed: false }],
        editState: { canUndo: false, canRedo: false } });
      expect((await session.read(1)).stories).toEqual(before.stories);
      expect((await session.read(1)).stories[0]!.text).toBe('1tail');
      expect(session.document().snapshotJson()).toBe(snapshot);
      const result = await session.rpc.call.saveCommandEditor(session.fence(1));
      expect(new Uint8Array(result.bytes)).toEqual(new Uint8Array(saved));
      const roundTrip = await JSZip.loadAsync(result.bytes);
      expect(await roundTrip.file(path)!.async('string')).toMatch(/<a:fld\b[^>]*\btype="slidenum"/);
      const reopened = session.Document.openCollaborative(new Uint8Array(result.bytes), 7812);
      try {
        const restored = content(JSON.parse(reopened.readContentJson('{}')) as PptxReadResult).stories[0]!;
        expect(restored.text).toBe('1tail');
        expect(restored.paragraphs[0]!.fields).toEqual(story.paragraphs[0]!.fields);
        expect(restored.paragraphs[0]!.paragraphId).toBe(story.paragraphs[0]!.paragraphId);
      } finally { reopened.free(); }
    } finally { await session.dispose(); }
  }
});

test('removed_field_history_allows_editing_another_paragraph', async () => {
  const zip = await JSZip.loadAsync(await fieldDeck('a', '12', 'z'));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string')).replace('</p:txBody>',
    '<a:p><a:r><a:rPr sz="2400"/><a:t>other</a:t></a:r></a:p></p:txBody>'));
  const session = await open(undefined, false, await zip.generateAsync({ type: 'uint8array' }));
  try {
    const story = (await session.read()).stories[0]!;
    const removed = await session.rpc.call.command(envelope(session.state, 1, { kind: 'replaceText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: '' } }));
    expect(removed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const undone = await session.rpc.call.command(envelope(removed, 2, { kind: 'undo', payload: input }));
    const redone = await session.rpc.call.command(envelope(undone, 3, { kind: 'redo', payload: input }));
    const live = (await session.read(3)).stories[0]!;
    expect(live.text).toBe('az\nother');
    expect(live.paragraphs[0]!.fields).toEqual([]);
    const inserted = await session.rpc.call.command(envelope(redone, 4, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(live, live.paragraphs[1]!.start) }, text: 'Q' } }));
    expect(inserted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect((await session.read(4)).stories[0]!.text).toBe('az\nQother');
    const saved = await JSZip.loadAsync((await session.rpc.call.saveCommandEditor(session.fence(4))).bytes);
    expect(await saved.file(path)!.async('string')).not.toMatch(/<a:fld\b/);
  } finally { await session.dispose(); }
});

for (const scenario of [
  { name: 'p1_consecutive_deletions_refuse_before_surviving_field_is_lost',
    before: 'X12', value: '12', after: '12Y', firstEnd: 3, firstText: '', secondStart: 2, secondEnd: 5,
    secondText: '', expected: '1212Y', fieldStart: 0 },
  { name: 'p1_diff_cell_limit_refuses_before_surviving_field_is_lost',
    before: 'a'.repeat(600), value: '1', after: 'b'.repeat(600), firstEnd: 1, firstText: 'X',
    secondStart: 1200, secondEnd: 1201, secondText: 'Y',
    expected: `X${'a'.repeat(599)}1${'b'.repeat(600)}`, fieldStart: 600 },
]) test(scenario.name, async () => {
  const session = await open(undefined, false, await fieldDeck(scenario.before, scenario.value, scenario.after));
  try {
    const document = session.document();
    const before = await session.read();
    const story = before.stories[0]!;
    expect(story.text).toBe(scenario.before + scenario.value + scenario.after);
    expect(story.paragraphs[0]!.fields).toEqual([{ start: scenario.before.length,
      end: scenario.before.length + scenario.value.length, fieldType: 'slidenum' }]);
    const first = await session.rpc.call.command(envelope(session.state, 1, { kind: 'replaceText',
      payload: { ...input, at: { source: 'confirmed', selection: selection(story, 0, scenario.firstEnd) },
        text: scenario.firstText } }));
    expect(first).toMatchObject({ throughSeq: 1,
      outcomes: [{ seq: 1, status: 'applied', committed: true }],
      editState: { canUndo: true, canRedo: false } });
    const afterFirst = await session.read(1);
    const edited = afterFirst.stories[0]!;
    const fields = [{ start: scenario.fieldStart, end: scenario.fieldStart + scenario.value.length,
      fieldType: 'slidenum' }];
    expect(edited.text).toBe(scenario.expected);
    expect(edited.paragraphs[0]!.fields).toEqual(fields);
    const saved = await session.rpc.call.saveCommandEditor(session.fence(1));
    const undone = await session.rpc.call.command(envelope(first, 2, { kind: 'undo', payload: input }));
    expect(undone.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect((await session.read(2)).stories).toEqual(before.stories);
    const undoZip = await JSZip.loadAsync((await session.rpc.call.saveCommandEditor(session.fence(2))).bytes);
    expect(await undoZip.file('ppt/slides/slide1.xml')!.async('string')).toMatch(/<a:fld\b[^>]*\btype="slidenum"/);
    const redone = await session.rpc.call.command(envelope(undone, 3, { kind: 'redo', payload: input }));
    expect(redone.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect((await session.read(3)).stories).toEqual(afterFirst.stories);
    const redoZip = await JSZip.loadAsync((await session.rpc.call.saveCommandEditor(session.fence(3))).bytes);
    expect(await redoZip.file('ppt/slides/slide1.xml')!.async('string')).toMatch(/<a:fld\b[^>]*\btype="slidenum"/);
    const snapshot = document.snapshotJson();
    const second = await session.rpc.call.command(envelope(redone, 4, { kind: 'replaceText',
      payload: { ...input, at: { source: 'confirmed',
        selection: selection(edited, scenario.secondStart, scenario.secondEnd) }, text: scenario.secondText } }));
    expect(second).toMatchObject({ throughSeq: 4, version: redone.version, dirtySlideIds: [],
      outcomes: [{ seq: 4, status: 'refused', reason: 'unsupported', committed: false }],
      editState: { canUndo: true, canRedo: false } });
    expect(second.frame!.slideRevision).toBe(redone.frame!.slideRevision);
    const after = await session.read(4);
    expect(after.stories).toEqual(afterFirst.stories);
    expect(after.stories[0]!.text).toBe(scenario.expected);
    expect(after.stories[0]!.paragraphs[0]!.fields).toEqual(fields);
    expect(document.snapshotJson()).toBe(snapshot);
    const result = await session.rpc.call.saveCommandEditor(session.fence(4));
    expect(new Uint8Array(result.bytes)).toEqual(new Uint8Array(saved.bytes));
    const zip = await JSZip.loadAsync(result.bytes);
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toMatch(/<a:fld\b[^>]*\btype="slidenum"/);
    const reopened = session.Document.openCollaborative(new Uint8Array(result.bytes), 7812);
    try {
      const restored = content(JSON.parse(reopened.readContentJson('{}')) as PptxReadResult).stories[0]!;
      expect(restored.text).toBe(scenario.expected);
      expect(restored.paragraphs[0]!.fields).toEqual(fields);
    } finally { reopened.free(); }
    document.undoJson();
    expect(content(JSON.parse(document.readContentJson('{}')) as PptxReadResult).stories).toEqual(before.stories);
    expect(document.canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

for (const kind of ['insertText', 'replaceText', 'compositionCommit'] as const) {
  test(`p1_${kind}_without_style_keeps_effective_selection_formatting`, async () => {
    for (const item of [
      { fontSizePt: 42, anchor: 5, focus: 5, text: 'X', expectedText: 'alphaX' },
      { fontSizePt: 24, anchor: 5, focus: 5, text: 'X', expectedText: 'alphaX' },
      { fontSizePt: 42, anchor: 5, focus: 0, text: 'X\nY', expectedText: 'X\nY' },
    ]) {
      const session = await open(undefined, false, await styledDeck(item.fontSizePt));
      try {
        const story = (await session.read()).stories[0]!;
        const at = selection(story, item.anchor, item.focus);
        const expected = effectiveStyleFromSelection(session.handle().story(story.storyId),
          at.anchor, at.focus, DEFAULT_EDITOR_TEXT_STYLE);
        expect(expected).toMatchObject({ bold: false, fontSizePt: item.fontSizePt });
        const payload = { ...input, at: { source: 'confirmed' as const, selection: at },
          compositionId: 'styled-ime' };
        let state: PresentationCommandAck = session.state;
        let seq = 1;
        if (kind === 'compositionCommit') {
          state = await session.rpc.call.command(envelope(state, seq++, { kind: 'compositionStart', payload }));
          expect(state.outcomes).toEqual([{ seq: 1, status: 'applied', committed: false }]);
        }
        const textPayload = { ...payload, text: item.text };
        const intent: PresentationCommandIntent = kind === 'compositionCommit'
          ? { kind: 'compositionCommit', payload: textPayload }
          : kind === 'replaceText' ? { kind: 'replaceText', payload: textPayload }
          : { kind: 'insertText', payload: textPayload };
        const ack = await session.rpc.call.command(envelope(state, seq, intent));
        expect(ack.outcomes).toEqual([{ seq, status: 'applied', committed: true }]);
        expect((await session.read(seq)).stories[0]!.text).toBe(item.expectedText);
        const paragraphs = session.handle().story(story.storyId).paragraphs;
        const runs = paragraphs[paragraphs.length - 1]!.runs;
        expect(runs[runs.length - 1]!.text).toEndWith(item.text.slice(-1));
        expect(runs[runs.length - 1]!.style).toMatchObject(expected);
        expect(ack.editState.toolbar.storedMarks).toEqual({});
      } finally { await session.dispose(); }
    }
  });
}

test('p1_payload_typing_style_expires_on_the_insert_caret_move', async () => {
  const session = await open(undefined, false, await styledDeck());
  try {
    const story = (await session.read()).stories[0]!;
    const marks = { bold: true };
    const marked = await session.rpc.call.command(envelope(session.state, 1, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story) }, text: 'M', style: marks,
    } }));
    expect(marked.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect(marked.editState.toolbar.storedMarks).toEqual({});
    const current = (await session.read(1)).stories[0]!;
    const at = selection(current, current.text.length);
    const expected = effectiveStyleFromSelection(session.handle().story(story.storyId),
      at.anchor, at.focus, DEFAULT_EDITOR_TEXT_STYLE);
    expect(expected.bold).toBe(false);
    const selected = await session.rpc.call.command(envelope(marked, 2, { kind: 'setSelection', payload: {
      ...input, selection: at,
    } }));
    expect(selected.editState.toolbar.storedMarks).toEqual({});
    const ack = await session.rpc.call.command(envelope(selected, 3, insert(at, 'X')));
    expect(ack.outcomes).toEqual([{ seq: 3, status: 'applied', committed: true }]);
    expect((await session.read(3)).stories[0]!.text).toBe('MalphaX');
    const runs = session.handle().story(story.storyId).paragraphs[0]!.runs;
    expect(runs[runs.length - 1]!.text).toEndWith('X');
    expect(runs[runs.length - 1]!.style).toMatchObject(expected);
    expect(ack.editState.toolbar.storedMarks).toEqual({});
  } finally { await session.dispose(); }
});

test('p1_confirmed_basis_surrogate_boundary_is_checked_before_rebase', async () => {
  const session = await open(['A😀B']);
  try {
    const story = (await session.read()).stories[0]!;
    const replaced = await session.rpc.call.command(envelope(session.state, 1, { kind: 'replaceText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: 'QQ',
    } }));
    expect(replaced.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect((await session.read(1)).stories[0]!.text).toBe('AQQB');
    const invalid = await session.rpc.call.command(envelope(session.state, 2, insert(selection(story, 2), 'X')));
    expect(invalid.outcomes).toEqual([{ seq: 2, status: 'refused', reason: 'invalid-offset', committed: false }]);
    expect((await session.read(2)).stories[0]!.text).toBe('AQQB');
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('A😀B');
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

test('p1_composition_commit_uses_rebased_confirmed_caret', async () => {
  const session = await open(['alpha']);
  try {
    const story = (await session.read()).stories[0]!;
    const payload = { ...input, at: { source: 'confirmed' as const, selection: selection(story) }, compositionId: 'ime' };
    const started = await session.rpc.call.command(envelope(session.state, 1, { kind: 'compositionStart', payload }));
    expect(started.outcomes).toEqual([{ seq: 1, status: 'applied', committed: false }]);
    const inserted = await session.rpc.call.command(envelope(session.state, 2, insert(selection(story), 'AA')));
    expect(inserted.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect((await session.read(2)).stories[0]!.text).toBe('AAalpha');
    const committed = await session.rpc.call.command(envelope(session.state, 3,
      { kind: 'compositionCommit', payload: { ...payload, text: 'X' } }));
    expect(committed.outcomes).toEqual([{ seq: 3, status: 'applied', committed: true }]);
    expect((await session.read(3)).stories[0]!.text).toBe('AAXalpha');
  } finally { await session.dispose(); }
});

test('p1_text_preflight_reads_only_the_touched_story', async () => {
  const session = await open(['alpha', 'beta', 'gamma']);
  const read = spyOn(session.handle(), 'readContent');
  const storyRead = spyOn(session.handle(), 'story');
  try {
    const story = content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!;
    const inserted = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story), 'AA')));
    expect(inserted.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    const broken = await session.rpc.call.command(envelope(inserted, 2, { kind: 'paragraphBreak', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 2) },
    } }));
    expect(broken.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    const invalid = await session.rpc.call.command(envelope(broken, 3, insert(selection(story, 999), 'bad')));
    expect(invalid.outcomes).toEqual([{ seq: 3, status: 'refused', reason: 'invalid-offset', committed: false }]);
    expect(read).not.toHaveBeenCalled();
    expect(storyRead.mock.calls).toEqual([[story.storyId], [story.storyId]]);
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult)
      .stories.map((value) => value.text)).toEqual(['AA\nalpha', 'beta', 'gamma']);
  } finally { read.mockRestore(); storyRead.mockRestore(); await session.dispose(); }
});

test('p1_confirmed_ranges_rebase_unlinked_mutations_and_other_story_is_unchanged', async () => {
  const session = await open(['alpha', 'beta']);
  try {
    const before = await session.read();
    const [story, other] = before.stories;
    const first = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story!), 'AA')));
    expect(first.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect((await session.read(1)).stories.map((value) => value.text)).toEqual(['AAalpha', 'beta']);
    const replaced = await session.rpc.call.command(envelope(session.state, 2, { kind: 'replaceText',
      payload: { ...input, at: { source: 'confirmed', selection: selection(story!, 1, 3) }, text: 'X' } }));
    expect(replaced.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect((await session.read(2)).stories.map((value) => value.text)).toEqual(['AAaXha', 'beta']);
    const independent = await session.rpc.call.command(envelope(session.state, 3, { kind: 'replaceText',
      payload: { ...input, at: { source: 'confirmed', selection: selection(other!, 1, 3) }, text: 'X' } }));
    expect(independent.outcomes).toEqual([{ seq: 3, status: 'applied', committed: true }]);
    expect((await session.read(3)).stories.map((value) => value.text)).toEqual(['AAaXha', 'bXa']);
    for (const expected of [['AAaXha', 'beta'], ['AAalpha', 'beta'], ['alpha', 'beta']]) {
      session.document().undoJson();
      expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult)
        .stories.map((value) => value.text)).toEqual(expected);
    }
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

test('p1_public_client_rebases_original_range_without_implicit_afterSeq', async () => {
  const host = await engineHost();
  const posted: PresentationCommandEnvelope[] = [];
  const transport: SessionTransport = { ...host.transport, post(message, transfer) {
    if (isClientMessage(message) && message.kind === 'call' && message.method === 'command') {
      posted.push(message.args[0] as PresentationCommandEnvelope);
    }
    host.transport.post(message, transfer);
  } };
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck(['alpha']), host.options, transport);
    const story = content((await client.read()).result).stories[0]!;
    const basis = { epoch: client.state.epoch, version: client.state.version };
    const first = client.dispatch(insert(selection(story), 'AA'));
    const second = client.dispatch({ kind: 'replaceText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: 'X' } });
    expect(client.state.throughSeq).toBe(0);
    expect(posted.map((command) => command.basis)).toEqual([basis, basis]);
    expect(posted.every((command) => !('afterSeq' in command))).toBe(true);
    const replies = await Promise.all([first, second]);
    expect(replies.map((ack) => ack.outcomes)).toEqual([
      [{ seq: 1, status: 'applied', committed: true }],
      [{ seq: 2, status: 'applied', committed: true }],
    ]);
    expect(content((await client.read()).result).stories[0]!.text).toBe('AAaXha');
    host.document().undoJson();
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('AAalpha');
    host.document().undoJson();
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(host.document().canUndo()).toBe(false);
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_rebase_insertion_ties_preserve_carets_and_range_edges', async () => {
  const cases = [
    { anchor: 1, focus: 1, expected: 'aAAXlpha' },
    { anchor: 0, focus: 1, expected: 'XAAlpha' },
    { anchor: 1, focus: 0, expected: 'XAAlpha' },
    { anchor: 1, focus: 3, expected: 'aAAXha' },
    { anchor: 3, focus: 1, expected: 'aAAXha' },
  ];
  for (const item of cases) {
    const session = await open(['alpha']);
    try {
      const story = (await session.read()).stories[0]!;
      const first = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 1), 'AA')));
      expect(first.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      const second = await session.rpc.call.command(envelope(session.state, 2, { kind: 'replaceText', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, item.anchor, item.focus) }, text: 'X',
      } }));
      expect(second.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
      expect((await session.read(2)).stories[0]!.text).toBe(item.expected);
    } finally { await session.dispose(); }
  }
});

test('p1_rebase_deleted_overlap_refuses_without_text_or_undo_changes', async () => {
  const ranges = [[2, 4], [0, 2], [0, 5], [1, 3], [3, 1]] as const;
  for (const [anchor, focus] of ranges) {
    const session = await open(['alpha']);
    try {
      const story = (await session.read()).stories[0]!;
      const first = await session.rpc.call.command(envelope(session.state, 1, { kind: 'deleteForward', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, 1, 3) }, unit: 'codePoint',
      } }));
      expect(first.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      expect((await session.read(1)).stories[0]!.text).toBe('aha');
      const bytes = session.document().saveBytes();
      const selected = await session.rpc.call.command(envelope(session.state, 2, { kind: 'setSelection', payload: {
        ...input, selection: selection(story, 3, 5),
      } }));
      expect(selected).toMatchObject({ selection: selection(story, 1, 3),
        outcomes: [{ seq: 2, status: 'applied', committed: false }] });
      const second = await session.rpc.call.command(envelope(session.state, 3, { kind: 'replaceText', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, anchor, focus) }, text: 'X',
      } }));
      expect(second).toMatchObject({ version: first.version, dirtySlideIds: [], selection: selected.selection,
        editState: selected.editState,
        outcomes: [{ seq: 3, status: 'refused', reason: 'stale-basis', committed: false }] });
      expect((await session.read(3)).stories[0]!.text).toBe('aha');
      expect(session.document().saveBytes()).toEqual(bytes);
      session.document().undoJson();
      expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
      expect(session.document().canUndo()).toBe(false);
    } finally { await session.dispose(); }
  }
});

test('p1_rebase_replacement_caret_and_paragraph_break_use_story_offsets', async () => {
  const cases = [
    { kind: 'replaceText', anchor: 2, focus: 2, expected: 'aQQXha' },
    { kind: 'replaceText', anchor: 0, focus: 1, expected: 'XQQha' },
    { kind: 'replaceText', anchor: 3, focus: 5, expected: 'aQQX' },
    { kind: 'paragraphBreak', anchor: 1, focus: 3, expected: 'a\nXha' },
  ] as const;
  for (const { kind, anchor, focus, expected } of cases) {
    const session = await open(['alpha']);
    try {
      const story = (await session.read()).stories[0]!;
      const at = { source: 'confirmed' as const, selection: selection(story, 1, kind === 'replaceText' ? 3 : 1) };
      const intent: PresentationCommandIntent = kind === 'replaceText'
        ? { kind, payload: { ...input, at, text: 'QQ' } } : { kind, payload: { ...input, at } };
      const first = await session.rpc.call.command(envelope(session.state, 1, intent));
      expect(first.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      const second = await session.rpc.call.command(envelope(session.state, 2, { kind: 'replaceText', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, anchor, focus) }, text: 'X',
      } }));
      expect(second.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
      expect((await session.read(2)).stories[0]!.text).toBe(expected);
    } finally { await session.dispose(); }
  }
});

test('p1_fixed_basis_mutations_cap_retained_versions_and_text', async () => {
  const session = await open(['A😀B', 'beta']);
  try {
    const [story, other] = (await session.read()).stories;
    if (!story || !other) throw new Error('Missing stories');
    const retained = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story), 'x')));
    expect(retained.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    for (let seq = 2; seq <= 257; seq += 1) {
      const intent: PresentationCommandIntent = seq === 2 ? { kind: 'replaceText', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: 'QQ',
      } } : insert(selection(story), 'x');
      const ack = await session.rpc.call.command(envelope(session.state, seq, intent));
      expect(ack.outcomes).toEqual([{ seq, status: 'applied', committed: true }]);
    }
    const expected = 'x'.repeat(256) + 'AQQB';
    expect((await session.read(257)).stories.map((value) => value.text)).toEqual([expected, 'beta']);
    const expired = await session.rpc.call.command(envelope(session.state, 258, insert(selection(other), 'bad')));
    expect(expired.outcomes).toEqual([{ seq: 258, status: 'refused', reason: 'stale-basis', committed: false }]);
    const usable = await session.rpc.call.command(envelope(retained, 259, insert(selection(story, 1), '')));
    expect(usable).toMatchObject({ selection: selection(story, 256),
      outcomes: [{ seq: 259, status: 'noop', committed: false }] });
    const invalid = await session.rpc.call.command(envelope(retained, 260, insert(selection(story, 3), 'bad')));
    expect(invalid.outcomes).toEqual([{ seq: 260, status: 'refused', reason: 'invalid-offset', committed: false }]);
    const advanced = await session.rpc.call.command(envelope(retained, 261, insert(selection(other), 'B')));
    expect(advanced.outcomes).toEqual([{ seq: 261, status: 'applied', committed: true }]);
    const newlyExpired = await session.rpc.call.command(envelope(retained, 262, insert(selection(other), 'bad')));
    expect(newlyExpired.outcomes).toEqual([{ seq: 262, status: 'refused', reason: 'stale-basis', committed: false }]);
    expect((await session.read(262)).stories.map((value) => value.text)).toEqual([expected, 'Bbeta']);
  } finally { await session.dispose(); }
}, 30_000);

test('p1_fixed_basis_noops_and_refusals_cap_results_and_outcomes', async () => {
  for (const refused of [false, true]) {
    const session = await open(['A😀B']);
    try {
      const story = (await session.read()).stories[0]!;
      const replaced = await session.rpc.call.command(envelope(session.state, 1, { kind: 'replaceText', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: 'QQ',
      } }));
      expect(replaced.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      const at = refused ? { ...selection(story), shapeId: 'missing' } : selection(story);
      const outcome = (seq: number): PresentationCommandOutcome => refused
        ? { seq, status: 'refused', reason: 'bad-target', committed: false }
        : { seq, status: 'noop', committed: false };
      for (let seq = 2; seq <= 258; seq += 1) {
        const ack = await session.rpc.call.command(envelope(session.state, seq, insert(at, '')));
        expect(ack.outcomes).toEqual([outcome(seq)]);
        expect(ack.version).toBe(replaced.version);
      }
      const retained = await session.rpc.call.command(envelope(session.state, 3, insert(at, '')));
      expect(retained.outcomes).toEqual([outcome(3)]);
      const expired = await session.rpc.call.command(envelope(session.state, 2, insert(at, '')));
      expect(expired.outcomes).toEqual([{ seq: 2, status: 'refused', reason: 'stale-basis', committed: false }]);
      const usable = await session.rpc.call.command(envelope(session.state, 259, { kind: 'insertText', payload: {
        ...input, at: { source: 'afterSeq', seq: 3 }, text: '',
      } }, undefined, 3));
      expect(usable.outcomes).toEqual([refused
        ? { seq: 259, status: 'cancelled', reason: 'dependency-refused', committed: false }
        : { seq: 259, status: 'noop', committed: false }]);
      const missing = await session.rpc.call.command(envelope(session.state, 260, { kind: 'insertText', payload: {
        ...input, at: { source: 'afterSeq', seq: 2 }, text: 'bad',
      } }, undefined, 2));
      expect(missing.outcomes).toEqual([{ seq: 260, status: 'refused', reason: 'stale-basis', committed: false }]);
      const invalid = await session.rpc.call.command(envelope(session.state, 261, insert(selection(story, 2), 'bad')));
      expect(invalid.outcomes).toEqual([{ seq: 261, status: 'refused', reason: 'invalid-offset', committed: false }]);
      expect((await session.read(261)).stories[0]!.text).toBe('AQQB');
      const confirmed = await session.rpc.call.command(envelope(session.state, 262, insert(selection(story), 'X')));
      expect(confirmed.outcomes).toEqual([{ seq: 262, status: 'applied', committed: true }]);
      expect((await session.read(262)).stories[0]!.text).toBe('XAQQB');
    } finally { await session.dispose(); }
  }
}, 30_000);

for (const expiry of ['basis', 'result', 'duplicate'] as const) {
  test(`session_expires_old_${expiry}`, async () => {
    const session = await open(['alpha', 'beta']);
    try {
      const [story, other] = (await session.read()).stories;
      if (!story || !other) throw new Error('Missing stories');
      const replacement = (seq: number): PresentationCommandIntent => ({ kind: 'replaceText', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story, 0, 1) }, text: seq % 2 ? 'x' : 'y',
      } });
      const original = envelope(session.state, 1, replacement(1));
      let state: PresentationCommandAck = session.state;
      for (let seq = 1; seq <= 513; seq += 1) {
        state = await session.rpc.call.command(envelope(state, seq, replacement(seq)));
        expect(state.outcomes).toEqual([{ seq, status: 'applied', committed: true }]);
      }
      expect((await session.read(513)).stories[0]!.text).toBe('xlpha');
      const refused = await session.rpc.call.command(expiry === 'basis'
        ? envelope(session.state, 514, insert(selection(other), 'bad'))
        : expiry === 'result' ? envelope(state, 514, { kind: 'insertText', payload: {
          ...input, at: { source: 'afterSeq', seq: 1 }, text: 'bad',
        } }, undefined, 1) : original);
      expect(refused.outcomes).toEqual([{ seq: expiry === 'duplicate' ? 1 : 514,
        status: 'refused', reason: 'stale-basis', committed: false }]);
      expect((await session.read(expiry === 'duplicate' ? 513 : 514)).stories.map((value) => value.text)).toEqual(['xlpha', 'beta']);
    } finally { await session.dispose(); }
  }, 30_000);
}

test('p1_pending_refusal_cancels_only_explicit_result_dependencies', async () => {
  const host = await engineHost();
  const posted: PresentationCommandEnvelope[] = [];
  const transport: SessionTransport = { ...host.transport, post(message, transfer) {
    if (isClientMessage(message) && message.kind === 'call' && message.method === 'command') {
      posted.push(message.args[0] as PresentationCommandEnvelope);
    }
    host.transport.post(message, transfer);
  } };
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck(['alpha', 'beta']), host.options, transport);
    const story = content((await client.read()).result).stories[0]!;
    const started = await client.dispatch({ kind: 'compositionStart', payload: { ...input,
      at: { source: 'confirmed', selection: selection(story) }, compositionId: 'pending-ime' } });
    expect(started.outcomes).toEqual([{ seq: 1, status: 'applied', committed: false }]);
    const replies = await Promise.all([
      client.dispatch(insert(selection(story, 999), 'bad')),
      client.dispatch(insert(selection(story), 'B')),
      client.dispatch({ kind: 'setReadOnly', payload: { ...input, readOnly: true } }),
      client.dispatch({ kind: 'setSelection', payload: { ...input, selection: selection(story) } }),
      client.dispatch({ kind: 'compositionEnd', payload: { ...input, compositionId: 'pending-ime' } }),
      client.dispatch({ kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: 2 }, text: 'dependent' } }),
    ]);
    expect(posted.slice(1).map((command) => command.afterSeq)).toEqual([
      undefined, undefined, undefined, undefined, undefined, 2,
    ]);
    expect(replies.map((ack) => ack.outcomes)).toEqual([
      [{ seq: 2, status: 'refused', reason: 'invalid-offset', committed: false }],
      [{ seq: 3, status: 'applied', committed: true }],
      [{ seq: 4, status: 'applied', committed: false }],
      [{ seq: 5, status: 'applied', committed: false }],
      [{ seq: 6, status: 'applied', committed: false }],
      [{ seq: 7, status: 'cancelled', reason: 'dependency-refused', committed: false }],
    ]);
    expect(client.state.editState.readOnly).toBe(true);
    expect(content((await client.read()).result).stories.map((value) => value.text)).toEqual(['Balpha', 'beta']);
    host.document().undoJson();
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult)
      .stories.map((value) => value.text)).toEqual(['alpha', 'beta']);
    expect(host.document().canUndo()).toBe(false);
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

async function unknownOutcome(call: Promise<unknown>, seq: number, typed = true): Promise<void> {
  const error: unknown = await call.then(() => undefined, (error: unknown) => error);
  expect(error).toBeInstanceOf(Error);
  expect(error).toMatchObject({ name: 'CommandOutcomeUnknown', refusal: { seq, reason: 'unknown-outcome' } });
  if (typed) {
    expect(error).toBeInstanceOf(CommandOutcomeUnknown);
    expect(error).toMatchObject({ seq });
  }
}

test('p1_mid_replacement_throw_rejects_every_later_call_without_acknowledgements', async () => {
  const session = await open(['alpha']);
  let restore = () => {};
  try {
    const story = (await session.read()).stories[0]!;
    const initial = session.state;
    const acknowledgements: PresentationCommandAck[] = [];
    session.rpc.on('ack', (ack) => acknowledgements.push(ack));
    const insertion = spyOn(session.handle(), 'insertText').mockImplementation(() => { throw new Error('insertion trapped'); });
    restore = () => insertion.mockRestore();
    await unknownOutcome(session.rpc.call.command(envelope(initial, 1, { kind: 'replaceText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: 'X' } })), 1, false);
    expect(insertion).toHaveBeenCalledTimes(1);
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('aha');
    await unknownOutcome(session.rpc.call.flushCommandEditor(session.fence(1)), 1, false);
    await unknownOutcome(session.rpc.call.saveCommandEditor(session.fence(1)), 1, false);
    await unknownOutcome(session.rpc.call.readEditor(session.fence(1)), 1, false);
    await unknownOutcome(session.rpc.call.command(envelope(initial, 2, insert(selection(story), 'later'))), 1, false);
    await unknownOutcome(session.rpc.call.command(envelope(initial, 1, insert(selection(story), 'retry'))), 1, false);
    await unknownOutcome(session.rpc.call.thumbnailFrame({ ...identity, slideId: story.slideId }), 1, false);
    expect(acknowledgements).toEqual([]);
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('aha');
  } finally { restore(); await session.dispose(); }
});

test('p1_mid_replacement_throw_notifies_client_failure_once_and_keeps_last_frame', async () => {
  const host = await engineHost();
  let client: PresentationCommandClient | undefined;
  let restore = () => {};
  try {
    client = await createPresentationCommandClient(await deck(['alpha']), host.options, host.transport);
    const story = content((await client.read()).result).stories[0]!;
    const initial = client.state;
    const acknowledgements: PresentationCommandAck[] = [];
    const failures: Error[] = [];
    client.subscribe((ack) => acknowledgements.push(ack));
    client.subscribeFailure((error) => failures.push(error));
    const insertion = spyOn(host.handle(), 'insertText').mockImplementation(() => { throw new Error('insertion trapped'); });
    restore = () => insertion.mockRestore();
    await unknownOutcome(client.dispatch({ kind: 'replaceText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: 'X' } }), 1);
    expect(insertion).toHaveBeenCalledTimes(1);
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('aha');
    await unknownOutcome(client.flush(), 1);
    await unknownOutcome(client.dispatch(insert(selection(story), 'later')), 1);
    await unknownOutcome(client.save(), 1);
    await unknownOutcome(client.read(), 1);
    await unknownOutcome(client.thumbnail(story.slideId), 1);
    expect(acknowledgements).toEqual([]);
    expect(failures).toHaveLength(1);
    expect(failures[0]).toBeInstanceOf(CommandOutcomeUnknown);
    expect(client.state.frame).toBe(initial.frame);
    expect(client.state).toMatchObject({ throughSeq: 0, editState: { readOnly: true, toolbar: { canEditText: false } } });
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('aha');
  } finally { restore(); try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_command_post_throw_is_terminal_without_acknowledgements', async () => {
  const host = await engineHost();
  const posted: PresentationCommandEnvelope[] = [];
  let throws = true;
  const transport: SessionTransport = { ...host.transport, post(message, transfer) {
    if (isClientMessage(message) && message.kind === 'call' && message.method === 'command') {
      posted.push(message.args[0] as PresentationCommandEnvelope);
      if (throws) { throws = false; throw new Error('command post failed'); }
    }
    host.transport.post(message, transfer);
  } };
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck(['alpha']), host.options, transport);
    const story = content((await client.read()).result).stories[0]!;
    const initial = client.state;
    const failures: Error[] = [];
    const acknowledgements: PresentationCommandAck[] = [];
    client.subscribeFailure((error) => failures.push(error));
    client.subscribe((ack) => acknowledgements.push(ack));
    const error: unknown = await client.dispatch(insert(selection(story), 'bad'))
      .then(() => undefined, (error: unknown) => error);
    if (!(error instanceof Error)) throw new Error('Expected command failure');
    expect(error).toMatchObject({ name: 'SessionFailure', code: 'crash', message: 'command post failed' });
    expect(failures).toEqual([error]);
    for (const call of [client.dispatch(insert(selection(story), 'later')), client.flush(),
      client.save(), client.read(), client.thumbnail(story.slideId)]) {
      expect(await call.then(() => undefined, (error: unknown) => error)).toBe(error);
    }
    expect(posted).toHaveLength(1);
    expect(acknowledgements).toEqual([]);
    expect(failures).toHaveLength(1);
    expect(client.state).toMatchObject({ throughSeq: 0, editState: { readOnly: true, toolbar: { canEditText: false } } });
    expect(client.state.frame).toBe(initial.frame);
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(host.document().canUndo()).toBe(false);
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_fence_and_thumbnail_post_throws_are_terminal', async () => {
  for (const method of ['flushCommandEditor', 'saveCommandEditor', 'readEditor', 'thumbnailFrame'] as const) {
    const host = await engineHost();
    let throws = false;
    const transport: SessionTransport = { ...host.transport, post(message, transfer) {
      if (throws && isClientMessage(message) && message.kind === 'call' && message.method === method) {
        throw new Error(`${method} post failed`);
      }
      host.transport.post(message, transfer);
    } };
    let client: PresentationCommandClient | undefined;
    try {
      client = await createPresentationCommandClient(await deck(['alpha']), host.options, transport);
      const story = content((await client.read()).result).stories[0]!;
      const failures: Error[] = [];
      const acknowledgements: PresentationCommandAck[] = [];
      client.subscribeFailure((error) => failures.push(error));
      client.subscribe((ack) => acknowledgements.push(ack));
      throws = true;
      const call = method === 'flushCommandEditor' ? client.flush() : method === 'saveCommandEditor'
        ? client.save() : method === 'readEditor' ? client.read() : client.thumbnail(story.slideId);
      const error: unknown = await call.then(() => undefined, (error: unknown) => error);
      if (!(error instanceof Error)) throw new Error('Expected transport failure');
      expect(error).toMatchObject({ name: 'SessionFailure', code: 'crash', message: `${method} post failed` });
      expect(failures).toEqual([error]);
      for (const later of [client.dispatch(insert(selection(story), 'bad')), client.flush(), client.save(), client.read()]) {
        expect(await later.then(() => undefined, (error: unknown) => error)).toBe(error);
      }
      expect(failures).toHaveLength(1);
      expect(acknowledgements).toEqual([]);
      expect(client.state.editState.readOnly).toBe(true);
      expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
      expect(host.document().canUndo()).toBe(false);
    } finally { try { await client?.dispose(); } finally { host.dispose(); } }
  }
});

test('p1_subscriber_flush_throw_stops_frame_and_ack_notifications', async () => {
  for (const subscriber of ['frame', 'ack'] as const) {
    const host = await engineHost();
    const transport: SessionTransport = { ...host.transport, post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'flushCommandEditor') {
        throw new SessionFailure('crash', 'flush post failed', 'frame subscriber');
      }
      host.transport.post(message, transfer);
    } };
    let client: PresentationCommandClient | undefined;
    try {
      client = await createPresentationCommandClient(await deck(['alpha']), host.options, transport);
      const opened = client;
      const story = content((await opened.read()).result).stories[0]!;
      const notified: string[] = [];
      const failures: PresentationCommandFailure[] = [];
      const secondFailures: PresentationCommandFailure[] = [];
      let rejected: Promise<unknown> | undefined;
      const crash = () => { rejected = opened.flush().catch((error: unknown) => error); };
      opened.subscribeFailure((error) => {
        failures.push(error);
        if ('code' in error) {
          expect(error.code).toBe('crash');
          expect(error.diagnostics).toBe('frame subscriber');
        } else throw new Error('Expected transport failure');
      });
      opened.subscribeFailure((error) => secondFailures.push(error));
      opened.subscribeFrame(() => { notified.push('frame-1'); if (subscriber === 'frame') crash(); });
      opened.subscribeFrame(() => notified.push('frame-2'));
      opened.subscribe(() => { notified.push('ack-1'); if (subscriber === 'ack') crash(); });
      opened.subscribe(() => notified.push('ack-2'));
      const ack = await opened.dispatch(insert(selection(story), 'X'));
      expect(ack.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      expect(notified).toEqual(subscriber === 'frame' ? ['frame-1'] : ['frame-1', 'frame-2', 'ack-1']);
      expect(failures).toHaveLength(1);
      expect(secondFailures).toEqual(failures);
      expect(failures[0]).toMatchObject({ name: 'SessionFailure', code: 'crash', diagnostics: 'frame subscriber' });
      expect(await rejected).toBe(failures[0]);
      expect(opened.state.editState).toMatchObject({ readOnly: true, toolbar: { canEditText: false } });
      expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('Xalpha');
    } finally { try { await client?.dispose(); } finally { host.dispose(); } }
  }
});

test('p1_idle_transport_crash_notifies_failure_once_and_rejects_later_calls', async () => {
  const host = await engineHost();
  let crash: ((error: unknown) => void) | undefined;
  const transport: SessionTransport = { ...host.transport, onError(listener) {
    crash = listener;
    return host.transport.onError(listener);
  } };
  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(await deck(['alpha']), host.options, transport);
    const story = content((await client.read()).result).stories[0]!;
    const initial = client.state;
    const failures: Error[] = [];
    const removed: Error[] = [];
    const acknowledgements: PresentationCommandAck[] = [];
    const unsubscribe = client.subscribeFailure((error) => removed.push(error));
    unsubscribe();
    client.subscribeFailure((error) => failures.push(error));
    client.subscribe((ack) => acknowledgements.push(ack));
    expect(crash).toBeDefined();
    crash!(new Error('worker crashed while idle'));
    crash!(new Error('duplicate crash'));
    expect(failures).toHaveLength(1);
    expect(failures[0]).toBeInstanceOf(Error);
    expect(failures[0]).toMatchObject({ name: 'SessionFailure', code: 'crash', message: 'worker crashed while idle' });
    expect(removed).toEqual([]);
    expect(client.state.frame).toBe(initial.frame);
    expect(client.state.editState).toMatchObject({ readOnly: true, toolbar: { canEditText: false } });
    for (const call of [client.flush(), client.dispatch(insert(selection(story), 'bad')), client.save(), client.read()]) {
      expect(await call.then(() => undefined, (error: unknown) => error)).toBe(failures[0]);
    }
    expect(acknowledgements).toEqual([]);
    expect(failures).toHaveLength(1);
    expect(content(JSON.parse(host.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(host.document().canUndo()).toBe(false);
  } finally { try { await client?.dispose(); } finally { host.dispose(); } }
});

test('p1_duplicate_applies_once_and_wrong_identity_is_ignored', async () => {
  const session = await open();
  try {
    const story = (await session.read()).stories[0]!;
    const command = envelope(session.state, 1, insert(selection(story), 'once'));
    for (const wrong of [{ session: 'other' }, { incarnation: 'old' }, { seq: 2 }]) {
      const ignored = await session.rpc.call.command({ ...command, ...wrong });
      expect(ignored).toMatchObject({ ...identity, throughSeq: 0, version: session.state.version,
        outcomes: [], dirtySlideIds: [] });
      expect((await session.read()).stories[0]!.text).toBe('alpha');
      expect(session.document().canUndo()).toBe(false);
    }
    const first = await session.rpc.call.command(command);
    const duplicate = await session.rpc.call.command(command);
    expect(duplicate).toMatchObject({ throughSeq: 1, version: first.version, outcomes: first.outcomes });
    expect((await session.read(1)).stories[0]!.text).toBe('oncealpha');
    const next = await session.rpc.call.command(envelope(first, 2, { kind: 'insertText', payload: { ...input, at: { source: 'confirmed', selection: selection(story, 4) }, text: '!', undoGroup: 'second' } }));
    expect(next.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect((await session.read(2)).stories[0]!.text).toBe('once!alpha');
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('oncealpha');
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

test('p1_atomic_replacement_break_and_unicode_delete', async () => {
  const cases: Array<{ text: string; anchor: number; focus: number; kind: 'replaceText' | 'paragraphBreak' | 'deleteBackward' | 'deleteForward'; expected: string }> = [
    { text: 'A😀B', anchor: 0, focus: 3, kind: 'replaceText', expected: 'XB' },
    { text: 'A😀B', anchor: 1, focus: 3, kind: 'paragraphBreak', expected: 'A\nB' },
    { text: 'A😀B', anchor: 3, focus: 3, kind: 'deleteBackward', expected: 'AB' },
    { text: 'A😀B', anchor: 1, focus: 1, kind: 'deleteForward', expected: 'AB' },
    { text: 'A😀B', anchor: 3, focus: 1, kind: 'deleteBackward', expected: 'AB' },
  ];
  for (const item of cases) {
    const session = await open([item.text]);
    try {
      const before = await session.read();
      const story = before.stories[0]!;
      const payload = { ...input, at: { source: 'confirmed' as const, selection: selection(story, item.anchor, item.focus) } };
      const intent: PresentationCommandIntent = item.kind === 'replaceText'
        ? { kind: item.kind, payload: { ...payload, text: 'X' } }
        : item.kind === 'paragraphBreak' ? { kind: item.kind, payload }
          : { kind: item.kind, payload: { ...payload, unit: 'codePoint' } };
      const ack = await session.rpc.call.command(envelope(session.state, 1, intent));
      expect(ack.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
      expect(ack.editState).toMatchObject({ canUndo: true, canRedo: false });
      const after = await session.read(1);
      expect(after.stories[0]!.text).toBe(item.expected);
      expect(after.stories[0]!.paragraphs).toHaveLength(item.expected.split('\n').length);
      covering(ack, story.slideId);
      const document = session.document();
      document.undoJson();
      expect(content(JSON.parse(document.readContentJson('{}')) as PptxReadResult).stories).toEqual(before.stories);
      expect(document.canUndo()).toBe(false);
      expect(document.canRedo()).toBe(true);
      document.redoJson();
      expect(content(JSON.parse(document.readContentJson('{}')) as PptxReadResult).stories).toEqual(after.stories);
      expect(document.canUndo()).toBe(true);
      expect(document.canRedo()).toBe(false);
    } finally { await session.dispose(); }
  }

  const session = await open(['alpha\nbeta']);
  try {
    const before = await session.read();
    const story = before.stories[0]!;
    const at = { source: 'confirmed' as const, selection: selection(story, 3, 8) };
    const ack = await session.rpc.call.command(envelope(session.state, 1,
      { kind: 'replaceText', payload: { ...input, at, text: 'X' } }));
    expect(ack.outcomes).toEqual([{ seq: 1, status: 'refused', reason: 'unsupported', committed: false }]);
    expect((await session.read(1)).stories).toEqual(before.stories);
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

function frameTag(frame: PresentationCommandFrame): PresentationCommandFrameTag {
  const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = frame;
  return { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation };
}

test('setSelection_refuses_malformed_supplied_references', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const malformed: unknown[] = [undefined, null, false, true, 0, '', 'confirmed', [], {},
      { source: 'unknown' }, { source: 'confirmed' }, { source: 'confirmed', selection: null },
      { source: 'confirmed', selection: false }, { source: 'confirmed', selection: {} },
      { source: 'afterSeq' }, { source: 'afterSeq', seq: null }, { source: 'afterSeq', seq: '1' },
      { source: 'afterSeq', seq: 0 }, { source: 'afterSeq', seq: 1.5 }];
    for (const [index, at] of malformed.entries()) {
      for (const [offset, target] of [selection(story, 2), null].entries()) {
        const seq = index * 2 + offset + 1;
        const command = { ...envelope(session.state, seq, { kind: 'setSelection', payload: {
          ...input, selection: target,
        } }), payload: { ...input, selection: target, at } } as unknown as PresentationCommandEnvelope;
        const ack = await session.rpc.call.command(command);
        expect(ack.outcomes).toEqual([{ seq, status: 'refused', reason: 'invalid-payload', committed: false }]);
        expect(ack.selection).toEqual(session.state.selection);
        expect(ack.version).toBe(session.state.version);
        expect(ack.editState.canUndo).toBe(false);
      }
    }
    const cleared = await session.rpc.call.command(envelope(session.state, malformed.length * 2 + 1, {
      kind: 'setSelection', payload: { ...input, selection: null },
    }));
    expect(cleared.selection).toBeNull();
    expect(cleared.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    const selected = await session.rpc.call.command(envelope(cleared, malformed.length * 2 + 2, {
      kind: 'setSelection', payload: { ...input, selection: selection(story, 1),
        at: { source: 'confirmed', selection: selection(story) } },
    }));
    expect(selected.selection).toEqual(selection(story, 1));
    expect(selected.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    const chained = await session.rpc.call.command(envelope(selected, malformed.length * 2 + 3, {
      kind: 'setSelection', payload: { ...input, selection: null,
        at: { source: 'afterSeq', seq: selected.throughSeq } },
    }, undefined, selected.throughSeq));
    expect(chained.selection).toBeNull();
    expect(chained.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect((await session.read(malformed.length * 2 + 3)).stories[0]!.text).toBe('abc');
  } finally { await session.dispose(); }
});

function framePoint(frame: PresentationCommandFrame, position: number): { x: number; y: number } {
  const box = decode(frame).primitives.find((primitive) => primitive.kind === 'textBox');
  if (box?.kind !== 'textBox') throw new Error('Expected text geometry');
  const line = box.lines.find((line) => line.caretStops.some((stop) => stop.position === position));
  const stop = line?.caretStops.find((stop) => stop.position === position);
  if (!line || !stop) throw new Error('Expected caret stop');
  return { x: stop.x, y: line.y + line.height / 2 };
}

test('expired_basis_refuses_absolute_selection_after_latest_insert', async () => {
  for (const absolute of [true, false]) {
    const session = await open(['abc']);
    try {
      const story = (await session.read()).stories[0]!;
      let latest: PresentationCommandAck = session.state;
      for (let seq = 1; seq <= 257; seq += 1) {
        latest = await session.rpc.call.command(envelope(session.state, seq, seq === 1
          ? insert(selection(story, 1), 'x')
          : { kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: seq - 1 }, text: 'x' } },
        undefined, seq === 1 ? undefined : seq - 1));
        expect(latest.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      }
      const intent: PresentationCommandIntent = absolute
        ? { kind: 'setSelection', payload: { ...input, selection: selection(story, 1),
          at: { source: 'afterSeq', seq: 257 } } }
        : { kind: 'insertText', payload: { ...input, text: 'bad',
          at: { source: 'confirmed', selection: selection(story, 1) } } };
      const refused = await session.rpc.call.command(envelope(session.state, 258, intent,
        undefined, absolute ? 257 : undefined));
      expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-basis', committed: false });
      expect(refused.selection).toEqual(latest.selection);
      expect((await session.read(258)).stories[0]!.text).toBe(`a${'x'.repeat(257)}bc`);
    } finally { await session.dispose(); }
  }
}, 30_000);

test('compositionCommit_refuses_null_or_missing_payload_and_allows_next_command', async () => {
  for (const payload of [null, undefined]) {
    const session = await open(['abc']);
    try {
      const story = (await session.read()).stories[0]!;
      const started = await session.rpc.call.command(envelope(session.state, 1, {
        kind: 'compositionStart', payload: { ...input, compositionId: 'ime',
          at: { source: 'confirmed', selection: selection(story, 1) } },
      }));
      const command = { ...envelope(started, 2, insert(selection(story, 1), 'bad')),
        kind: 'compositionCommit', payload } as unknown as PresentationCommandEnvelope;
      if (payload === undefined) delete (command as unknown as { payload?: unknown }).payload;
      const refused = await session.rpc.call.command(command);
      expect(refused.throughSeq).toBe(2);
      expect(refused.outcomes).toEqual([{ seq: 2, status: 'refused', reason: 'invalid-payload', committed: false }]);
      expect(refused.selection).toEqual(started.selection);
      expect(refused.version).toBe(started.version);
      expect(refused.editState.canUndo).toBe(false);
      const typed = await session.rpc.call.command(envelope(refused, 3, insert(selection(story, 1), 'Z')));
      expect(typed.outcomes).toEqual([{ seq: 3, status: 'applied', committed: true }]);
      expect((await session.read(3)).stories[0]!.text).toBe('aZbc');
    } finally { await session.dispose(); }
  }
});

test('selection_intents_refuse_malformed_references_before_early_returns', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const list = decode(session.state.frame);
    const box = list.primitives.find((primitive) => primitive.kind === 'textBox');
    if (box?.kind !== 'textBox') throw new Error('Expected text box');
    const blank = { x: list.width - 1, y: list.height - 1 };
    const shape = { x: box.x + box.w / 2, y: box.y + box.h - 1 };
    const geometry = indexHitGeometry(list);
    expect(hitGeometry(geometry, blank)).toBeNull();
    expect(hitGeometry(geometry, shape)?.kind).toBe('shape');
    const malformed = [
      ...[NaN, -1, 1.5, Number.MAX_SAFE_INTEGER, Infinity].flatMap((offset) => [
        { at: { source: 'confirmed', selection: selection(story, offset, 1) }, reason: 'invalid-offset' as const },
        { at: { source: 'confirmed', selection: selection(story, 1, offset) }, reason: 'invalid-offset' as const },
      ]),
      { at: { source: 'wrong', selection: selection(story, 1) }, reason: 'invalid-payload' as const },
      ...[NaN, -1, 1.5, Number.MAX_SAFE_INTEGER, Infinity].map((seq) => ({
        at: { source: 'afterSeq', seq }, reason: 'invalid-payload' as const,
      })),
    ];
    let seq = 0;
    for (const { at, reason } of malformed) {
      const intents: PresentationCommandIntent[] = [
        { kind: 'setSelection', payload: { ...input, selection: null } },
        { kind: 'setSelection', payload: { ...input, selection: selection(story, 2) } },
        { kind: 'selectAtPoint', payload: { ...input, at: { source: 'confirmed', selection: selection(story) },
          frame: frameTag(session.state.frame), ...blank, extend: false } },
        { kind: 'selectAtPoint', payload: { ...input, at: { source: 'confirmed', selection: selection(story) },
          frame: frameTag(session.state.frame), ...shape, extend: false } },
        { kind: 'selectAtPoint', payload: { ...input, at: { source: 'confirmed', selection: selection(story) },
          frame: frameTag(session.state.frame), ...framePoint(session.state.frame, 1), extend: false } },
        { kind: 'moveSelection', payload: { ...input, at: { source: 'confirmed', selection: selection(story) },
          motion: 'left', extend: false } },
      ];
      for (const intent of intents) {
        const command = { ...envelope(session.state, ++seq, intent),
          payload: { ...intent.payload, at } } as unknown as PresentationCommandEnvelope;
        const ack = await session.rpc.call.command(command);
        expect(ack.outcomes).toEqual([{ seq, status: 'refused', reason, committed: false }]);
        expect(ack.selection).toEqual(session.state.selection);
        expect(ack.version).toBe(session.state.version);
        expect(ack.editState.canUndo).toBe(false);
      }
    }
    expect((await session.read(seq)).stories[0]!.text).toBe('abc');
  } finally { await session.dispose(); }
});

test('compact_geometry_preserves_hits_and_motions_without_paint_data', async () => {
  const session = await open(['abc\nxyz']);
  try {
    const story = (await session.read()).stories[0]!;
    const list = decode(session.state.frame);
    const box = list.primitives.find((primitive) => primitive.kind === 'textBox');
    if (box?.kind !== 'textBox') throw new Error('Expected text box');
    const nested: SlideDisplayList = { ...list, primitives: [{ kind: 'table', objectId: 99,
      shapeId: 'table', name: 'table', label: '', x: box.x, y: box.y, w: box.w, h: box.h,
      transform: { rotationDeg: 15, flipH: true },
      primitives: [{ ...box, transform: { flipV: true } }],
    }] };
    for (const source of [list, nested]) {
      const compact = indexHitGeometry(source);
      const full = indexGeometry(source);
      for (let x = box.x - 20; x <= box.x + box.w + 20; x += 20) {
        for (let y = box.y - 20; y <= box.y + box.h + 20; y += 20) {
          expect(hitGeometry(compact, { x, y })).toEqual(hitGeometry(full, { x, y }));
        }
      }
      const motions: PresentationCommandMotion[] = ['left', 'right', 'lineStart', 'lineEnd', 'up', 'down'];
      for (const motion of motions) for (const extend of [false, true]) {
        expect(navigateGeometry(compact, selection(story, 1, 3), motion, extend))
          .toEqual(navigateGeometry(full, selection(story, 1, 3), motion, extend));
      }
      for (const hit of compact.hits) for (const primitive of [hit.primitive, ...hit.transforms, ...hit.clips]) {
        expect(primitive).not.toHaveProperty('paragraphs');
        expect(primitive).not.toHaveProperty('primitives');
        expect(primitive).not.toHaveProperty('fill');
      }
      for (const lines of compact.lines.values()) for (const { line } of lines) {
        expect(line).not.toHaveProperty('runs');
        expect(line.caretStops).not.toBe(box.lines[0].caretStops);
      }
    }
  } finally { await session.dispose(); }
});

test('non_text_frame_hits_preserve_selection_and_chained_typing', async () => {
  for (const blank of [false, true]) {
    const session = await open(['abc']);
    try {
      const story = (await session.read()).stories[0]!;
      const box = decode(session.state.frame).primitives.find((primitive) => primitive.kind === 'textBox');
      if (box?.kind !== 'textBox') throw new Error('Expected text box');
      const point = blank ? { x: -1, y: -1 } : { x: box.x + 1, y: box.y + box.h - 1 };
      const first = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 3), 'X')));
      const clicked = await session.rpc.call.command(envelope(session.state, 2, { kind: 'selectAtPoint', payload: {
        ...input, at: { source: 'afterSeq', seq: 1 }, frame: frameTag(session.state.frame), ...point, extend: false,
      } }, undefined, 1));
      expect(clicked.outcomes).toEqual([{ seq: 2, status: 'noop', committed: false }]);
      expect(clicked.selection).toEqual(first.selection);
      const typed = await session.rpc.call.command(envelope(clicked, 3, { kind: 'insertText', payload: {
        ...input, at: { source: 'afterSeq', seq: 2 }, text: 'Y',
      } }, undefined, 2));
      expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      expect((await session.read(3)).stories[0]!.text).toBe('abcXY');
    } finally { await session.dispose(); }
  }
});

test('afterSeq_navigation_survives_expired_entry_basis', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    for (let seq = 1; seq <= 257; seq += 1) {
      const intent: PresentationCommandIntent = seq === 1 ? insert(selection(story, 1), 'x')
        : { kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: seq - 1 }, text: 'x' } };
      const ack = await session.rpc.call.command(envelope(session.state, seq, intent,
        undefined, seq === 1 ? undefined : seq - 1));
      expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    }
    const expired = await session.rpc.call.command(envelope(session.state, 258, insert(selection(story, 1), 'bad')));
    expect(expired.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-basis', committed: false });
    const moved = await session.rpc.call.command(envelope(session.state, 259, { kind: 'moveSelection', payload: {
      ...input, at: { source: 'afterSeq', seq: 257 }, motion: 'left', extend: false,
    } }, undefined, 257));
    expect(moved.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(moved.selection).toEqual(selection(story, 257));
    const typed = await session.rpc.call.command(envelope(session.state, 260, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 259 }, text: 'Z',
    } }, undefined, 259));
    expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const wrongEpoch = await session.rpc.call.command(envelope(session.state, 261, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 260 }, text: 'bad',
    } }, { epoch: session.state.epoch + 1, version: session.state.version }, 260));
    expect(wrongEpoch.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-basis', committed: false });
    expect((await session.read(261)).stories[0]!.text).toBe(`a${'x'.repeat(256)}Zxbc`);
  } finally { await session.dispose(); }
}, 30_000);

test('p4_moveSelection_resolves_each_motion_after_held_insert', async () => {
  const cases: { motion: PresentationCommandMotion; position: number }[] = [
    { motion: 'left', position: 1 }, { motion: 'right', position: 3 },
    { motion: 'lineStart', position: 0 }, { motion: 'lineEnd', position: 4 },
    { motion: 'up', position: 2 }, { motion: 'down', position: 7 },
  ];
  for (const { motion, position } of cases) {
    const session = await open(['aaa\naaa']);
    try {
      const story = (await session.read()).stories[0]!;
      let release!: () => void;
      let reached!: () => void;
      let delivered = false;
      const gate = new Promise<void>((resolve) => { release = resolve; });
      const held = new Promise<void>((resolve) => { reached = resolve; });
      const inserted = session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 1), 'a')))
        .then(async (ack) => { reached(); await gate; delivered = true; return ack; });
      const moved = session.rpc.call.command(envelope(session.state, 2, { kind: 'moveSelection', payload: {
        ...input, at: { source: 'afterSeq', seq: 1 }, motion, extend: false,
      } }, undefined, 1));
      await held;
      const ack = await moved;
      expect(delivered).toBe(false);
      release();
      const first = await inserted;
      expect(first.selection).toMatchObject({ anchor: 2, focus: 2 });
      expect(ack.outcomes).toEqual([{ seq: 2, status: 'applied', committed: false }]);
      const { goalX, ...at } = ack.selection!;
      expect(at).toEqual({ ...selection(story, position),
        affinity: motion === 'lineEnd' ? 'upstream' : 'downstream' });
      if (motion === 'up' || motion === 'down') expect(goalX).toEqual(expect.any(Number));
      else expect(goalX).toBeUndefined();
      const typed = await session.rpc.call.command(envelope(ack, 3, { kind: 'insertText', payload: {
        ...input, at: { source: 'afterSeq', seq: 2 }, text: 'X',
      } }, undefined, 2));
      expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      const text = 'aaaa\naaa';
      expect((await session.read(3)).stories[0]!.text).toBe(text.slice(0, position) + 'X' + text.slice(position));
    } finally { await session.dispose(); }
  }
});

test('p4_moveSelection_extends_and_collapses_authoritative_ranges', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const moved = await session.rpc.call.command(envelope(session.state, 1, { kind: 'moveSelection', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 1, 3) }, motion: 'left', extend: true,
    } }));
    expect(moved.selection).toEqual(selection(story, 1, 2));
    const collapsed = await session.rpc.call.command(envelope(moved, 2, { kind: 'moveSelection', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, motion: 'right', extend: false,
    } }, undefined, 1));
    expect(collapsed.selection).toEqual(selection(story, 2));
  } finally { await session.dispose(); }
});

test('p4_moveSelection_uses_geometry_after_paragraph_split', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const split = session.rpc.call.command(envelope(session.state, 1, { kind: 'paragraphBreak', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 1) },
    } }));
    const moved = session.rpc.call.command(envelope(session.state, 2, { kind: 'moveSelection', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, motion: 'lineEnd', extend: false,
    } }, undefined, 1));
    expect((await split).outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const ack = await moved;
    expect(ack.selection).toEqual({ ...selection(story, 4), affinity: 'upstream' });
    expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    const typed = await session.rpc.call.command(envelope(ack, 3, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 2 }, text: 'Y',
    } }, undefined, 2));
    expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect((await session.read(3)).stories[0]!.text).toBe('a\nbcY');
    expect((await session.read(3)).stories[0]!.paragraphs).toHaveLength(2);
  } finally { await session.dispose(); }
});

test('p7_fresh_paragraph_edge_noop_acknowledges_caret_for_chained_end_and_typing', async () => {
  for (const kind of ['deleteBackward', 'deleteForward'] as const) {
    const session = await open(['a\nbc']);
    try {
      const story = (await session.read()).stories[0]!;
      expect(session.state.selection).toBeNull();
      const at = selection(story, kind === 'deleteBackward' ? 2 : 1);
      const deleted = await session.rpc.call.command(envelope(session.state, 1, {
        kind, payload: { ...input, at: { source: 'confirmed', selection: at }, unit: 'codePoint' },
      }));
      expect(deleted.outcomes).toEqual([{ seq: 1, status: 'noop', committed: false }] as const);
      expect(deleted.selection).toEqual(at);
      expect(deleted.version).toBe(session.state.version);
      expect(deleted.dirtySlideIds).toEqual([]);
      expect(deleted.editState.canUndo).toBe(session.state.editState.canUndo);
      expect(deleted.editState.canRedo).toBe(session.state.editState.canRedo);
      expect((await session.read(1)).stories[0]!.text).toBe('a\nbc');
      const moved = await session.rpc.call.command(envelope(deleted, 2, {
        kind: 'moveSelection', payload: { ...input, at: { source: 'afterSeq', seq: 1 },
          motion: 'lineEnd', extend: false },
      }, undefined, 1));
      expect(moved.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
      expect(moved.selection).toEqual({ ...selection(story, kind === 'deleteBackward' ? 4 : 1), affinity: 'upstream' });
      const typed = await session.rpc.call.command(envelope(moved, 3, {
        kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: 2 }, text: 'Y' },
      }, undefined, 2));
      expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      const final = (await session.read(3)).stories[0]!;
      expect(final.text).toBe(kind === 'deleteBackward' ? 'a\nbcY' : 'aY\nbc');
      expect(final.paragraphs.map((paragraph) => final.text.slice(paragraph.start, paragraph.end)))
        .toEqual(kind === 'deleteBackward' ? ['a', 'bcY'] : ['aY', 'bc']);
    } finally { await session.dispose(); }
  }
});

test('collapsed_paragraph_boundary_deletes_are_noops_on_confirmed_and_afterSeq_paths', async () => {
  for (const kind of ['deleteBackward', 'deleteForward'] as const) {
    for (const relative of [false, true]) {
      const session = await open(['a\nbc']);
      try {
        const story = (await session.read()).stories[0]!;
        const at = selection(story, kind === 'deleteBackward' ? 2 : 1);
        const selected = await session.rpc.call.command(envelope(session.state, 1, {
          kind: 'setSelection', payload: { ...input, selection: at },
        }));
        const deleted = await session.rpc.call.command(envelope(selected, 2, { kind, payload: {
          ...input, at: relative ? { source: 'afterSeq', seq: 1 } : { source: 'confirmed', selection: at },
          unit: 'codePoint', count: 3,
        } }, undefined, relative ? 1 : undefined));
        expect(deleted.outcomes).toEqual([{ seq: 2, status: 'noop', committed: false }]);
        expect(deleted.selection).toEqual(selected.selection);
        expect(deleted.version).toBe(selected.version);
        expect(deleted.dirtySlideIds).toEqual([]);
        expect(deleted.editState.canUndo).toBe(selected.editState.canUndo);
        expect(deleted.editState.canRedo).toBe(selected.editState.canRedo);
        expect((await session.read(2)).stories[0]!.text).toBe('a\nbc');
        const typed = await session.rpc.call.command(envelope(deleted, 3, { kind: 'insertText', payload: {
          ...input, at: { source: 'afterSeq', seq: 2 }, text: 'Y',
        } }, undefined, 2));
        expect(typed.outcomes).toEqual([{ seq: 3, status: 'applied', committed: true }]);
        expect((await session.read(3)).stories[0]!.text).toBe(kind === 'deleteBackward' ? 'a\nYbc' : 'aY\nbc');
      } finally { await session.dispose(); }
    }
  }
});

test('p4_selectAtPoint_hits_emitted_geometry_and_maps_later_insert_or_covering_deletion', async () => {
  for (const deletion of [false, true]) {
    const session = await open(['abc']);
    try {
      const story = (await session.read()).stories[0]!;
      const point = framePoint(session.state.frame, 2);
      const first = await session.rpc.call.command(envelope(session.state, 1, deletion
        ? { kind: 'replaceText', payload: { ...input,
          at: { source: 'confirmed', selection: selection(story, 1, 3) }, text: '' } }
        : insert(selection(story, 1), 'MMMM')));
      const clicked = await session.rpc.call.command(envelope(session.state, 2, { kind: 'selectAtPoint', payload: {
        ...input, at: { source: 'afterSeq', seq: 1 }, frame: frameTag(session.state.frame),
        ...point, extend: false,
      } }, undefined, 1));
      expect(first.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      expect(clicked.outcomes).toEqual([{ seq: 2, status: 'applied', committed: false }]);
      expect(clicked.selection).toEqual(selection(story, deletion ? 1 : 6));
      const typed = await session.rpc.call.command(envelope(clicked, 3, { kind: 'insertText', payload: {
        ...input, at: { source: 'afterSeq', seq: 2 }, text: 'Y',
      } }, undefined, 2));
      expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      expect((await session.read(3)).stories[0]!.text).toBe(deletion ? 'aY' : 'aMMMMbYc');
    } finally { await session.dispose(); }
  }
});

test('p4_selectAtPoint_extend_keeps_reference_anchor', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const ack = await session.rpc.call.command(envelope(session.state, 1, { kind: 'selectAtPoint', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 1) },
      frame: frameTag(session.state.frame), ...framePoint(session.state.frame, 3), extend: true,
    } }));
    expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(ack.selection).toEqual({ ...selection(story, 1, 3), affinity: 'upstream' });
    const inserted = await session.rpc.call.command(envelope(session.state, 2, insert(selection(story, 1), 'X')));
    expect(inserted.selection).toEqual(selection(story, 2));
    const extended = await session.rpc.call.command(envelope(session.state, 3, { kind: 'selectAtPoint', payload: {
      ...input, at: { source: 'afterSeq', seq: 2 }, frame: frameTag(session.state.frame),
      ...framePoint(session.state.frame, 3), extend: true,
    } }, undefined, 2));
    expect(extended.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(extended.selection).toEqual({ ...selection(story, 2, 4), affinity: 'upstream' });
  } finally { await session.dispose(); }
});

test('pending_shift_click_without_compatible_predecessor_selects_normally', async () => {
  for (const otherStory of [false, true]) {
    const zip = await JSZip.loadAsync(await deck(['abc']));
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
    zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Other"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="3000000"/><a:ext cx="6000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>other</a:t></a:r></a:p></p:txBody></p:sp></p:spTree>`));
    const session = await open(undefined, false, await zip.generateAsync({ type: 'uint8array' }));
    try {
      const stories = (await session.read()).stories;
      expect(stories).toHaveLength(2);
      const cleared = await session.rpc.call.command(envelope(session.state, 1, { kind: 'setSelection', payload: {
        ...input, selection: otherStory ? selection(stories[1]!, 1) : null,
      } }));
      const clicked = await session.rpc.call.command(envelope(cleared, 2, { kind: 'selectAtPoint', payload: {
        ...input, at: { source: 'afterSeq', seq: 1 }, frame: frameTag(cleared.frame!),
        ...framePoint(cleared.frame!, 2), extend: true,
      } }, undefined, 1));
      expect(clicked.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
      expect(clicked.selection).toEqual(selection(stories[0]!, 2));
      const typed = await session.rpc.call.command(envelope(clicked, 3, { kind: 'insertText', payload: {
        ...input, at: { source: 'afterSeq', seq: 2 }, text: 'X',
      } }, undefined, 2));
      expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      expect((await session.read(3)).stories[0]!.text).toBe('abXc');
    } finally { await session.dispose(); }
  }
});

test('selection_frames_reuse_layout_until_slide_revision_changes', async () => {
  const original = wasm.presentationDisplayListJson;
  const layout = spyOn(wasm, 'presentationDisplayListJson').mockImplementation(original);
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const initialCalls = layout.mock.calls.length;
    const selected = await session.rpc.call.command(envelope(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(story, 1),
    } }));
    expect(layout).toHaveBeenCalledTimes(initialCalls);
    expect(selected.frame).toMatchObject({ coversThroughSeq: 1, slideRevision: 0 });
    const edited = await session.rpc.call.command(envelope(selected, 2, insert(selection(story, 1), 'X')));
    expect(layout).toHaveBeenCalledTimes(initialCalls + 1);
    expect(edited.frame).toMatchObject({ coversThroughSeq: 2, slideRevision: 1 });
    expect((decode(edited.frame!).primitives.find((primitive) => primitive.kind === 'textBox') as { paragraphs: { runs: { text: string }[] }[] })
      .paragraphs[0]!.runs.map((run) => run.text).join('')).toBe('aXbc');
    const undone = await session.rpc.call.command(envelope(edited, 3, { kind: 'undo', payload: input }));
    expect(layout).toHaveBeenCalledTimes(initialCalls + 2);
    expect((await session.read(3)).stories[0]!.text).toBe('abc');
    expect(undone.frame?.slideRevision).toBeGreaterThan(edited.frame!.slideRevision);
  } finally { await session.dispose(); layout.mockRestore(); }
});

test('p4_selectAtPoint_refuses_unknown_and_foreign_tags', async () => {
  const session = await open(['abc', 'other']);
  try {
    const story = (await session.read()).stories[0]!;
    const tag = frameTag(session.state.frame);
    const tags = [{ ...tag, slideRevision: 99 }, { ...tag, coversThroughSeq: 99 },
      { ...tag, slideId: session.state.slides[1]!.id }, { ...tag, documentEpoch: 2 },
      { ...tag, incarnation: 'other-worker' }];
    for (const [index, frame] of tags.entries()) {
      const ack = await session.rpc.call.command(envelope(session.state, index + 1, { kind: 'selectAtPoint', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story) },
        frame, ...framePoint(session.state.frame, 1), extend: false,
      } }));
      expect(ack.outcomes).toEqual([{ seq: index + 1, status: 'refused', reason: 'stale-basis', committed: false }]);
    }
    expect((await session.read(tags.length)).stories[0]!.text).toBe('abc');
  } finally { await session.dispose(); }
});

test('p4_emitted_frame_retention_evicts_old_tags_and_keeps_latest', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    let latest: PresentationCommandAck = session.state;
    for (let seq = 1; seq <= 258; seq += 1) {
      latest = await session.rpc.call.command(envelope(session.state, seq, { kind: 'setSelection', payload: {
        ...input, selection: selection(story, 1),
      } }));
    }
    const expired = await session.rpc.call.command(envelope(session.state, 259, { kind: 'selectAtPoint', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story) },
      frame: frameTag(session.state.frame), ...framePoint(session.state.frame, 1), extend: false,
    } }));
    expect(expired.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(expired.selection).toEqual(selection(story, 1));
    const fresh = await session.rpc.call.command(envelope(session.state, 260, { kind: 'selectAtPoint', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story) },
      frame: frameTag(latest.frame!), ...framePoint(latest.frame!, 2), extend: false,
    } }));
    expect(fresh.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(fresh.selection).toEqual(selection(story, 2));
    expect((await session.read(260)).stories[0]!.text).toBe('abc');
  } finally { await session.dispose(); }
});

test('p4_newer_command_basis_preserves_queued_click_frames', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const inserted = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 1), 'X')));
    const ack = await session.rpc.call.command(envelope(inserted, 2, { kind: 'selectAtPoint', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, frame: frameTag(session.state.frame),
      ...framePoint(session.state.frame, 1), extend: false,
    } }, undefined, 1));
    expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(ack.selection).toEqual(selection(story, 2));
    const latest = await session.rpc.call.command(envelope(inserted, 3, { kind: 'selectAtPoint', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story, 2) }, frame: frameTag(ack.frame!),
      ...framePoint(ack.frame!, 2), extend: false,
    } }));
    expect(latest.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    const typed = await session.rpc.call.command(envelope(latest, 4, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 3 }, text: 'Y',
    } }, undefined, 3));
    expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect((await session.read(4)).stories[0]!.text).toBe('aXYbc');
  } finally { await session.dispose(); }
});

test('p4_afterSeq_accepts_successful_predecessor_already_covered_by_basis', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const inserted = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 1), 'X')));
    const ack = await session.rpc.call.command(envelope(inserted, 2, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, text: 'Y',
    } }, undefined, 1));
    expect(ack.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect((await session.read(2)).stories[0]!.text).toBe('aXYbc');
  } finally { await session.dispose(); }
});

test('p7_afterSeq_null_selection_survives_expired_entry_basis', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    for (let seq = 1; seq <= 257; seq += 1) {
      const intent: PresentationCommandIntent = seq === 1 ? insert(selection(story, 1), 'x')
        : { kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: seq - 1 }, text: 'x' } };
      const ack = await session.rpc.call.command(envelope(session.state, seq, intent,
        undefined, seq === 1 ? undefined : seq - 1));
      expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    }
    const cleared = await session.rpc.call.command(envelope(session.state, 258, {
      kind: 'setSelection', payload: { ...input, selection: null, at: { source: 'afterSeq', seq: 257 } },
    }, undefined, 257));
    expect(cleared.outcomes).toEqual([{ seq: 258, status: 'applied', committed: false }] as const);
    expect(cleared.selection).toBeNull();
    expect(cleared.editState.toolbar.target).toBeNull();
    expect((await session.read(258)).stories[0]!.text).toBe(`a${'x'.repeat(257)}bc`);
  } finally { await session.dispose(); }
}, 30_000);

test('p4_setSelection_null_remains_admissible_with_a_result_dependency', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const inserted = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 1), 'X')));
    const cleared = await session.rpc.call.command(envelope(inserted, 2, { kind: 'setSelection', payload: {
      ...input, selection: null, at: { source: 'afterSeq', seq: 1 },
    } }, undefined, 1));
    expect(cleared.outcomes).toEqual([{ seq: 2, status: 'applied', committed: false }]);
    expect(cleared.selection).toBeNull();
    const repeated = await session.rpc.call.command(envelope(cleared, 3, { kind: 'setSelection', payload: {
      ...input, selection: null, at: { source: 'afterSeq', seq: 2 },
    } }, undefined, 2));
    expect(repeated.outcomes).toEqual([{ seq: 3, status: 'applied', committed: false }]);
    expect(repeated.selection).toBeNull();
    expect((await session.read(3)).stories[0]!.text).toBe('aXbc');
  } finally { await session.dispose(); }
});

test('p4_covered_refused_predecessor_still_cancels_dependents', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const refused = await session.rpc.call.command(envelope(session.state, 1, insert(selection(story, 99), 'bad')));
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'invalid-offset' });
    const inserted = await session.rpc.call.command(envelope(session.state, 2, insert(selection(story, 1), 'X')));
    const cancelled = await session.rpc.call.command(envelope(inserted, 3, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 1 }, text: 'Y',
    } }, undefined, 1));
    expect(cancelled.outcomes).toEqual([{ seq: 3, status: 'cancelled', reason: 'dependency-refused', committed: false }]);
    expect((await session.read(3)).stories[0]!.text).toBe('aXbc');
  } finally { await session.dispose(); }
});

test('p1_unchanged_slides_keep_revision_and_dirty_ids_are_exact', async () => {
  const session = await open();
  try {
    const before = await session.read();
    const [active, other, untouched] = session.state.slides;
    const thumbnails = await Promise.all([other!, untouched!].map((slide) =>
      session.rpc.call.thumbnailFrame({ ...identity, slideId: slide.id })));
    expect(thumbnails.map((thumbnail) => thumbnail.status)).toEqual(['frame', 'frame']);
    const activeStory = before.stories.find((story) => story.slideId === active!.id)!;
    const one = await session.rpc.call.command(envelope(session.state, 1, insert(selection(activeStory), 'A')));
    expect(one.dirtySlideIds).toEqual([active!.id]);
    const activeFrame = covering(one, active!.id);
    for (const thumbnail of thumbnails) {
      expect(await session.rpc.call.thumbnailFrame({ ...identity, slideId: thumbnail.slideId,
        knownRevision: thumbnail.slideRevision })).toMatchObject({ status: 'unchanged', slideRevision: thumbnail.slideRevision });
    }
    const otherStory = before.stories.find((story) => story.slideId === other!.id)!;
    const selected = await session.rpc.call.command(envelope(one, 2,
      { kind: 'setSelection', payload: { ...input, selection: selection(otherStory) } }));
    expect(selected).toMatchObject({ dirtySlideIds: [], version: one.version,
      outcomes: [{ seq: 2, status: 'applied', committed: false }] });
    expect(covering(selected, other!.id).slideRevision).toBe(thumbnails[0]!.slideRevision);
    const two = await session.rpc.call.command(envelope(selected, 3, insert(selection(otherStory), 'B')));
    expect(two.dirtySlideIds).toEqual([other!.id]);
    expect(covering(two, other!.id).slideRevision).toBeGreaterThan(thumbnails[0]!.slideRevision);
    expect(await session.rpc.call.thumbnailFrame({ ...identity, slideId: active!.id,
      knownRevision: activeFrame.slideRevision })).toMatchObject({ status: 'unchanged', slideRevision: activeFrame.slideRevision });
    const changed = await session.rpc.call.thumbnailFrame({ ...identity, slideId: other!.id,
      knownRevision: thumbnails[0]!.slideRevision });
    expect(changed.status).toBe('frame');
    expect(changed.slideRevision).toBeGreaterThan(thumbnails[0]!.slideRevision);
    expect(await session.rpc.call.thumbnailFrame({ ...identity, slideId: untouched!.id,
      knownRevision: thumbnails[1]!.slideRevision })).toMatchObject({ status: 'unchanged', slideRevision: thumbnails[1]!.slideRevision });
    const returned = await session.rpc.call.command(envelope(two, 4,
      { kind: 'setSelection', payload: { ...input, selection: selection(activeStory, 1) } }));
    expect(returned).toMatchObject({ dirtySlideIds: [], version: two.version,
      outcomes: [{ seq: 4, status: 'applied', committed: false }] });
    expect(covering(returned, active!.id).slideRevision).toBe(activeFrame.slideRevision);
  } finally { await session.dispose(); }
});

test('p1_readOnly_command_refuses_later_writes', async () => {
  const session = await open();
  try {
    const before = await session.read();
    const story = before.stories[0]!;
    const locked = await session.rpc.call.command(envelope(session.state, 1,
      { kind: 'setReadOnly', payload: { ...input, readOnly: true } }));
    expect(locked).toMatchObject({ throughSeq: 1, version: before.version, dirtySlideIds: [],
      editState: { readOnly: true, canUndo: false }, outcomes: [{ seq: 1, status: 'applied', committed: false }] });
    const write = await session.rpc.call.command(envelope(locked, 2, insert(selection(story), 'bad')));
    expect(write).toMatchObject({ throughSeq: 2, version: before.version, dirtySlideIds: [],
      outcomes: [{ seq: 2, status: 'refused', reason: 'read-only', committed: false }] });
    expect(await session.read(2)).toEqual(before);
    expect(session.document().canUndo()).toBe(false);
    const unlocked = await session.rpc.call.command(envelope(write, 3,
      { kind: 'setReadOnly', payload: { ...input, readOnly: false } }));
    const valid = await session.rpc.call.command(envelope(unlocked, 4, insert(selection(story), 'ok')));
    expect(valid.outcomes).toEqual([{ seq: 4, status: 'applied', committed: true }]);
    expect((await session.read(4)).stories[0]!.text).toBe('okalpha');
  } finally { await session.dispose(); }
  const initiallyLocked = await open(['alpha'], true);
  try {
    expect(initiallyLocked.state.editState.readOnly).toBe(true);
    const story = (await initiallyLocked.read()).stories[0]!;
    expect((await initiallyLocked.rpc.call.command(envelope(initiallyLocked.state, 1, insert(selection(story), 'bad')))).outcomes)
      .toEqual([{ seq: 1, status: 'refused', reason: 'read-only', committed: false }]);
    expect(initiallyLocked.document().canUndo()).toBe(false);
  } finally { await initiallyLocked.dispose(); }
});

test('p1_composition_commit_is_atomic_and_dependents_cancel_after_refusal', async () => {
  const session = await open(['alpha']);
  try {
    const story = (await session.read()).stories[0]!;
    const payload = { ...input, at: { source: 'confirmed' as const, selection: selection(story, 0, 5) }, compositionId: 'ime-1' };
    const started = await session.rpc.call.command(envelope(session.state, 1, { kind: 'compositionStart', payload }));
    expect(started).toMatchObject({ version: session.state.version, dirtySlideIds: [], editState: { canUndo: false } });
    const committed = await session.rpc.call.command(envelope(started, 2,
      { kind: 'compositionCommit', payload: { ...payload, text: '日本語' } }));
    expect(committed.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect((await session.read(2)).stories[0]!.text).toBe('日本語');
    const ended = await session.rpc.call.command(envelope(committed, 3,
      { kind: 'compositionEnd', payload: { ...input, compositionId: 'ime-1' } }));
    expect(ended.version).toBe(committed.version);
    expect(ended.dirtySlideIds).toEqual([]);
    const bad = envelope(ended, 4, insert({ ...selection(story), shapeId: 'missing' }, 'bad'));
    const dependent = envelope(ended, 5, { kind: 'insertText', payload: { ...input,
      at: { source: 'afterSeq', seq: 4 }, text: 'dependent' } }, undefined, 4);
    const [refused, cancelled] = await Promise.all([
      session.rpc.call.command(bad), session.rpc.call.command(dependent),
    ]);
    expect(refused.outcomes).toEqual([{ seq: 4, status: 'refused', reason: 'bad-target', committed: false }]);
    expect(cancelled).toMatchObject({ throughSeq: 5, version: ended.version, dirtySlideIds: [],
      outcomes: [{ seq: 5, status: 'cancelled', reason: 'dependency-refused', committed: false }] });
    expect((await session.read(5)).stories[0]!.text).toBe('日本語');
    const independent = await session.rpc.call.command(envelope(ended, 6, insert(selection(story), '!')));
    expect(independent.outcomes).toEqual([{ seq: 6, status: 'applied', committed: true }]);
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('日本語');
    session.document().undoJson();
    expect(content(JSON.parse(session.document().readContentJson('{}')) as PptxReadResult).stories[0]!.text).toBe('alpha');
    expect(session.document().canUndo()).toBe(false);
  } finally { await session.dispose(); }
});

test('p8_evicted_click_uses_newest_frame_and_preserves_following_insert', async () => {
  for (const mutationCount of [0, 513]) {
    const session = await open(['abc', 'other']);
    try {
      const stories = (await session.read()).stories;
      const story = stories[0]!;
      const expired = frameTag(session.state.frame);
      const count = mutationCount || 258;
      let latest: PresentationCommandAck = session.state;
      for (let seq = 1; seq <= count; seq += 1) {
        const intent: PresentationCommandIntent = mutationCount === 0
          ? { kind: 'setSelection', payload: { ...input, selection: selection(story, 1) } }
          : seq === 1 ? insert(selection(story, 1), 'x')
            : { kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: seq - 1 }, text: 'x' } };
        latest = await session.rpc.call.command(envelope(latest, seq, intent,
          undefined, mutationCount > 0 && seq > 1 ? seq - 1 : undefined));
        expect(latest.outcomes[0]).toMatchObject({ status: 'applied' });
      }
      const clicked = await session.rpc.call.command(envelope(latest, count + 1, {
        kind: 'selectAtPoint', payload: { ...input, at: { source: 'afterSeq', seq: count },
          frame: expired, ...framePoint(mutationCount === 0 ? session.state.frame : latest.frame!, 2), extend: false },
      }, undefined, count));
      expect(clicked.outcomes).toEqual([{ seq: count + 1, status: 'applied', committed: false }] as const);
      expect(clicked.selection).toEqual(selection(story, 2));
      const typed = await session.rpc.call.command(envelope(clicked, count + 2, {
        kind: 'insertText', payload: { ...input, at: { source: 'afterSeq', seq: count + 1 }, text: 'Z' },
      }, undefined, count + 1));
      expect(typed.outcomes).toEqual([{ seq: count + 2, status: 'applied', committed: true }] as const);
      expect((await session.read(count + 2)).stories[0]!.text)
        .toBe(mutationCount === 0 ? 'abZc' : `axZ${'x'.repeat(512)}bc`);
    } finally { await session.dispose(); }
  }
}, 30_000);

test('p8_malformed_cloneable_frame_refuses_without_throwing_and_next_command_applies', async () => {
  const session = await open(['abc']);
  try {
    const story = (await session.read()).stories[0]!;
    const tag = frameTag(session.state.frame);
    const malformed: unknown[] = [{ ...tag, coversThroughSeq: 1n },
      { ...tag, slideRevision: Number.MAX_SAFE_INTEGER + 1 }, { ...tag, coversThroughSeq: -1 },
      { ...tag, documentEpoch: 1.5 }, { ...tag, slideId: 1n }, { ...tag, incarnation: {} }, null];
    let latest: PresentationCommandAck = session.state;
    for (const [index, frame] of malformed.entries()) {
      const seq = index * 2 + 1;
      const command = envelope(latest, seq, { kind: 'selectAtPoint', payload: {
        ...input, at: { source: 'confirmed', selection: selection(story) },
        frame: frame as PresentationCommandFrameTag, ...framePoint(session.state.frame, 1), extend: false,
      } });
      const refused = await session.rpc.call.command(command);
      expect(refused.outcomes).toEqual([{ seq, status: 'refused', reason: 'invalid-payload', committed: false }] as const);
      expect(refused.throughSeq).toBe(seq);
      expect(refused.version).toBe(latest.version);
      latest = await session.rpc.call.command(envelope(refused, seq + 1, insert(selection(story, 1), 'Z')));
      expect(latest.outcomes).toEqual([{ seq: seq + 1, status: 'applied', committed: true }] as const);
    }
    expect((await session.read(latest.throughSeq)).stories[0]!.text).toBe(`a${'Z'.repeat(malformed.length)}bc`);
  } finally { await session.dispose(); }
});
