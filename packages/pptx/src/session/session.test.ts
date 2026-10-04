import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { runInNewContext } from 'node:vm';
import {
  createSessionClient, createSessionHost, isHostMessage, SESSION_SUPERSEDED, SessionFailure,
  type MethodPolicy, type SessionTransport,
} from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isTiff } from '../../../../shared/media';
import type { PptxEditRequest, PptxReadResult } from '../edits';
import {
  decodeTiffImage, initWasm, openPresentation, type PresentationHandle,
} from '../wasm/loader';
import { createPresentationSession, type PresentationSession } from './client';
import { createPresentationSessionHost } from './host';
import {
  PRESENTATION_SESSION_METHODS,
  PRESENTATION_SESSION_POLICIES,
  type PresentationFrame,
  type PresentationSessionEvents,
  type PresentationSessionMethods,
  type PresentationWireFrame,
} from './methods';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;
let tiffFixture: Uint8Array;
let fontBytes: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx, tiff, font] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(import.meta.dir, '../render/fixtures/tiff-image.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = new Uint8Array(pptx);
  tiffFixture = new Uint8Array(tiff);
  fontBytes = new Uint8Array(font);
});

function read(result: PptxReadResult): Extract<PptxReadResult, { ok: true }> {
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

function content(result: PptxReadResult): Omit<Extract<PptxReadResult, { ok: true }>, 'version'> {
  const { version: _, ...value } = read(result);
  return value;
}

function editBatch(main: PresentationHandle, expectVersion: string): PptxEditRequest {
  const steps = main.snapshot().slides.map<PptxEditRequest['steps'][number]>((slide) => {
    const shape = slide.shapes[0];
    if (!shape) throw new Error('Fixture slide has no editable shape');
    return { op: 'setShapeRect', target: { slideId: slide.id, shapeId: shape.id }, rect: {
      x: shape.x + 91440, y: shape.y, width: shape.width, height: shape.height,
    } };
  });
  return { expectVersion, steps };
}

function batchRequests(transport: SessionTransport): SessionTransport {
  return { ...transport, listen(listener) {
    let messages: unknown[] = [];
    let timer: ReturnType<typeof setTimeout> | undefined;
    const off = transport.listen((message) => {
      messages.push(message);
      if (timer !== undefined) return;
      timer = setTimeout(() => {
        timer = undefined;
        const batch = messages;
        messages = [];
        for (const queued of batch) listener(queued);
      }, 0);
    });
    return () => {
      off();
      if (timer !== undefined) clearTimeout(timer);
      messages = [];
    };
  } };
}

async function session(
  clientId: number, bytes = fixture, wrapHost?: (transport: SessionTransport) => SessionTransport
): Promise<PresentationSession> {
  const pair = createInProcessPair();
  createPresentationSessionHost(wrapHost ? wrapHost(pair.host) : pair.host);
  return createPresentationSession(bytes, {
    clientId,
    fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
  }, pair.client);
}

async function matchingFrame(
  worker: PresentationSession, main: PresentationHandle, slideIndex: number
): Promise<PresentationFrame> {
  const frame = await worker.call.frame(slideIndex);
  expect(frame.displayList).toEqual(main.layoutSlide(slideIndex));
  expect(frame.version).toBe(await worker.call.version());
  expect(frame.slideIndex).toBe(slideIndex);
  return frame;
}

describe('presentation sessions', () => {
  test('declares ordered methods and marks edits as user input', () => {
    expect(Object.keys(PRESENTATION_SESSION_METHODS)).toEqual(Object.keys(PRESENTATION_SESSION_POLICIES));
    expect(PRESENTATION_SESSION_POLICIES.applyEdits).toEqual({
      lane: 'input', mutates: true, userInput: true,
    });
    expect(PRESENTATION_SESSION_POLICIES.frame).toEqual({
      lane: 'interactive', reframes: true, key: 'frame', replaceableBy: 'frame',
    });
    for (const policy of Object.values(PRESENTATION_SESSION_POLICIES)) {
      expect(typeof policy).toBe('object');
      expect((policy as MethodPolicy).reorderable).not.toBe(true);
    }
  });

  test('matches main-thread projections, reads, edits and saved bytes', async () => {
    const main = openPresentation(fixture, {
      clientId: 9701, fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
    });
    let worker: PresentationSession | undefined;
    try {
      worker = await session(9701);
      const snapshot = main.snapshot();
      const summaries = snapshot.slides.map((slide, index) => ({
        id: slide.id, index, name: slide.name, layoutPartPath: slide.layoutPartPath,
      }));
      expect(worker.state).toEqual({
        format: 'pptx', stage: 'ready', version: 0, dirty: false,
        slides: summaries, size: { width: snapshot.widthEmu, height: snapshot.heightEmu },
      });
      expect(await worker.call.slides()).toEqual(summaries);
      expect(await worker.call.slideSize()).toEqual(worker.state.size);
      expect(await worker.save()).toEqual(main.save());
      const mainBefore = read(main.readContent());
      const workerBefore = read(await worker.call.readContent());
      expect(await worker.call.version()).toBe(workerBefore.version);
      expect(mainBefore.version).toBe(main.version());
      expect(workerBefore.version).not.toBe(mainBefore.version);
      expect(content(workerBefore)).toEqual(content(mainBefore));
      const story = mainBefore.stories.find((candidate) => candidate.paragraphs[0]?.editable);
      if (!story) throw new Error('Fixture has no editable story');
      const request: PptxEditRequest = {
        expectVersion: workerBefore.version,
        steps: [
          { op: 'insertText', at: 'start', text: 'Session: ', target: {
            kind: 'range', slideId: story.slideId, shapeId: story.shapeId,
            storyId: story.storyId, start: 0, end: 0,
          } },
          { op: 'setSlideNotes', target: { slideId: story.slideId }, text: 'Session notes' },
        ],
      };
      expect(await worker.call.validateEdits(request)).toMatchObject({ ok: true, wouldApply: true });
      expect(worker.state.dirty).toBe(false);
      const applied = await worker.call.applyEdits(request);
      const mainApplied = main.applyEdits({ ...request, expectVersion: mainBefore.version });
      if (!applied.ok || !mainApplied.ok) throw new Error('Parity batch was refused');
      expect(applied.applied).toBe(true);
      expect(applied.baseVersion).toBe(workerBefore.version);
      expect(mainApplied.baseVersion).toBe(mainBefore.version);
      const { baseVersion: workerBase, version: workerVersion, ...workerReceipt } = applied;
      const { baseVersion: mainBase, version: mainVersion, ...mainReceipt } = mainApplied;
      expect(workerBase).not.toBe(mainBase);
      expect(workerVersion).not.toBe(mainVersion);
      expect(workerReceipt).toEqual(mainReceipt);
      expect(await worker.call.version()).toBe(applied.version);
      expect(main.version()).toBe(mainApplied.version);
      expect(applied.version).not.toBe(workerBefore.version);
      expect(mainApplied.version).not.toBe(mainBefore.version);
      const workerAfter = read(await worker.call.readContent());
      const mainAfter = read(main.readContent());
      expect(workerAfter.version).toBe(applied.version);
      expect(mainAfter.version).toBe(mainApplied.version);
      expect(content(workerAfter)).toEqual(content(mainAfter));
      const filtered = read(await worker.call.readContent({ slideIds: [story.slideId] }));
      const mainFiltered = read(main.readContent({ slideIds: [story.slideId] }));
      expect(filtered.version).toBe(applied.version);
      expect(mainFiltered.version).toBe(mainApplied.version);
      expect(content(filtered)).toEqual(content(mainFiltered));
      const found = await worker.call.findText({ text: 'Session: ' });
      const mainFound = main.findText({ text: 'Session: ' });
      expect(found.version).toBe(applied.version);
      expect(mainFound.version).toBe(mainApplied.version);
      const { version: foundVersion, ...matches } = found;
      const { version: mainFoundVersion, ...mainMatches } = mainFound;
      expect(foundVersion).not.toBe(mainFoundVersion);
      expect(matches).toEqual(mainMatches);
      expect(await worker.save()).toEqual(main.save());
      expect<Uint8Array>(new Uint8Array(await worker.call.save())).toEqual(main.save());
      expect(worker.state).toMatchObject({ version: 1, dirty: true });
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  for (const name of ['betteroffice-demo.pptx', 'tiff-image.pptx'] as const) {
    test(`matches every ${name} slide before and after edits with increasing epochs`, async () => {
      const bytes = name === 'betteroffice-demo.pptx' ? fixture : tiffFixture;
      const main = openPresentation(bytes, {
        clientId: 9711, fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
      });
      let worker: PresentationSession | undefined;
      try {
        worker = await session(9711, bytes);
        const before: PresentationFrame[] = [];
        for (const slide of worker.state.slides) {
          const frame = await matchingFrame(worker, main, slide.index);
          expect(frame.epoch).toBe(before.length + 1);
          before.push(frame);
        }
        const request = editBatch(main, before[0].version);
        const applied = await worker.call.applyEdits(request);
        const mainApplied = main.applyEdits({ ...request, expectVersion: main.version() });
        expect(applied).toMatchObject({ ok: true, applied: true });
        expect(mainApplied).toMatchObject({ ok: true, applied: true });
        for (const slide of worker.state.slides) {
          const after = await matchingFrame(worker, main, slide.index);
          expect(after.epoch).toBe(before.length + slide.index + 1);
          expect(after.version).not.toBe(before[slide.index].version);
          expect(after.displayList).not.toEqual(before[slide.index].displayList);
        }
      } finally {
        main.dispose();
        await worker?.dispose();
      }
    });

    test(`transfers ${name} JSON and new media once and exposes cached media`, async () => {
      const bytes = name === 'betteroffice-demo.pptx' ? fixture : tiffFixture;
      const buffers: ArrayBuffer[] = [];
      const sent: PresentationWireFrame['media'][] = [];
      const worker = await session(9712, bytes, (transport) => ({ ...transport, post(message, transfer) {
        if (isHostMessage(message) && message.kind === 'reply' && message.ok &&
          message.value !== null && typeof message.value === 'object' && 'epoch' in message.value) {
          const wire = message.value as PresentationWireFrame;
          expect(wire.displayList).toBeInstanceOf(ArrayBuffer);
          expect(transfer).toEqual([wire.displayList, ...wire.media.map(({ bytes }) => bytes)]);
          expect(JSON.stringify(message)).not.toContain('"primitives"');
          expect(wire.displayList.byteLength).toBeGreaterThan(0);
          for (const asset of wire.media) expect(asset.bytes).toBeInstanceOf(ArrayBuffer);
          buffers.push(wire.displayList, ...wire.media.map(({ bytes }) => bytes));
          sent.push(wire.media.map((asset) => ({ ...asset, bytes: asset.bytes.slice(0) })));
        }
        transport.post(message, transfer);
      } }));
      const main = openPresentation(bytes, {
        clientId: 9712, fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
      });
      try {
        const first = await matchingFrame(worker, main, 0);
        expect(first.media.size).toBeGreaterThan(0);
        expect(sent[0].map(({ assetId }) => assetId)).toEqual([...first.media.keys()]);
        if (name === 'tiff-image.pptx') expect(sent[0]).toHaveLength(1);
        for (const [assetId, media] of first.media) {
          const raw = main.mediaBytes(assetId);
          expect(isTiff(raw)).toBe(name === 'tiff-image.pptx');
          const expected = isTiff(raw) ? decodeTiffImage(raw) : raw;
          expect(media).toEqual(expected);
          expect(new Uint8Array(sent[0].find((asset) => asset.assetId === assetId)!.bytes))
            .toEqual(new Uint8Array(expected));
          if (isTiff(raw)) expect(media.subarray(0, 8)).toEqual(
            new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])
          );
        }
        const later = await matchingFrame(worker, main, 0);
        expect(sent[1]).toEqual([]);
        expect(later.media).toEqual(first.media);
        expect(later.epoch).toBe(first.epoch + 1);
        for (const [assetId, media] of first.media) expect(later.media.get(assetId)).toBe(media);
        if (name === 'betteroffice-demo.pptx') {
          const withoutImages = await matchingFrame(worker, main, 1);
          expect(withoutImages.media.size).toBe(0);
          expect(sent[2]).toEqual([]);
          const revisited = await matchingFrame(worker, main, 0);
          expect(revisited.media).toEqual(first.media);
          expect(sent[3]).toEqual([]);
        }
        expect(buffers.every((buffer) => buffer.byteLength === 0)).toBe(true);
      } finally {
        main.dispose();
        await worker.dispose();
      }
    });
  }

  test('supersedes the middle of three queued frames and remains usable', async () => {
    const worker = await session(9713, fixture, batchRequests);
    try {
      const results = await Promise.allSettled([
        worker.call.frame(0), worker.call.frame(1), worker.call.frame(2),
      ]);
      for (const result of results.slice(0, 2)) {
        if (result.status !== 'rejected') throw new Error('Queued frame was not superseded');
        expect(result.reason).toMatchObject({ name: SESSION_SUPERSEDED });
      }
      const last = results[2];
      if (last.status !== 'fulfilled') throw last.reason;
      expect(last.value.slideIndex).toBe(2);
      expect(last.value.epoch).toBe(1);
      expect(worker.failure).toBeUndefined();
      expect((await worker.call.frame(0)).epoch).toBe(2);
    } finally { await worker.dispose(); }
  });

  test('keeps an edit between queued frames ahead of the later frame', async () => {
    const main = openPresentation(fixture, {
      clientId: 9714, fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
    });
    let worker: PresentationSession | undefined;
    try {
      worker = await session(9714, fixture, batchRequests);
      const before = await matchingFrame(worker, main, 0);
      const request = editBatch(main, before.version);
      const replaced = worker.call.frame(1).catch((error: unknown) => error);
      const editing = worker.call.applyEdits(request);
      const framing = worker.call.frame(0);
      const [dropped, applied, after] = await Promise.all([replaced, editing, framing]);
      expect(dropped).toMatchObject({ name: SESSION_SUPERSEDED });
      if (!applied.ok) throw new Error(applied.failure.message);
      expect(applied.applied).toBe(true);
      expect(main.applyEdits({ ...request, expectVersion: main.version() }))
        .toMatchObject({ ok: true, applied: true });
      expect(after.displayList).toEqual(main.layoutSlide(0));
      expect(after.version).toBe(applied.version);
      expect(after.version).toBe(await worker.call.version());
      expect(after.version).not.toBe(before.version);
      expect(after.epoch).toBe(before.epoch + 1);
      expect(worker.state).toMatchObject({ version: 1, dirty: true });
      expect(worker.failure).toBeUndefined();
    } finally {
      main.dispose();
      await worker?.dispose();
    }
  });

  test('rejects invalid slide indices and leaves the session usable', async () => {
    const worker = await session(9715);
    try {
      const state = worker.state;
      const version = await worker.call.version();
      const saved = await worker.save();
      const changes: PresentationSessionEvents['changed'][] = [];
      worker.on('changed', (change) => { changes.push(change); });
      for (const index of [-1, 0.5, state.slides.length, 2 ** 32, NaN, Infinity]) {
        await expect(worker.call.frame(index)).rejects.toMatchObject({ name: 'RangeError' });
      }
      expect(worker.failure).toBeUndefined();
      const frame = await worker.call.frame(0);
      expect(frame.epoch).toBe(1);
      expect(frame.version).toBe(version);
      expect(worker.state).toEqual(state);
      expect(changes).toEqual([]);
      expect(await worker.save()).toEqual(saved);
    } finally { await worker.dispose(); }
  });

  test('returns refusals as data and orders edits, reads, saves and changed events', async () => {
    const worker = await session(9702);
    const changes: PresentationSessionEvents['changed'][] = [];
    const off = worker.on('changed', (change) => { changes.push(change); });
    try {
      const initial = await worker.call.version();
      const slide = worker.state.slides[0];
      const request: PptxEditRequest = {
        expectVersion: initial,
        steps: [{ op: 'setSlideNotes', target: { slideId: slide.id }, text: 'First notes' }],
      };
      expect(await worker.call.applyEdits({ ...request, expectVersion: 'stale' }))
        .toMatchObject({ ok: false, version: initial, failure: { code: 'stale-version' } });
      expect(await worker.call.validateEdits({ ...request, expectVersion: 'stale' }))
        .toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      const originalNotes = read(await worker.call.readContent()).slides[0].notes ?? '';
      expect(await worker.call.applyEdits({
        expectVersion: initial,
        steps: [{ op: 'setSlideNotes', target: { slideId: slide.id }, text: originalNotes }],
      })).toMatchObject({ ok: true, applied: false, version: initial });
      expect(worker.state).toMatchObject({ version: 0, dirty: false });
      expect(changes).toEqual([]);
      const first = await worker.call.applyEdits(request);
      if (!first.ok) throw new Error(first.failure.message);
      expect(first.applied).toBe(true);
      expect(await worker.call.applyEdits(request)).toMatchObject({
        ok: false, version: first.version, failure: { code: 'stale-version' },
      });
      const noOp = await worker.call.applyEdits({ ...request, expectVersion: first.version });
      expect(noOp).toMatchObject({ ok: true, applied: false, version: first.version });
      expect(changes).toEqual([{ version: 1, dirty: true }]);
      const beforeWrite = worker.call.readContent();
      const saveBefore = worker.save();
      const second = worker.call.applyEdits({
        expectVersion: first.version,
        steps: [{ op: 'setSlideNotes', target: { slideId: slide.id }, text: 'Second notes' }],
      });
      const afterWrite = worker.call.readContent();
      const saveAfter = worker.save();
      const [before, savedBefore, applied, after, savedAfter] =
        await Promise.all([beforeWrite, saveBefore, second, afterWrite, saveAfter]);
      if (!applied.ok) throw new Error(applied.failure.message);
      expect(applied.applied).toBe(true);
      expect(read(before).version).toBe(first.version);
      expect(read(before).slides[0].notes).toBe('First notes');
      expect(read(after).version).toBe(applied.version);
      expect(read(after).slides[0].notes).toBe('Second notes');
      expect(savedBefore).not.toEqual(savedAfter);
      expect(changes).toEqual([{ version: 1, dirty: true }, { version: 2, dirty: true }]);
      expect(worker.state).toMatchObject({ version: 2, dirty: true });
      expect(worker.failure).toBeUndefined();
      off();
      await worker.call.applyEdits({
        expectVersion: applied.version,
        steps: [{ op: 'setSlideNotes', target: { slideId: slide.id }, text: 'Third notes' }],
      });
      expect(changes).toHaveLength(2);
    } finally {
      off();
      await worker.dispose();
    }
    await expect(worker.call.version()).rejects.toMatchObject({ code: 'disposed' });
    await expect(worker.call.frame(0)).rejects.toMatchObject({ code: 'disposed' });
    await expect(worker.save()).rejects.toMatchObject({ code: 'disposed' });
    await worker.dispose();
  });

  test('refuses calls before open, a second open, and calls after RPC disposal', async () => {
    const pair = createInProcessPair();
    let initializations = 0;
    createPresentationSessionHost(pair.host, { initWasm: async (source) => {
      initializations += 1;
      await initWasm(source);
    } });
    const client = createSessionClient<PresentationSessionMethods, {
      changed: PresentationSessionEvents['changed'];
    }>(pair.client, { methods: PRESENTATION_SESSION_METHODS });
    try {
      const calls = [
        () => client.call.version(), () => client.call.readContent(),
        () => client.call.findText({ text: 'text' }),
        () => client.call.validateEdits({ expectVersion: 'stale', steps: [] }),
        () => client.call.applyEdits({ expectVersion: 'stale', steps: [] }),
        () => client.call.frame(0),
        () => client.call.slides(), () => client.call.slideSize(),
        () => client.call.save(), () => client.call.dispose(),
      ];
      for (const call of calls) await expect(call()).rejects.toThrow('not open');
      expect(initializations).toBe(0);
      const bytes = new Uint8Array(fixture).buffer;
      await client.call.open(bytes, { clientId: 9703 });
      await expect(client.call.open(bytes, { clientId: 9703 })).rejects.toThrow('already open');
      expect(initializations).toBe(1);
      await client.call.dispose();
      await expect(client.call.version()).rejects.toThrow('disposed');
      await expect(client.call.frame(0)).rejects.toThrow('disposed');
      await expect(client.call.open(bytes)).rejects.toThrow('disposed');
    } finally { await client.dispose(); }
  });

  test('opens from an initial update with fallback fonts and preserves caller buffers', async () => {
    const source = openPresentation(fixture, { clientId: 9705 });
    try {
      source.setSlideNotes(source.snapshot().slides[0].id, 'Restored notes');
      const update = source.encodeStateAsUpdate();
      const updateBuffer = new Uint8Array(update.byteLength + 16);
      updateBuffer.set(update, 8);
      const initialUpdate = updateBuffer.subarray(8, updateBuffer.byteLength - 8);
      const retainedUpdate = initialUpdate.slice();
      const retainedUpdateBuffer = updateBuffer.slice();
      const faces = [{ family: 'Liberation Sans', bytes: fontBytes }];
      const main = openPresentation(fixture, {
        clientId: 9706, initialUpdate, fonts: faces, fallbackFonts: faces,
      });
      let worker: PresentationSession | undefined;
      try {
        const pair = createInProcessPair();
        createPresentationSessionHost(pair.host);
        const transferred: ArrayBuffer[] = [];
        const transport: SessionTransport = { ...pair.client, post(message, transfer) {
          transferred.push(...(transfer ?? []) as ArrayBuffer[]);
          pair.client.post(message, transfer);
        } };
        worker = await createPresentationSession(fixture, {
          clientId: 9706, initialUpdate, fonts: faces, fallbackFonts: faces,
        }, transport);
        expect(transferred).toHaveLength(4);
        expect(transferred.every((buffer) => buffer.byteLength === 0)).toBe(true);
        expect(initialUpdate).toEqual(retainedUpdate);
        expect(updateBuffer).toEqual(retainedUpdateBuffer);
        expect(faces[0].bytes).toBe(fontBytes);
        expect(fontBytes.byteLength).toBeGreaterThan(0);
        expect(content(await worker.call.readContent())).toEqual(content(main.readContent()));
        expect(await worker.save()).toEqual(main.save());
        expect(worker.state).toMatchObject({ version: 0, dirty: false });
      } finally {
        main.dispose();
        await worker?.dispose();
      }
    } finally { source.dispose(); }
  });

  test('transfers owned copies of document and font buffers, including subviews and other realms', async () => {
    const main = openPresentation(fixture, { clientId: 9704 });
    let saved: Uint8Array;
    try { saved = main.save(); } finally { main.dispose(); }
    for (const kind of ['buffer', 'view', 'foreign'] as const) {
      const asView = kind === 'view';
      const buffer: ArrayBuffer = kind === 'foreign'
        ? runInNewContext('new ArrayBuffer(size)', { size: fixture.byteLength })
        : new ArrayBuffer(fixture.byteLength + (asView ? 16 : 0));
      if (kind === 'foreign') expect(buffer instanceof ArrayBuffer).toBe(false);
      const source = new Uint8Array(buffer);
      source.set(fixture, asView ? 8 : 0);
      const retainedSource = source.slice();
      const document = asView ? source.subarray(8, source.byteLength - 8) : source.buffer;
      const font = new Uint8Array(fontBytes.byteLength + 16);
      font.set(fontBytes, 8);
      const face = font.subarray(8, font.byteLength - 8);
      const pair = createInProcessPair();
      createPresentationSessionHost(pair.host);
      const transferred: ArrayBuffer[] = [];
      const transport: SessionTransport = { ...pair.client, post(message, transfer) {
        transferred.push(...(transfer ?? []) as ArrayBuffer[]);
        pair.client.post(message, transfer);
      } };
      const worker = await createPresentationSession(document, {
        clientId: 9704, fonts: [{ family: 'Liberation Sans', bytes: face }],
      }, transport);
      try {
        expect(transferred).toHaveLength(2);
        expect(transferred.every((buffer) => buffer.byteLength === 0)).toBe(true);
        expect(source.byteLength).toBe(fixture.byteLength + (asView ? 16 : 0));
        expect(source).toEqual(retainedSource);
        expect<Uint8Array>(new Uint8Array(document)).toEqual(fixture);
        expect<Uint8Array>(face).toEqual(fontBytes);
        expect(await worker.save()).toEqual(saved);
      } finally { await worker.dispose(); }
    }
  });

  test('sets the failed stage before notifying listeners of traps and crashes', async () => {
    for (const failure of [
      new WebAssembly.RuntimeError('unreachable'), new SessionFailure('crash', 'Host crashed'),
    ]) {
      const pair = createInProcessPair();
      createSessionHost<Pick<PresentationSessionMethods, 'open' | 'version' | 'dispose'>, {}, null>(
        pair.host, {
          context: null,
          policies: {
            open: PRESENTATION_SESSION_POLICIES.open,
            version: PRESENTATION_SESSION_POLICIES.version,
            dispose: PRESENTATION_SESSION_POLICIES.dispose,
          },
          handlers: {
            open: () => ({
              format: 'pptx', stage: 'ready', version: 0, dirty: false,
              slides: [], size: { width: 0, height: 0 },
            }),
            version: () => { throw failure; },
            dispose: () => {},
          },
        }
      );
      const worker = await createPresentationSession(new Uint8Array(), {}, pair.client);
      const stages: PresentationSession['state']['stage'][] = [];
      worker.onFailure(() => { stages.push(worker.state.stage); });
      try {
        expect(worker.state.stage).toBe('ready');
        await expect(worker.call.version()).rejects.toMatchObject({
          code: failure instanceof SessionFailure ? 'crash' : 'trap',
        });
        expect(stages).toEqual(['failed']);
        expect(worker.state.stage).toBe('failed');
      } finally { await worker.dispose(); }
    }
  });

  test('closes the transport when the caller buffers cannot be copied', async () => {
    const pair = createInProcessPair();
    const detached = new Uint8Array(8);
    structuredClone(detached, { transfer: [detached.buffer] });
    await expect(createPresentationSession(detached, {}, pair.client)).rejects.toThrow();
    expect(() => pair.host.post({})).toThrow();
  });

  test('closes the transport when the session client cannot attach', async () => {
    const pair = createInProcessPair();
    let closed = 0;
    const transport: SessionTransport = {
      post: (message, transfer) => pair.client.post(message, transfer),
      listen: (listener) => pair.client.listen(listener),
      onError() { throw new Error('attach'); },
      close() { closed += 1; pair.client.close(); },
    };
    await expect(createPresentationSession(fixture, {}, transport)).rejects.toThrow('attach');
    expect(closed).toBe(1);
  });
});
