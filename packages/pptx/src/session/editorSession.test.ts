import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createSessionClient } from '../../../../shared/office-session/client';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { initWasm } from '../wasm/loader';
import { PptxDocument } from '../wasm/generated/pptx_wasm.js';
import { createPresentationSession } from './client';
import { createPptxWorkerEditorSession, createPresentationEditorSession } from './editorSession';
import { createPresentationSessionHost } from './host';
import { PRESENTATION_SESSION_METHODS, PRESENTATION_SESSION_POLICIES } from './methods';
import { PptxPeerNotReadyError, PptxWorkerEditorCollaborationError } from './peerHydrationError';
import { PRESENTATION_EDITOR_METHODS, type EditorBaseline, type PresentationEditorMethods } from './replay';

let source: Uint8Array;
let module: WebAssembly.Module;
let font: Uint8Array;
beforeAll(async () => {
  const root = resolve(import.meta.dir, '../../../..');
  const [wasm, deck, face] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  module = await WebAssembly.compile(wasm);
  await initWasm(module);
  source = deck;
  font = face;
});

function host() {
  const pair = createInProcessPair();
  createPresentationSessionHost(pair.host, { initEditorWasm: async (input) => {
    expect(input).toBeInstanceOf(WebAssembly.Module);
    await initWasm(input);
    return input as WebAssembly.Module;
  } });
  return pair;
}

test('hydrates_exactly_one_peer_eagerly', async () => {
  const pair = host();
  const opened = spyOn(PptxDocument, 'openPeerDeckJson');
  const fonts = [{ family: 'Liberation Sans', bytes: font }];
  const owner = createPresentationEditorSession(source, { wasm: module, fonts, fallbackFonts: fonts }, pair.client);
  const stages: string[] = [owner.state.stage];
  owner.subscribe((state) => { stages.push(state.stage); });
  try {
    expect(() => owner.apply({ method: 'insertSlide', args: [1] })).toThrow(PptxPeerNotReadyError);
    await owner.whenHydrated();
    const access = await owner.handleAsync();
    expect(opened).toHaveBeenCalledTimes(1);
    expect(stages).toEqual(['opening', 'hydrating', 'ready']);
    expect(owner.state.initialFrame).toMatchObject({ sequence: 0, version: access.version() });
    expect(access.undoCaptureMode()).toBe('manual');
    expect(access.interaction.snapshot().snapshot).toEqual(access.snapshot());
    expect('save' in access).toBe(false);
    expect('dispose' in access).toBe(false);
    await expect(owner.frame('missing')).rejects.toMatchObject({ name: 'RangeError' });
    expect(owner.state.stage).toBe('ready');
    await owner.handleAsync();
    expect(opened).toHaveBeenCalledTimes(1);
  } finally { opened.mockRestore(); await owner.dispose(); }

  const partial = host();
  const original = PptxDocument.openPeerDeckJson;
  const free = spyOn(PptxDocument.prototype, 'free');
  const peerOpen = spyOn(PptxDocument, 'openPeerDeckJson').mockImplementation((...args) => {
    const doc = original(...args);
    queueMicrotask(() => { void cancelled.dispose(); });
    return doc;
  });
  const cancelled = createPresentationEditorSession(source, { wasm: module }, partial.client);
  try {
    await expect(cancelled.whenHydrated()).rejects.toMatchObject({ code: 'editor-disposed' });
    await cancelled.dispose();
    expect(peerOpen).toHaveBeenCalledTimes(1);
    expect(free.mock.calls.length).toBeGreaterThanOrEqual(2);
    expect(cancelled.state.stage).toBe('disposed');
  } finally { peerOpen.mockRestore(); free.mockRestore(); await cancelled.dispose(); }
});

test('input_requires_matching_attachment', async () => {
  for (const mismatch of ['version', 'sequence', 'twice'] as const) {
    const pair = host();
    const client = createSessionClient<PresentationEditorMethods, {}>(pair.client, { methods: PRESENTATION_EDITOR_METHODS });
    try {
      const baseline = await client.call.beginEditor(new Uint8Array(source).buffer, { wasm: module });
      await expect(client.call.flush(0, baseline.version)).rejects.toMatchObject({ refusal: { code: 'stage' } });
      if (mismatch === 'twice') await client.call.attachPeer(baseline.version, 0);
      await expect(client.call.attachPeer(mismatch === 'version' ? 'wrong' : baseline.version,
        mismatch === 'sequence' ? 1 : 0)).rejects.toMatchObject({ name: 'PptxPeerHydrationError' });
      await expect(client.call.attachPeer(baseline.version, 0)).rejects.toThrow();
      await expect(client.call.flush(0, baseline.version)).rejects.toThrow();
    } finally { await client.dispose(); }
  }
  const rejecting = createInProcessPair();
  createPresentationSessionHost({ ...rejecting.host, post(message, transfer) {
    if (isHostMessage(message) && message.kind === 'reply' && message.ok &&
      typeof message.value === 'object' && message.value !== null && 'hydration' in message.value) {
      const baseline = message.value as EditorBaseline;
      message = { ...message, value: { ...baseline, version: 'wrong' } };
    }
    rejecting.host.post(message, transfer);
  } }, { initEditorWasm: async () => module });
  const free = spyOn(PptxDocument.prototype, 'free');
  const owner = createPresentationEditorSession(source, { wasm: module }, rejecting.client);
  try {
    await expect(owner.whenHydrated()).rejects.toMatchObject({ cause: { code: 'version' } });
    expect(owner.state.stage).toBe('failed');
    expect(free.mock.calls.length).toBeGreaterThan(0);
  } finally { free.mockRestore(); await owner.dispose(); }
});

test('collaboration_error_precedes_worker_creation', () => {
  let constructed = 0;
  const initialize = spyOn(WebAssembly, 'instantiate');
  try {
    expect(() => createPptxWorkerEditorSession(source, {
      collaboration: {}, worker: () => { constructed += 1; throw new Error('worker'); },
    })).toThrow(PptxWorkerEditorCollaborationError);
    expect(constructed).toBe(0);
    expect(initialize).not.toHaveBeenCalled();
  } finally { initialize.mockRestore(); }
});

test('cold_editor_hydration_reuses_worker_module', async () => {
  const script = `
    import assert from 'node:assert/strict';
    import { createInProcessPair } from ${JSON.stringify(resolve(import.meta.dir, '../../../../shared/office-session/testing/inProcessTransport.ts'))};
    import { createPresentationSessionHost, initializePresentationEditorWasm } from ${JSON.stringify(resolve(import.meta.dir, 'host.ts'))};
    import { createPresentationEditorSession } from ${JSON.stringify(resolve(import.meta.dir, 'editorSession.ts'))};
    import { wasmVersion } from ${JSON.stringify(resolve(import.meta.dir, '../wasm/loader.ts'))};
    assert.throws(() => wasmVersion(), /wasm is not initialized/);
    const wasm = await Bun.file(${JSON.stringify(resolve(import.meta.dir, '../wasm/generated/pptx_wasm_bg.wasm'))}).arrayBuffer();
    const source = await Bun.file(${JSON.stringify(resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.pptx'))}).arrayBuffer();
    const font = new Uint8Array(await Bun.file(${JSON.stringify(resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))}).arrayBuffer());
    const compile = WebAssembly.compile.bind(WebAssembly);
    let compilations = 0;
    WebAssembly.compile = (bytes) => { compilations++; return compile(bytes); };
    const pair = createInProcessPair();
    const calls = [];
    let module;
    createPresentationSessionHost(pair.host, { initEditorWasm: async (input) => {
      assert.equal(input, undefined);
      module = await initializePresentationEditorWasm(wasm);
      return module;
    } });
    const owner = createPresentationEditorSession(source, {
      clientId: 7101, fonts: [{ family: 'Liberation Sans', bytes: font }],
    }, {
      ...pair.client,
      post(message, transfer) { calls.push(message.method ?? message.kind); pair.client.post(message, transfer); },
      listen(listener) { return pair.client.listen((message) => {
        if (message.value?.module) assert.ok(message.value.module instanceof WebAssembly.Module);
        listener(message);
      }); },
    });
    try {
      await owner.whenHydrated();
      assert.ok(module instanceof WebAssembly.Module);
      assert.equal(compilations, 1);
      assert.deepEqual(calls, ['beginEditor', 'attachPeer', 'editorFrame']);
      assert.equal((await owner.handleAsync()).clientId, 7101);
    } finally { await owner.dispose(); }
  `;
  const child = Bun.spawn([process.execPath, '--eval', script], {
    cwd: resolve(import.meta.dir, '../../../..'), stdout: 'ignore', stderr: 'pipe',
  });
  const error = await new Response(child.stderr).text();
  expect({ exit: await child.exited, error }).toEqual({ exit: 0, error: '' });
});

test('ordinary_session_wire_shapes_are_unchanged', async () => {
  expect(Object.keys(PRESENTATION_SESSION_METHODS)).toEqual([
    'open', 'version', 'readContent', 'findText', 'validateEdits', 'applyEdits', 'frame',
    'slides', 'slideSize', 'save', 'dispose',
  ]);
  expect(PRESENTATION_SESSION_POLICIES.frame).toEqual({
    lane: 'interactive', reframes: true, key: 'frame', replaceableBy: 'frame',
  });
  const pair = host();
  const session = await createPresentationSession(source, {
    wasm: module, fonts: [{ family: 'Liberation Sans', bytes: font }],
  }, pair.client);
  try {
    expect(Object.keys(session.state).sort()).toEqual(['dirty', 'format', 'size', 'slides', 'stage', 'version']);
    expect(Object.keys(await session.call.frame(0)).sort()).toEqual([
      'displayList', 'epoch', 'media', 'slideIndex', 'version',
    ]);
    const result = await session.call.applyEdits({ expectVersion: await session.call.version(), steps: [] });
    expect(result.ok).toBe(true);
    expect(await session.save()).toBeInstanceOf(Uint8Array);
  } finally { await session.dispose(); }
});
