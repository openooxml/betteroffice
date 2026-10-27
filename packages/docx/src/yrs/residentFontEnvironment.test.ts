import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from './wasm/index';
import { preloadOpcWasm } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { DocxSourceWorkerClient as ResidentEngineWorkerClient } from './residentSourceWorkerClient';
import { createResidentEngineSession } from './residentEngineSession';
import { setEditWasmHeapLimit } from '../wasm/edit';
import { wasmModuleMemories } from './index';

const TIMEOUT = 20_000;
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
const font = (extension: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `crates/docx-edit/tests/fonts/LiberationSans-Regular.${extension}`)));

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
  startWorker = await residentWorkerFactory();
});

test('oversized WOFF table leaves wasm memory and a bounded editing session intact', async () => {
  const engine = await createResidentEngineSession(16 * 1024 * 1024);
  try {
    engine.openDocument({ package: { document: { content: [{ type: 'paragraph', content: [
      { type: 'run', content: [{ type: 'text', text: 'Text' }] },
    ] }] } } }, {});
    engine.decodeFont(font('woff'));
    const bytes = new Uint8Array(65);
    bytes.set(new TextEncoder().encode('wOFF'));
    const view = new DataView(bytes.buffer);
    view.setUint32(4, 0x10000);
    view.setUint32(8, 65);
    view.setUint16(12, 1);
    view.setUint32(16, 28);
    view.setUint32(48, 64);
    view.setUint32(52, 1);
    view.setUint32(56, 65 * 1024 * 1024);
    const before = wasmModuleMemories().find(entry => entry.label === 'docx-edit')!;
    expect(() => engine.decodeFont(bytes)).toThrow('Invalid resident font size');
    const after = wasmModuleMemories().find(entry => entry.label === 'docx-edit')!;
    expect(after.bufferBytes).toBe(before.bufferBytes);
    expect(after.liveBytes! - before.liveBytes!).toBeLessThan(65536);
    expect(after.failedAllocationBytes).toBe(0);
    engine.insertText({ story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 }, 'Still ');
    expect(engine.paragraphs('body')[0]!.text).toBe('Still Text');
  } finally { engine.destroy(); setEditWasmHeapLimit(undefined); }
}, TIMEOUT);

for (const extension of ['woff', 'woff2']) {
  test(`${extension} font environments decode in the worker and publish readiness and presentation`, async () => {
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    const events: unknown[] = [];
    const memory: unknown[] = [];
    const document = { package: { document: { content: [{ type: 'paragraph' as const, content: [{ type: 'run' as const, content: [{ type: 'text' as const, text: 'Font test' }] }] }] } } };
    try {
      await client.openSource({ document });
      const initial = await client.commandInit({
        session: 'font-worker', incarnation: 1, epoch: 1, readOnly: true,
        fonts: [font('ttf')], fontRegistrationSlots: true,
        layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ properties: {} }] },
          renderEnv: {}, measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } } }),
        layoutWithRegions: true, layoutRevision: 1, extras: '{}', displayWindow: [0, 1],
      });
      const version = worker.sessions[0]!.proposalEngine.version();
      const unsubscribe = client.onFontEnvironment(event => events.push(event));
      client.onMemory(sample => memory.push(sample));
      const envelope = {
        session: initial.session, incarnation: initial.incarnation, seq: 1,
        basis: { epoch: initial.epoch, version: initial.version }, kind: 'setFontEnvironment',
        payload: { faces: [{ family: 'Liberation Sans', data: font(extension) }],
          chains: { 'liberation sans|0|0': [0] }, families: ['Liberation Sans'] },
      };
      const ack = await client.edit(envelope);
      expect(ack.outcomes[0]?.status).toBe('noop');
      expect(ack.version).toBe(initial.version);
      expect(worker.sessions[0]!.proposalEngine.version()).toBe(version);
      expect(ack.frame?.payload.fontEnvironment).toEqual({ revision: 1, families: ['Liberation Sans'] });
      expect(worker.fontFaces.size).toBe(1);
      expect(events).toMatchObject([{ phase: 'ready', revision: 1 }, { phase: 'presentation', revision: 1, frameEpoch: ack.frame?.revision }]);
      expect(memory.length).toBeGreaterThan(0);
      expect(client.memory()?.some(sample => sample.label === 'docx-edit')).toBe(true);
      const duplicate = await client.edit(envelope);
      expect(duplicate.version).toBe(ack.version);
      expect(events).toHaveLength(2);
      const oldFace = [...worker.fontFaces][0];
      const replaced = await client.edit({ ...envelope, seq: 2 });
      expect(replaced.frame?.payload.fontEnvironment?.revision).toBe(2);
      expect(worker.fontFaces.size).toBe(1);
      expect(worker.fontFaces.has(oldFace!)).toBe(false);
      expect(events).toHaveLength(4);
      unsubscribe();
      const refused = await client.edit({ ...envelope, seq: 3, payload: { ...envelope.payload, faces: [{ family: 'Bad', data: new Uint8Array(24) }] } });
      expect(refused.outcomes[0]?.status).toBe('refused');
      expect(refused.frame?.payload.fontEnvironment?.revision).toBe(2);
      expect(events).toHaveLength(4);
    } finally {
      client.destroy();
      expect(worker.fontFaces.size).toBe(0);
    }
  }, TIMEOUT);
}

test('decoded worker fonts stay in their owning measurement session', async () => {
  const { createResidentEngineSession } = await import('./residentEngineSession');
  const first = await createResidentEngineSession();
  const second = await createResidentEngineSession();
  try {
    const decoded = first.decodeFont(font('woff2'));
    const id = first.registerFont(decoded);
    expect(first.outlineGlyphJson(id, 1)).toBeDefined();
    expect(() => second.outlineGlyphJson(id, 1)).toThrow();
  } finally {
    first.destroy();
    second.destroy();
  }
}, TIMEOUT);
