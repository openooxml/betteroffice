import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { SessionFailure } from '../../../../shared/office-session';
import { NATIVE_MEMORY_BUDGET_EXCEEDED } from '../../../../shared/office-session/trap';
import {
  EditSession,
  engineKind,
  rezip_docx,
  set_wasm_heap_counting,
  wasm_failed_allocation_bytes,
} from '../wasm/engineEdit';
import { preloadEditWasm, setEditWasmHeapLimit } from '../wasm/edit';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import type { DocxCommandEnvelope } from './residentCommandProtocol';
import { DocxSourceWorkerClient } from './residentSourceWorkerClient';

let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(engineKind === 'native' ? undefined
    : new Uint8Array(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  startWorker = await residentWorkerFactory();
}, 30_000);

async function expectPanicToTrap(): Promise<void> {
  const worker = startWorker();
  const client = new DocxSourceWorkerClient(worker);
  try {
    await client.openSource({ document: { package: { document: { content: [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Panic test' }] }] },
    ] } } } });
    const font = new Uint8Array(readFileSync(new URL(
      '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url
    )));
    await client.commandInit({
      session: 'engine-panic', incarnation: 1, epoch: 1, readOnly: false,
      fonts: [font], fontRegistrationSlots: true,
      layoutInput: JSON.stringify({
        bodyStory: 'body', regions: { sections: [{ properties: {} }] }, renderEnv: {},
        measurement: {
          defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
          fontChains: { 'liberation sans|0|0': [0] },
        },
      }),
      layoutWithRegions: true, layoutRevision: 1, extras: '{}', displayWindow: [0, 1],
    });
    const posted = spyOn(worker, 'postMessage');
    const failing = spyOn(EditSession.prototype, 'encode_state').mockImplementation(() => {
      throw engineKind === 'native'
        ? Object.assign(new Error('panicked at engine'), { code: 'GenericFailure' })
        : new WebAssembly.RuntimeError('unreachable');
    });
    try {
      await expect(client.encodeState()).rejects.toBeInstanceOf(SessionFailure);
      expect(client.commandFailure).toMatchObject({ code: 'trap' });
      expect(client.hasFailed()).toBe(true);
      await expect(client.encodeState()).rejects.toBe(client.commandFailure);
      expect(failing).toHaveBeenCalledTimes(1);
      expect(posted).toHaveBeenCalledTimes(1);
    } finally {
      failing.mockRestore();
      posted.mockRestore();
    }
  } finally {
    client.destroy();
  }
}

test('an engine panic fails the document session as a trap and refuses later calls', expectPanicToTrap, 20_000);

// A wasm module that runs out of memory is gone, so only a native thread can panic again after one.
test.skipIf(engineKind !== 'native')('an engine panic after a call ran past the heap limit is still a trap', async () => {
  expect(set_wasm_heap_counting(true)).toBe(true);
  setEditWasmHeapLimit(0);
  try {
    expect(() => rezip_docx({ 'part.bin': new Uint8Array(1024) })).toThrow(NATIVE_MEMORY_BUDGET_EXCEEDED);
  } finally {
    setEditWasmHeapLimit(undefined);
  }
  expect(wasm_failed_allocation_bytes()).toBeGreaterThan(0);
  await expectPanicToTrap();
}, 20_000);

test.skipIf(engineKind !== 'native')('a command edit past the heap limit fails the session as out of memory', async () => {
  const worker = startWorker();
  const client = new DocxSourceWorkerClient(worker);
  try {
    await client.openSource({ document: { package: { document: { content: [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Budget test' }] }] },
    ] } } } });
    const point = { story: 'body', paraId: worker.sessions[0]!.paragraphs('body')[0]!.paraId, offset: 0 };
    const font = new Uint8Array(readFileSync(new URL(
      '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url
    )));
    const initial = await client.commandInit({
      session: 'engine-budget', incarnation: 1, epoch: 1, readOnly: false,
      selection: { anchor: point, focus: point },
      fonts: [font], fontRegistrationSlots: true,
      layoutInput: JSON.stringify({
        bodyStory: 'body', regions: { sections: [{ properties: {} }] }, renderEnv: {},
        measurement: {
          defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
          fontChains: { 'liberation sans|0|0': [0] },
        },
      }),
      layoutWithRegions: true, layoutRevision: 1, extras: '{}', displayWindow: [0, 1],
    });
    const command: DocxCommandEnvelope = {
      session: initial.session, incarnation: initial.incarnation, seq: 1,
      basis: { epoch: initial.epoch, version: initial.version }, kind: 'insertText', payload: { text: 'x' },
    };
    setEditWasmHeapLimit(0);
    try {
      await expect(client.edit(command)).rejects.toMatchObject({ code: 'out-of-memory' });
    } finally {
      setEditWasmHeapLimit(undefined);
    }
    expect(client.commandFailure).toMatchObject({ code: 'out-of-memory' });
  } finally {
    client.destroy();
  }
}, 20_000);
