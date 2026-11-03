import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { SessionFailure } from '../../../../shared/office-session';
import { EditSession, engineKind } from '../wasm/engineEdit';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { DocxSourceWorkerClient } from './residentSourceWorkerClient';

let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(engineKind === 'native' ? undefined
    : new Uint8Array(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  startWorker = await residentWorkerFactory();
}, 30_000);

test('an engine panic fails the document session as a trap and refuses later calls', async () => {
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
}, 20_000);
