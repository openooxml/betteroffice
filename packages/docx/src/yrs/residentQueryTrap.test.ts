import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../layout/render/frameDelta';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import type { DocxCommandEnvelope } from './residentCommandProtocol';
import type { DocxEngineWorkerResponse } from './residentEngineWorkerProtocol';
import { DocxSourceWorkerClient } from './residentSourceWorkerClient';

const TIMEOUT = 20_000;
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});

test('a display-list query trap terminates the worker instead of acknowledging fallback selection', async () => {
  const worker = startWorker();
  const client = new DocxSourceWorkerClient(worker);
  const replies: DocxEngineWorkerResponse[] = [];
  const failures: Error[] = [];
  const queryEngine = await loadRustDisplayListQueryEngine();
  const font = new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
  )));
  try {
    await client.openSource({ document: { package: { document: { content: [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Query trap' }] }] },
    ] } } } });
    const point = { story: 'body', paraId: worker.sessions[0]!.paragraphs('body')[0]!.paraId, offset: 0 };
    const initial = await client.commandInit({
      session: 'query-trap', incarnation: 1, epoch: 1, readOnly: false,
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
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(initial.frame.payload.frame));
    const page = frame.pages[0]!;
    const command: DocxCommandEnvelope = {
      session: initial.session, incarnation: initial.incarnation, seq: 1,
      basis: { epoch: initial.epoch, version: initial.version }, kind: 'selectPoint',
      payload: {
        page: { pageId: page.pageId.toString(), pageIndex: 0 }, x: 96, y: 100,
        layoutRevision: initial.frame.payload.layoutRevision, frameEpoch: initial.frame.revision,
        clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
      },
    };
    const receive = worker.onmessage!;
    worker.onmessage = event => { replies.push(event.data); receive(event); };
    client.onFailure(error => failures.push(error));
    const posted = spyOn(worker, 'postMessage');
    const hit = spyOn(queryEngine, 'hitTestRegionsByHandle').mockImplementation(() => {
      throw new WebAssembly.RuntimeError('query allocation failed');
    });
    const log = spyOn(console, 'error').mockImplementation(() => {});
    try {
      await expect(client.edit(command)).rejects.toMatchObject({
        code: 'trap', message: 'Document worker trapped: query allocation failed',
      });
      expect(hit).toHaveBeenCalledTimes(1);
      const request = posted.mock.calls[0]![0];
      if (request.type !== 'edit') throw new Error('Expected a worker edit request');
      expect(replies.filter(reply => reply.id === request.id)).toEqual([
        expect.objectContaining({ id: request.id, ok: false, terminal: true }),
      ]);
      expect(replies.some(reply => reply.ok && reply.commandAck)).toBe(false);
      expect(client.hasFailed()).toBe(true);
      expect(client.commandFailure).toBeDefined();
      expect(failures).toEqual([client.commandFailure!]);
      await expect(client.edit({ ...command, seq: 2 })).rejects.toBe(client.commandFailure);
      expect(posted).toHaveBeenCalledTimes(1);
    } finally {
      log.mockRestore();
      hit.mockRestore();
      posted.mockRestore();
    }
  } finally {
    client.destroy();
  }
}, TIMEOUT);
