import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { SessionFailure } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { engineKind } from '../wasm/engine';
import * as wasm from '../wasm/loader';
import { createPresentationSession } from './client';
import { createPresentationSessionHost } from './host';

let fixture: Uint8Array;

beforeAll(async () => {
  await wasm.initWasm(engineKind === 'native' ? undefined
    : await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  fixture = new Uint8Array(await readFile(new URL('../../../../apps/demo/public/betteroffice-demo.pptx', import.meta.url)));
});

test('an engine panic fails the presentation session as a trap and refuses later calls', async () => {
  const pair = createInProcessPair();
  createPresentationSessionHost(pair.host);
  const original = wasm.openPresentation;
  let handle: wasm.PresentationHandle | undefined;
  const opening = spyOn(wasm, 'openPresentation').mockImplementation((...args) => (handle = original(...args)));
  const session = await createPresentationSession(fixture, { clientId: 9705 }, pair.client)
    .finally(() => opening.mockRestore());
  const failing = spyOn(handle!, 'version').mockImplementation(() => {
    throw engineKind === 'native'
      ? Object.assign(new Error('panicked at engine'), { code: 'GenericFailure' })
      : new WebAssembly.RuntimeError('unreachable');
  });
  try {
    await expect(session.call.version()).rejects.toBeInstanceOf(SessionFailure);
    expect(session.failure?.code).toBe('trap');
    expect(session.state.stage).toBe('failed');
    await expect(session.call.version()).rejects.toBe(session.failure);
    await expect(session.save()).rejects.toBe(session.failure);
    expect(failing).toHaveBeenCalledTimes(1);
  } finally {
    failing.mockRestore();
    await session.dispose();
  }
});
