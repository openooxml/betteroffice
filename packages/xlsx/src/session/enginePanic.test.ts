import { beforeAll, expect, spyOn, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session';
import { engineKind } from '../wasm/engine';
import * as wasm from '../wasm/loader';
import type { WorkbookCommand } from './commandProtocol';
import { createTestWorkbookSession, initRecordedWasm, numbersWorkbook } from './testHelpers';

beforeAll(async () => {
  if (engineKind === 'native') await wasm.initWasm();
  else await initRecordedWasm();
});

test('an engine panic fails the workbook session as a trap and refuses later calls', async () => {
  const original = wasm.openWorkbook;
  let handle: wasm.WorkbookHandle | undefined;
  const opening = spyOn(wasm, 'openWorkbook').mockImplementation((...args) => (handle = original(...args)));
  const session = await createTestWorkbookSession(await numbersWorkbook(20)).finally(() => opening.mockRestore());
  const failing = spyOn(handle!, 'setNumberFormat').mockImplementation(() => {
    throw engineKind === 'native'
      ? Object.assign(new Error('panicked at engine'), { code: 'GenericFailure' })
      : new WebAssembly.RuntimeError('unreachable');
  });
  try {
    const command: WorkbookCommand = { ...session.commandIdentity, seq: 1, basis: session.confirmedBasis,
      kind: 'setNumberFormat', payload: { args: [0, 'A1', { type: 'custom', pattern: '0.0000' }] } };
    await expect(session.submit(command)).rejects.toBeInstanceOf(SessionFailure);
    expect(session.failure?.code).toBe('trap');
    await expect(session.submit({ ...command, seq: 2 })).rejects.toBe(session.failure);
    await expect(session.save()).rejects.toBe(session.failure);
    expect(failing).toHaveBeenCalledTimes(1);
  } finally {
    failing.mockRestore();
    await session.dispose();
  }
});
