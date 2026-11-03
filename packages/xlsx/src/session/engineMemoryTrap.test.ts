import { afterAll, beforeAll, expect, spyOn, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session';
import * as wasm from '../wasm/loader';
import { engineKind } from '../wasm/engine';
import { createTestWorkbookSession, initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('a trap after an earlier engine out of memory on the same thread stays a trap', async () => {
  if (engineKind === 'native') {
    const earlier = wasm.openWorkbook(await numbersWorkbook(20_000));
    limitEngineMemoryHeadroom(256 * 2 ** 10);
    expect(() => earlier.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:B20000' } }] })).toThrow();
    restoreEngineMemory();
    earlier.dispose();
  }
  const original = wasm.openWorkbook;
  let handle: wasm.WorkbookHandle | undefined;
  const opening = spyOn(wasm, 'openWorkbook').mockImplementation((...args) => (handle = original(...args)));
  const session = await createTestWorkbookSession(await numbersWorkbook(20)).finally(() => opening.mockRestore());
  try {
    spyOn(handle!, 'setNumberFormat').mockImplementation(() => { throw new WebAssembly.RuntimeError('unreachable'); });
    await expect(session.submit({ ...session.commandIdentity, seq: 1, basis: session.confirmedBasis,
      kind: 'setNumberFormat', payload: { args: [0, 'A1', { type: 'custom', pattern: '0.0000' }] } })).rejects.toBeInstanceOf(SessionFailure);
    expect(session.failure?.code).toBe('trap');
  } finally { await session.dispose(); }
});
