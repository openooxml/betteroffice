import { afterAll, beforeAll, expect, test } from 'bun:test';
import { openWorkbook } from '../wasm/loader';
import { engineKind } from '../wasm/engine';
import { openWorkbookSession } from './client';
import { createTestWorkbookSession, initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('an edit that runs out of memory fails its session for good while a session on another thread keeps working', async () => {
  const bytes = await numbersWorkbook(200);
  const session = await createTestWorkbookSession(bytes);
  limitEngineMemoryHeadroom(4 * 2 ** 20);
  expect((await session.call.sheets()).length).toBe(1);
  const input = 'x'.repeat(30_000);
  const edits = Array.from({ length: 400 }, (_, row) => ({ row, col: 3, input }));
  await expect(session.submit({ ...session.commandIdentity, seq: 1, basis: session.confirmedBasis,
    kind: 'editCells', payload: { args: [0, edits] } })).rejects.toMatchObject({ code: 'out-of-memory' });
  const failure = session.failure;
  expect(failure).toMatchObject({ code: 'out-of-memory' });
  await expect(session.call.sheets()).rejects.toMatchObject({ code: 'out-of-memory' });
  await expect(session.save()).rejects.toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toBe(failure);

  const other = await openWorkbookSession(bytes);
  try {
    expect((await other.call.sheets()).length).toBe(1);
    expect((await other.save()).byteLength).toBeGreaterThan(0);
  } finally {
    await other.dispose();
  }
});

test('a workbook handle whose call ran out of memory refuses every later call', async () => {
  const workbook = openWorkbook(await numbersWorkbook(20_000));
  const read = (a1: string) => workbook.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1 } }] });
  limitEngineMemoryHeadroom(256 * 2 ** 10);
  expect(() => read('A1:B20000')).toThrow();
  restoreEngineMemory();
  expect(() => read('A1')).toThrow();
  expect(() => workbook.save()).toThrow();
  if (engineKind === 'native') expect(() => workbook.dispose()).not.toThrow();
});
