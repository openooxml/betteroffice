import { afterAll, beforeAll, expect, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session';
import { openWorkbook } from '../wasm/loader';
import { engineOutOfMemory } from './host';
import { initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('a call that runs out of memory and then fails as an invalid request is classified as out of memory', async () => {
  const workbook = openWorkbook(await numbersWorkbook(20));
  const input = 'x'.repeat(100);
  const edits = [...Array.from({ length: 5_000 }, (_, row) => ({ row, col: 3, input })), { row: -1, col: 3, input }];
  limitEngineMemoryHeadroom(Math.round(JSON.stringify({ sheet: 0, edits }).length * 1.5));
  let failure: unknown;
  try { workbook.editCells(0, edits); } catch (error) { failure = error; }
  restoreEngineMemory();
  expect(failure).toBeInstanceOf(Error);
  expect(engineOutOfMemory(new SessionFailure('trap', (failure as Error).message))?.code).toBe('out-of-memory');
});
