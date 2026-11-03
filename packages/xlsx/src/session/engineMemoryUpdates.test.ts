import { afterAll, beforeAll, expect, test } from 'bun:test';
import { engineAllocationFailure, openWorkbook } from '../wasm/loader';
import { initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('an edit that runs out of memory with an update listener keeps its allocation failure', async () => {
  const workbook = openWorkbook(await numbersWorkbook(20_000));
  workbook.onUpdate(() => {});
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  expect(() => workbook.applyOps([{ type: 'insertRows', sheet: 0, at: 0, count: 1 }])).toThrow();
  expect(engineAllocationFailure()).not.toBeNull();
});
