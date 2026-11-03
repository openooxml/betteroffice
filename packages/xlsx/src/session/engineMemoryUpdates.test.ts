import { afterAll, beforeAll, expect, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session';
import { openWorkbook } from '../wasm/loader';
import { engineOutOfMemory } from './host';
import { initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('an edit that runs out of memory with an update listener is classified as out of memory', async () => {
  const workbook = openWorkbook(await numbersWorkbook(20_000));
  workbook.onUpdate(() => {});
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  let failure: unknown;
  try { workbook.applyOps([{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]); } catch (error) { failure = error; }
  expect(failure).toBeInstanceOf(Error);
  expect(engineOutOfMemory(new SessionFailure('trap', (failure as Error).message))?.code).toBe('out-of-memory');
});
