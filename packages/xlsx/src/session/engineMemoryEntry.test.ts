import { afterAll, beforeAll, expect, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session';
import { openWorkbook } from '../wasm/loader';
import { engineKind } from '../wasm/engine';
import { engineOutOfMemory } from './host';
import { initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, ranInOwnEngineProcess, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

const name = 'a call whose arguments run out of memory stays out of memory after a later call succeeds';
test(name, async () => {
  if (await ranInOwnEngineProcess(import.meta.path, name)) return;
  const workbook = openWorkbook(await numbersWorkbook(20));
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  let failure: unknown;
  try { workbook.editCell(0, 0, 0, 'x'.repeat(2 ** 20)); } catch (error) { failure = error; }
  expect(failure).toBeInstanceOf(Error);
  restoreEngineMemory();
  if (engineKind === 'native') expect(workbook.sheetCount()).toBe(1);
  expect(engineOutOfMemory(new SessionFailure('trap', (failure as Error).message))?.code).toBe('out-of-memory');
}, 60_000);
