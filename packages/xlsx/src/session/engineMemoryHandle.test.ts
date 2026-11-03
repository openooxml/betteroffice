import { afterAll, beforeAll, expect, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session';
import { isEngineTrap } from '../../../../shared/office-session/trap';
import { openWorkbook } from '../wasm/loader';
import { engineOutOfMemory } from './host';
import { engineKind } from '../wasm/engine';
import { initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('a workbook handle whose call ran out of memory refuses every later call as an out-of-memory trap', async () => {
  const workbook = openWorkbook(await numbersWorkbook(20_000));
  const read = (a1: string) => workbook.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1 } }] });
  limitEngineMemoryHeadroom(256 * 2 ** 10);
  expect(() => read('A1:B20000')).toThrow();
  restoreEngineMemory();
  let refused: unknown;
  try { read('A1'); } catch (error) { refused = error; }
  expect(refused).toBeInstanceOf(Error);
  if (engineKind === 'native') expect(isEngineTrap(refused)).toBe(true);
  expect(engineOutOfMemory(new SessionFailure('trap', (refused as Error).message))?.code).toBe('out-of-memory');
  expect(() => workbook.save()).toThrow();
  if (engineKind === 'native') expect(() => workbook.dispose()).not.toThrow();
});
