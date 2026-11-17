import { afterAll, beforeAll, expect, test } from 'bun:test';
import { engineKind } from '../wasm/engine';
import { createTestWorkbookSession, initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, ranInOwnEngineProcess, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

const name = 'a native save that crosses the memory budget only while it runs fails the session as out-of-memory';
test.skipIf(engineKind !== 'native')(name, async () => {
  if (await ranInOwnEngineProcess(import.meta.path, name)) return;
  const session = await createTestWorkbookSession(await numbersWorkbook(20_000));
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  await expect(session.save()).rejects.toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toMatchObject({ code: 'out-of-memory' });
  restoreEngineMemory();
  await expect(session.save()).rejects.toMatchObject({ code: 'out-of-memory' });
}, 60_000);
