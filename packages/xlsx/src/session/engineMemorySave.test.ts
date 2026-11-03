import { afterAll, beforeAll, expect, test } from 'bun:test';
import { createTestWorkbookSession, initRecordedWasm, limitEngineMemoryHeadroom, numbersWorkbook, restoreEngineMemory } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('a save that runs out of memory only while it runs fails the session as out-of-memory', async () => {
  const session = await createTestWorkbookSession(await numbersWorkbook(20_000));
  limitEngineMemoryHeadroom(64 * 2 ** 10);
  await expect(session.save()).rejects.toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toMatchObject({ code: 'out-of-memory' });
  restoreEngineMemory();
  await expect(session.save()).rejects.toMatchObject({ code: 'out-of-memory' });
});
