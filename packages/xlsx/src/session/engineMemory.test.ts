import { afterAll, beforeAll, expect, test } from 'bun:test';
import { createTestWorkbookSession, exhaustEngineMemory, initRecordedWasm, restoreEngineMemory, numbersWorkbook } from './testHelpers';

beforeAll(initRecordedWasm);
afterAll(restoreEngineMemory);

test('an engine out of memory while opening fails the session as out-of-memory, not a bare trap', async () => {
  const bytes = await numbersWorkbook(20_000);
  exhaustEngineMemory();
  await expect(createTestWorkbookSession(bytes)).rejects.toMatchObject({
    name: 'SessionFailure', code: 'out-of-memory', message: expect.stringContaining('too large'),
  });
});
