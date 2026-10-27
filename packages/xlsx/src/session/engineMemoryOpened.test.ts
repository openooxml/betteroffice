import { beforeAll, expect, test } from 'bun:test';
import { createTestWorkbookSession, exhaustEngineMemory, initRecordedWasm, numbersWorkbook } from './testHelpers';

beforeAll(initRecordedWasm);

test('an engine out of memory after open fails the session as out-of-memory, not a bare trap', async () => {
  const session = await createTestWorkbookSession(await numbersWorkbook(200));
  const failed = new Promise((resolve) => session.onFailure(resolve));
  exhaustEngineMemory();
  const request = { expectVersion: 'x'.repeat(128 * 2 ** 20), steps: [] };
  await expect(session.call.validateEdits(request)).rejects.toMatchObject({ code: 'out-of-memory' });
  expect(await failed).toMatchObject({ code: 'out-of-memory', message: expect.stringContaining('too large') });
});
