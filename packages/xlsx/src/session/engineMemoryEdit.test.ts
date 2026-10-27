import { beforeAll, expect, test } from 'bun:test';
import { createTestWorkbookSession, exhaustEngineMemory, initRecordedWasm, numbersWorkbook } from './testHelpers';

beforeAll(initRecordedWasm);

test('an engine out of memory in a cell edit after open fails the session as out-of-memory, not a bare trap', async () => {
  const session = await createTestWorkbookSession(await numbersWorkbook(200));
  exhaustEngineMemory();
  const input = 'x'.repeat(30_000);
  const edits = Array.from({ length: 400 }, (_, row) => ({ row, col: 3, input }));
  await expect(session.submit({ ...session.commandIdentity, seq: 1, basis: session.confirmedBasis,
    kind: 'editCells', payload: { args: [0, edits] } })).rejects.toMatchObject({ code: 'out-of-memory' });
  expect(session.failure).toMatchObject({ code: 'out-of-memory', message: expect.stringContaining('too large') });
});
