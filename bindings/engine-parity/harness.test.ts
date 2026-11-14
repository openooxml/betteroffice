import { describe, expect, test } from 'bun:test';
import { runRow, type EngineHandle, type Row } from './harness';

async function fakeEngine(): Promise<EngineHandle<object>> {
  return {
    target: {},
    generated: () => [],
    save: async () => new Uint8Array(),
    free: () => {},
  };
}

function throwingRow(declared: string, actual: string): Row<object> {
  const throws = { name: 'RangeError', reason: 'out of range', message: declared };
  return {
    method: 'Fake.call',
    throws,
    call: () => {
      throw new RangeError(actual);
    },
  };
}

describe('throws row with a declared message', () => {
  test('fails when both engines throw a different message', async () => {
    const row = throwingRow('index out of range', 'cell not found');
    await expect(runRow(row, fakeEngine, fakeEngine)).rejects.toThrow();
  });

  test('passes when both engines throw a message containing it', async () => {
    const row = throwingRow('index out of range', 'row index out of range: 7');
    await expect(runRow(row, fakeEngine, fakeEngine)).resolves.toBeUndefined();
  });
});
