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

function throwingRow(declared: string | RegExp, actual: string): Row<object> {
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

  test('fails when the declared message is empty', async () => {
    const row = throwingRow('', 'cell not found');
    await expect(runRow(row, fakeEngine, fakeEngine)).rejects.toThrow();
  });

  test('fails when the declared pattern is empty', async () => {
    const row = throwingRow(new RegExp(''), 'cell not found');
    await expect(runRow(row, fakeEngine, fakeEngine)).rejects.toThrow();
  });

  test('fails when the declared pattern matches an empty message', async () => {
    const row = throwingRow(/x*/, 'cell not found');
    await expect(runRow(row, fakeEngine, fakeEngine)).rejects.toThrow();
  });

  test('a global pattern matching the message passes on every run', async () => {
    const row = throwingRow(/out of range/g, 'row index out of range: 7');
    await expect(runRow(row, fakeEngine, fakeEngine)).resolves.toBeUndefined();
    await expect(runRow(row, fakeEngine, fakeEngine)).resolves.toBeUndefined();
  });

  test('fails when a sticky pattern matches only past the start of the message', async () => {
    const row = throwingRow(/out of range/y, 'row index out of range: 7');
    await expect(runRow(row, fakeEngine, fakeEngine)).rejects.toThrow();
  });

  test('passes when a sticky pattern matches at the start of the message', async () => {
    const row = throwingRow(/row index/y, 'row index out of range: 7');
    await expect(runRow(row, fakeEngine, fakeEngine)).resolves.toBeUndefined();
  });
});

type Side = { side: 'native' | 'wasm' };

function sidedEngine(side: Side['side']) {
  return async (): Promise<EngineHandle<Side>> => ({ ...await fakeEngine(), target: { side } });
}

function versionShape(value: unknown): string {
  if (typeof value !== 'string' || !/^[0-9a-f]{4}-\d+$/.test(value)) throw new Error(`${JSON.stringify(value)} is not a version`);
  return value.slice(value.indexOf('-'));
}

function maskedRow(values: Record<Side['side'], { version: string; text: string }>): Row<Side> {
  return {
    method: 'Fake.open',
    masked: [{ path: 'state.version', shape: versionShape }],
    call: ({ side }) => ({ state: values[side] }),
  };
}

describe('row with a masked path', () => {
  test('passes when only the masked value differs', async () => {
    const row = maskedRow({ native: { version: 'aaaa-0', text: 'same' }, wasm: { version: 'bbbb-0', text: 'same' } });
    await expect(runRow(row, sidedEngine('wasm'), sidedEngine('native'))).resolves.toBeUndefined();
  });

  test('fails when a value outside the masked path differs', async () => {
    const row = maskedRow({ native: { version: 'aaaa-0', text: 'native' }, wasm: { version: 'bbbb-0', text: 'wasm' } });
    await expect(runRow(row, sidedEngine('wasm'), sidedEngine('native'))).rejects.toThrow();
  });

  test('fails when a masked value has the wrong shape', async () => {
    const row = maskedRow({ native: { version: 'aaaa-0', text: 'same' }, wasm: { version: 'not a version', text: 'same' } });
    await expect(runRow(row, sidedEngine('wasm'), sidedEngine('native'))).rejects.toThrow('not a version');
  });

  test('fails when the part the shape keeps differs', async () => {
    const row = maskedRow({ native: { version: 'aaaa-0', text: 'same' }, wasm: { version: 'bbbb-1', text: 'same' } });
    await expect(runRow(row, sidedEngine('wasm'), sidedEngine('native'))).rejects.toThrow();
  });
});
