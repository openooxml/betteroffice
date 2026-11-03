import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { expect, test } from 'bun:test';

const { XlsxDocument } = createRequire(import.meta.url)('./index.js') as typeof import('./index.js');
const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url));
const failedObject = 'This engine object failed in an earlier call and can only be freed';

function thrown(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected an error');
}

test('a panic fails the raw xlsx object until it is freed', () => {
  const document = XlsxDocument.open(fixture);
  try {
    const panic = thrown(() => document.panicForTest());
    expect(panic).toBeInstanceOf(Error);
    expect(panic).toHaveProperty('code', 'GenericFailure');

    const failure = thrown(() => document.sheetInfoJson());
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe(failedObject);
    expect(failure).not.toHaveProperty('code');
  } finally {
    expect(() => document.free()).not.toThrow();
  }
});
