import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { expect, test } from 'bun:test';

const { PptxDocument } = createRequire(import.meta.url)('./index.js') as typeof import('./index.js');
const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/betteroffice-demo.pptx', import.meta.url));
const failedObject = 'This engine object failed in an earlier call and can only be freed';

function thrown(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected an error');
}

test('a panic fails the raw pptx object until it is freed', () => {
  const document = PptxDocument.openCollaborative(fixture, 1);
  try {
    const panic = thrown(() => document.panicForTest());
    expect(panic).toBeInstanceOf(Error);
    expect(panic).toHaveProperty('code', 'GenericFailure');

    const failure = thrown(() => document.snapshotJson());
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe(failedObject);
    expect(failure).not.toHaveProperty('code');
  } finally {
    expect(() => document.free()).not.toThrow();
  }
});
