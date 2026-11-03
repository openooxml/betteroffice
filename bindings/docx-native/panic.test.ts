import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { expect, test } from 'bun:test';

const { EditSession } = createRequire(import.meta.url)('./index.js') as typeof import('./index.js');
const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/betteroffice-demo.docx', import.meta.url));
const failedObject = 'This engine object failed in an earlier call and can only be freed';

function thrown(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected an error');
}

test.skipIf(process.env.BETTEROFFICE_NATIVE_TEST_HOOKS !== '1')('a panic fails the raw docx object until it is freed', () => {
  const document = new EditSession(1);
  try {
    document.open_docx(fixture, true);
    const panic = thrown(() => (document as unknown as { panicForTest(): void }).panicForTest());
    expect(panic).toBeInstanceOf(Error);
    expect(panic).toHaveProperty('code', 'GenericFailure');

    const failure = thrown(() => document.version());
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe(failedObject);
    expect(failure).not.toHaveProperty('code');
  } finally {
    expect(() => document.free()).not.toThrow();
  }
});

test.skipIf(process.env.BETTEROFFICE_NATIVE_TEST_HOOKS !== '1')('a panic mid-edit delivers no queued update and fails the raw docx object', () => {
  const document = new EditSession(2);
  const hooks = document as unknown as { panicAfterInsertForTest(story: string, paraId: string, text: string): void };
  try {
    document.open_docx(fixture, true);
    const paragraphId = (JSON.parse(document.read_paragraphs_json(JSON.stringify({ story: 'body', view: 'accepted' }))) as { paragraphs: { paraId: string }[] }).paragraphs[0]!.paraId;
    let delivered = 0;
    document.set_update_observer(() => { delivered += 1; });
    document.insert_text('body', paragraphId, 0, 'Delivered ');
    expect(delivered).toBeGreaterThan(0);
    delivered = 0;
    expect(thrown(() => hooks.panicAfterInsertForTest('body', paragraphId, 'Half applied '))).toHaveProperty('code', 'GenericFailure');
    expect(delivered).toBe(0);
    expect((thrown(() => document.version()) as Error).message).toBe(failedObject);
  } finally {
    expect(() => document.free()).not.toThrow();
  }
});

for (const swallowed of [false, true]) {
  test.skipIf(process.env.BETTEROFFICE_NATIVE_TEST_HOOKS !== '1')(`a panic in an edit made by an update observer stops delivery and fails the outer call (${swallowed ? 'observer swallows it' : 'observer lets it through'})`, () => {
    const document = new EditSession(3);
    const hooks = document as unknown as { panicAfterInsertForTest(story: string, paraId: string, text: string): void };
    try {
      document.open_docx(fixture, true);
      const paragraphId = (JSON.parse(document.read_paragraphs_json(JSON.stringify({ story: 'body', view: 'accepted' }))) as { paragraphs: { paraId: string }[] }).paragraphs[0]!.paraId;
      let panicked = false;
      let afterPanic = 0;
      document.set_update_observer(() => {
        if (panicked) { afterPanic += 1; return; }
        panicked = true;
        if (!swallowed) { hooks.panicAfterInsertForTest('body', paragraphId, 'Nested '); return; }
        try { hooks.panicAfterInsertForTest('body', paragraphId, 'Nested '); } catch {}
      });
      const outer = thrown(() => document.insert_text('body', paragraphId, 0, 'Outer '));
      expect(panicked).toBe(true);
      expect(afterPanic).toBe(0);
      expect(outer).toBeInstanceOf(Error);
      if (swallowed) expect((outer as Error).message).toBe(failedObject);
      else expect(outer).toHaveProperty('code', 'GenericFailure');
      expect((thrown(() => document.version()) as Error).message).toBe(failedObject);
      expect(afterPanic).toBe(0);
    } finally {
      expect(() => document.free()).not.toThrow();
    }
  });
}
