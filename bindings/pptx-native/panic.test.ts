import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { expect, test } from 'bun:test';

const { PptxDocument, PptxRenderer } = createRequire(import.meta.url)('./index.js') as typeof import('./index.js');
const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/betteroffice-demo.pptx', import.meta.url));
const failedObject = 'This engine object failed in an earlier call and can only be freed';

/** Raw `renderPngTarget` with the scale written as-is, so NaN and Infinity reach the binding as bad JSON. */
function slidePng(renderer: { renderPngTarget(document: unknown, target: string, options: string): Uint8Array }, document: unknown, slide: number, scale: number) {
  return renderer.renderPngTarget(document, `{"slide":${slide}}`, `{"scale":${scale}}`);
}

function thrown(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected an error');
}

test.skipIf(process.env.BETTEROFFICE_NATIVE_TEST_HOOKS !== '1')('a panic fails the raw pptx object until it is freed', () => {
  const document = PptxDocument.openCollaborative(fixture, 1);
  try {
    const panic = thrown(() => (document as unknown as { panicForTest(): void }).panicForTest());
    expect(panic).toBeInstanceOf(Error);
    expect(panic).toHaveProperty('code', 'GenericFailure');

    const failure = thrown(() => document.snapshotJson());
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe(failedObject);
    expect(failure).not.toHaveProperty('code');

    expect((thrown(() => document.validateProposalJson('{}')) as Error).message).toBe(failedObject);
    expect((thrown(() => document.proposalSlideStripJson('proposal')) as Error).message).toBe(failedObject);
    const renderer = new PptxRenderer();
    try {
      expect((thrown(() => renderer.layoutProposalDiffSlideByIdJson(document, 'proposal', 'slide')) as Error).message).toBe(failedObject);
    } finally {
      renderer.free();
    }
  } finally {
    expect(() => document.free()).not.toThrow();
  }
});

test.skipIf(process.env.BETTEROFFICE_NATIVE_TEST_HOOKS !== '1')('a panic inside renderPngTarget poisons the raw pptx renderer and document', () => {
  const document = PptxDocument.openCollaborative(fixture, 2);
  const renderer = new PptxRenderer();
  const hooks = renderer as unknown as { panicNextPngForTest(): void };
  try {
    hooks.panicNextPngForTest();
    const panic = thrown(() => slidePng(renderer, document, 0, 1));
    expect(panic).toBeInstanceOf(Error);
    expect(panic).toHaveProperty('code', 'GenericFailure');
    for (const call of [
      () => renderer.hitTestJson(0, 0),
      () => document.snapshotJson(),
      () => slidePng(renderer, document, 0, 1),
    ]) {
      const failure = thrown(call);
      expect((failure as Error).message).toBe(failedObject);
      expect(failure).not.toHaveProperty('code');
    }
  } finally {
    expect(() => renderer.free()).not.toThrow();
    expect(() => document.free()).not.toThrow();
  }
});
