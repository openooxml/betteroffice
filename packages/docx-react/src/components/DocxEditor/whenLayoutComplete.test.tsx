import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { pagedDocx } from './__fixtures__/pagedDocx';
import { setupWorkerEngine } from './__fixtures__/workerEngine';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { parseDocx } from '@betteroffice/docx/docx';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const { act, cleanup, render } = await import('@testing-library/react');
const { DocxEditor } = await import('../../index');
type DocxEditorRef = import('../../index').DocxEditorRef;

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function tick(ms = 10) {
  await act(async () => {
    await new Promise((done) => setTimeout(done, ms));
  });
}

/** Lets the editor work until `done()`, for at most `ms` on a busy machine. */
async function until(done: () => boolean, ms = 20_000) {
  const deadline = performance.now() + ms;
  while (!done() && performance.now() < deadline) await tick();
}

/** The page count `whenLayoutComplete()` resolves with, or its failure. */
async function layoutComplete(ref: React.RefObject<DocxEditorRef | null>) {
  let outcome = null as number | Error | null;
  ref.current!.whenLayoutComplete().then(
    (pages) => (outcome = pages),
    (error: Error) => (outcome = error)
  );
  await until(() => outcome !== null);
  return outcome;
}

async function mountTwoPages() {
  const ref = createRef<DocxEditorRef>();
  render(<DocxEditor ref={ref} documentBuffer={await pagedDocx(2)} />);
  await until(() => ref.current !== null);
  expect(await layoutComplete(ref)).toBe(2);
  return ref;
}

test('a parsed reload settles with the new document', async () => {
  const ref = await mountTwoPages();
  const document = await parseDocx(await pagedDocx(3));
  act(() => ref.current!.loadDocument(document));
  expect(await layoutComplete(ref)).toBe(3);
  expect(ref.current!.getTotalPages()).toBe(3);
}, 90_000);

test('a parsed reload with the same page count reports it again', async () => {
  const ref = await mountTwoPages();
  const document = ref.current!.getDocument()!;
  act(() => ref.current!.loadDocument(document));
  expect(await layoutComplete(ref)).toBe(2);
  expect(ref.current!.getTotalPages()).toBe(2);
}, 90_000);

test('a reload waits for the new document and reports no pages while it loads', async () => {
  const ref = await mountTwoPages();

  let release = () => {};
  const gate = new Promise<void>((done) => (release = done));
  const bytes = await pagedDocx(3);
  const blob = new Blob([bytes]);
  blob.arrayBuffer = async () => {
    await gate;
    return bytes;
  };
  let loading: Promise<void> = Promise.resolve();
  await act(async () => {
    loading = ref.current!.loadDocumentBuffer(blob);
  });
  let next = null as number | null;
  void ref.current!.whenLayoutComplete().then((pages) => (next = pages));
  await tick(200);
  expect(next).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);

  release();
  await until(() => next !== null);
  await loading;
  expect(next).toBe(3);
  expect(ref.current!.getTotalPages()).toBe(3);
}, 90_000);

test('each failed load rejects the wait', async () => {
  const ref = await mountTwoPages();
  const detached = await pagedDocx(1);
  structuredClone(detached, { transfer: [detached] });
  for (let load = 0; load < 2; load += 1) {
    await act(async () => {
      await ref.current!.loadDocumentBuffer(detached);
    });
    expect(await layoutComplete(ref)).toBeInstanceOf(Error);
  }
}, 90_000);

test('a failed load before any layout rejects the wait', async () => {
  const ref = createRef<DocxEditorRef>();
  render(<DocxEditor ref={ref} />);
  await until(() => ref.current !== null);
  const detached = await pagedDocx(1);
  structuredClone(detached, { transfer: [detached] });
  await act(async () => {
    await ref.current!.loadDocumentBuffer(detached);
  });
  await tick(200);
  expect(await layoutComplete(ref)).toBeInstanceOf(Error);
}, 90_000);

test('a failure without a message rejects waits during and after the load', async () => {
  const ref = await mountTwoPages();
  let fail = () => {};
  const blob = new Blob([]);
  blob.arrayBuffer = () => new Promise((_, reject) => (fail = () => reject(new Error())));
  await act(async () => {
    void ref.current!.loadDocumentBuffer(blob);
  });
  const during = layoutComplete(ref);
  fail();
  expect(await during).toBeInstanceOf(Error);
  expect(await layoutComplete(ref)).toBeInstanceOf(Error);
}, 90_000);

describe('DocxEditor layout waits (worker engine)', () => {
  const workers = setupWorkerEngine();

  test('a failed load before any layout rejects the wait on the worker engine', async () => {
    const ref = createRef<DocxEditorRef>();
    render(<DocxEditor ref={ref} experimentalWorkerOpen />);
    await until(() => ref.current !== null);
    const detached = await pagedDocx(1);
    structuredClone(detached, { transfer: [detached] });
    await act(async () => {
      await ref.current!.loadDocumentBuffer(detached);
    });
    await tick(200);
    expect(await layoutComplete(ref)).toBeInstanceOf(Error);

    const bytes = await pagedDocx(1);
    let loading: Promise<void> = Promise.resolve();
    await act(async () => {
      loading = ref.current!.loadDocumentBuffer(bytes);
    });
    expect(await layoutComplete(ref)).toBe(1);
    await loading;
    expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);
  }, 90_000);
});
