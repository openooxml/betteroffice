import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');

// The second session a load creates is the full one, after its preview.
const real = await import('@betteroffice/docx/yrs');
// Mocking rebinds the module's live exports, `real`'s included.
const { createYrsSession } = real;
let created = 0;
let fullSession: unknown = null;
let shownPages = false;
let fullOpen: 'fail' | 'open' | 'layout-fail' = 'fail';
let holdFullOpen: Promise<void> | null = null;
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: async (options: Parameters<typeof createYrsSession>[0]) => {
    created += 1;
    if (created === 2 && holdFullOpen) await holdFullOpen;
    if (created === 2 && fullOpen === 'fail') {
      // Fails once the preview shows, however long a loaded machine takes to paint it.
      for (let waited = 0; !shownPages && waited < 20_000; waited += 20) {
        await new Promise((done) => setTimeout(done, 20));
      }
      throw new Error('full open failed');
    }
    const session = await createYrsSession(options);
    if (created === 2) fullSession = session;
    if (created === 2 && fullOpen === 'layout-fail') {
      session.layoutFontRequirementsJson = () => {
        throw new Error('layout failed');
      };
    }
    return session;
  },
}));
const displayList = await import('./hooks/useDisplayList');
const { useCanvasRenderer } = displayList;
let renderer: ReturnType<typeof useCanvasRenderer> | null = null;
// Once the full session exists, its pages fail to render; or, until the full
// session's first frame shows, the preview's error stays set.
const renderFailure = new Error('render failed');
const previewFailure = new Error('preview render failed');
let failRender: 'full' | 'preview' | null = null;
mock.module('./hooks/useDisplayList', () => ({
  ...displayList,
  useCanvasRenderer: (...args: Parameters<typeof useCanvasRenderer>) => {
    renderer = useCanvasRenderer(...args);
    if (renderer.displayList) shownPages = true;
    const error =
      failRender === 'full' && created >= 2
        ? renderFailure
        : failRender === 'preview' && (created < 2 || renderer.presentedEngine !== fullSession)
          ? previewFailure
          : null;
    return error ? { ...renderer, error, status: 'error' as const } : renderer;
  },
}));
const { DocxEditor } = await import('../../index');
type Editor = import('../../index').DocxEditorRef;

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const PAGES = resolve(
  import.meta.dir,
  '../../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx'
);
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
  // React's scheduler still runs the last commit's passive effects, which read `window`.
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const documentBuffer = () => {
  const bytes = readFileSync(PAGES);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};
const load = (buffer: ArrayBuffer, onError: (error: Error) => void, ref = createRef<Editor>()) => (
  <DocxEditor ref={ref} previewFirstPage documentBuffer={buffer} onError={onError} />
);

test('a load whose full open fails after its preview painted keeps none of its pages', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'fail';
  failRender = null;
  const bytes = readFileSync(PAGES);
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const view = render(
    <DocxEditor
      ref={ref}
      previewFirstPage
      documentBuffer={
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
      }
      onError={(error) => errors.push(error.message)}
    />
  );
  await waitFor(() => expect(shownPages).toBe(true), { timeout: 10_000 });
  await waitFor(() => expect(errors).toEqual(['full open failed']), { timeout: 10_000 });
  await act(async () => {});
  expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
  expect(renderer!.displayList).toBeNull();
  expect(renderer!.presentedEngine).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
}, 30_000);

test('a load whose full session fails to render fails, and leaves no session behind', async () => {
  created = 0;
  fullOpen = 'open';
  failRender = 'full';
  const bytes = readFileSync(PAGES);
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const view = render(
    <DocxEditor
      ref={ref}
      previewFirstPage
      documentBuffer={
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
      }
      onError={(error) => errors.push(error.message)}
    />
  );
  await waitFor(() => expect(errors).toEqual(['render failed']), { timeout: 10_000 });
  await act(async () => {});
  expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
  expect(renderer!.displayList).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
}, 30_000);

test('a load whose full session fails to lay out fails once, and leaves no session behind', async () => {
  created = 0;
  fullOpen = 'layout-fail';
  failRender = null;
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const view = render(load(documentBuffer(), (error) => errors.push(error.message), ref));
  await waitFor(() => expect(errors).toContain('layout failed'), { timeout: 10_000 });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(errors).toEqual(['layout failed']);
  expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
}, 30_000);

test("a preview's render error is reported once and does not fail the full session", async () => {
  created = 0;
  fullOpen = 'open';
  failRender = 'preview';
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const buffer = documentBuffer();
  const view = render(load(buffer, (error) => errors.push(error.message), ref));
  await waitFor(() => expect(created).toBe(2), { timeout: 10_000 });
  // A host passing a new callback while the full session opens.
  view.rerender(load(buffer, (error) => errors.push(`again: ${error.message}`), ref));
  await waitFor(() => expect(ref.current!.getDocument()).not.toBeNull(), { timeout: 10_000 });
  expect(errors).toEqual(['preview render failed']);
  expect(view.container.querySelector('.docx-editor-error')).toBeNull();
}, 30_000);

test('while its preview shows, a load has no page count and its layout is not complete', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  let release = () => {};
  holdFullOpen = new Promise((done) => (release = done));
  try {
    const ref = createRef<Editor>();
    render(load(documentBuffer(), () => {}, ref));
    await waitFor(() => expect(ref.current).not.toBeNull());
    let pages = null as number | null;
    void ref.current!.whenLayoutComplete().then((count) => (pages = count));
    await waitFor(() => expect(shownPages).toBe(true), { timeout: 10_000 });
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(ref.current!.getTotalPages()).toBe(0);
    expect(pages).toBeNull();

    release();
    await waitFor(() => expect(pages).not.toBeNull(), { timeout: 20_000 });
    expect(renderer!.presentedEngine).toBe(fullSession);
    expect(pages).toBe(renderer!.displayList!.pages.length);
    expect(ref.current!.getTotalPages()).toBe(pages!);
  } finally {
    release();
    holdFullOpen = null;
  }
}, 30_000);
