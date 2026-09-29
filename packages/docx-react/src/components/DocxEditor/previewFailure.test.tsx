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
let fullOpen: 'fail' | 'open' = 'fail';
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: (options: Parameters<typeof createYrsSession>[0]) => {
    created += 1;
    if (created === 2 && fullOpen === 'fail') return Promise.reject(new Error('full open failed'));
    return createYrsSession(options);
  },
}));
const displayList = await import('./hooks/useDisplayList');
const { useCanvasRenderer } = displayList;
let renderer: ReturnType<typeof useCanvasRenderer> | null = null;
// Once the full session exists, its pages fail to render.
const renderFailure = new Error('render failed');
let failFullRender = false;
mock.module('./hooks/useDisplayList', () => ({
  ...displayList,
  useCanvasRenderer: (...args: Parameters<typeof useCanvasRenderer>) => {
    renderer = useCanvasRenderer(...args);
    return failFullRender && created >= 2
      ? { ...renderer, error: renderFailure, status: 'error' as const }
      : renderer;
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
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a load whose full open fails after its preview painted keeps none of its pages', async () => {
  created = 0;
  fullOpen = 'fail';
  failFullRender = false;
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
  await waitFor(() => expect(renderer?.displayList).not.toBeNull(), { timeout: 10_000 });
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
  failFullRender = true;
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
