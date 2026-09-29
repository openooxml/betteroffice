import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document } from '@betteroffice/docx/types/document';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { useHistory } from '../../../hooks/useHistory';
import { createCommentIdAllocator } from '../commentFactories';
import { useDocumentLoader } from './useDocumentLoader';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const ROOT = resolve(import.meta.dir, '../../../../../..');
const COMMENTED = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/structured-export/principal.docx'))
);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

afterAll(async () => {
  cleanup();
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('comments load from the full document, not from its preview', async () => {
  const previewSession = await createYrsSession();
  const fullSession = await createYrsSession();
  const preview = previewSession.openDocxPreview(COMMENTED, 1);
  const full = fullSession.openDocx(COMMENTED, true);
  const fullComments = full.document.package.document?.comments;
  expect(fullComments?.length).toBeGreaterThan(0);

  const loaded: Comment[][] = [];
  const sidebar: boolean[] = [];
  const { result, unmount } = renderHook(() =>
    useDocumentLoader({
      documentBuffer: null,
      initialDocument: null,
      externalContent: false,
      history: useHistory<Document | null>(null),
      pagedEditorRef: { current: null },
      setLoadingState: () => {},
      setComments: (next) => {
        if (typeof next !== 'function') loaded.push(next);
      },
      setShowCommentsSidebar: (next) => {
        if (typeof next !== 'function') sidebar.push(next);
      },
      onError: undefined,
      resetForNewDocument: () => {},
      commentsLoadedRef: { current: false },
      commentIdAllocator: createCommentIdAllocator(),
      setDocumentFonts: () => {},
    })
  );
  await act(async () => {
    void result.current.loadBuffer(COMMENTED.slice().buffer);
  });
  const generation = result.current.yrsSeedGeneration;

  await act(async () => {
    result.current.acceptHostDocument(preview, generation, { preview: true });
  });
  expect(loaded).toEqual([]);
  expect(sidebar).toEqual([true]);

  await act(async () => {
    result.current.acceptHostDocument(full, generation);
  });
  expect(loaded).toHaveLength(1);
  expect(loaded[0]).toBe(fullComments!);
  unmount();
  previewSession.destroy();
  fullSession.destroy();
});
