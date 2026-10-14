import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document } from '@betteroffice/docx/types/document';
import { createFontLoadScope, extractFontsFromDocument } from '@betteroffice/docx/utils';
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
  const preview = previewSession.openDocxPreview(COMMENTED, 1)!;
  const full = fullSession.openDocx(COMMENTED, true);
  const fullComments = full.document.package.document?.comments;
  expect(fullComments?.length).toBeGreaterThan(0);

  const fontScope = createFontLoadScope();
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
      fontScope,
    })
  );
  await act(async () => {
    void result.current.loadBuffer(COMMENTED.slice().buffer);
  });
  const generation = result.current.yrsSeedGeneration;

  await act(async () => {
    result.current.acceptHostDocument(preview, generation, undefined, { preview: true });
  });
  expect(loaded).toEqual([]);
  expect(sidebar).toEqual([true]);

  await act(async () => {
    result.current.acceptHostDocument(full, generation);
  });
  expect(loaded).toHaveLength(1);
  expect(loaded[0]).toBe(fullComments!);
  unmount();
  fontScope.dispose();
  previewSession.destroy();
  fullSession.destroy();
});

test('a load whose full open fails keeps nothing of its preview', async () => {
  const previewSession = await createYrsSession();
  const preview = previewSession.openDocxPreview(COMMENTED, 1)!;
  const fontScope = createFontLoadScope();
  const errors: Error[] = [];
  const { result, unmount } = renderHook(() => {
    const history = useHistory<Document | null>(null);
    const loader = useDocumentLoader({
      documentBuffer: null,
      initialDocument: null,
      externalContent: false,
      history,
      pagedEditorRef: { current: null },
      setLoadingState: () => {},
      setComments: () => {},
      setShowCommentsSidebar: () => {},
      onError: (error) => errors.push(error),
      resetForNewDocument: () => {},
      commentsLoadedRef: { current: false },
      commentIdAllocator: createCommentIdAllocator(),
      setDocumentFonts: () => {},
      fontScope,
    });
    return { history, loader };
  });
  await act(async () => {
    void result.current.loader.loadBuffer(COMMENTED.slice().buffer);
  });
  const generation = result.current.loader.yrsSeedGeneration;
  await act(async () => {
    result.current.loader.acceptHostDocument(preview, generation, undefined, { preview: true });
  });
  expect(result.current.history.state).toBe(preview.document);

  await act(async () => {
    result.current.loader.failHostDocument(new Error('full open failed'), generation);
  });
  expect(result.current.history.state).toBeNull();
  expect(errors.map((error) => error.message)).toEqual(['full open failed']);
  unmount();
  fontScope.dispose();
  previewSession.destroy();
});

test("a preview's font loads stop once its load's full document is accepted", async () => {
  const previewSession = await createYrsSession();
  const fullSession = await createYrsSession();
  const preview = previewSession.openDocxPreview(COMMENTED, 1)!;
  const full = fullSession.openDocx(COMMENTED, true);
  const fontScope = createFontLoadScope();
  const fontLoads: string[][] = [];
  fontScope.loadFontsWithMapping = async (families) => {
    fontLoads.push(families);
  };
  const { result, unmount } = renderHook(() =>
    useDocumentLoader({
      documentBuffer: null,
      initialDocument: null,
      externalContent: false,
      history: useHistory<Document | null>(null),
      pagedEditorRef: { current: null },
      setLoadingState: () => {},
      setComments: () => {},
      setShowCommentsSidebar: () => {},
      onError: undefined,
      resetForNewDocument: () => {},
      commentsLoadedRef: { current: false },
      commentIdAllocator: createCommentIdAllocator(),
      setDocumentFonts: () => {},
      fontScope,
    })
  );
  await act(async () => {
    void result.current.loadBuffer(COMMENTED.slice().buffer);
  });
  const generation = result.current.yrsSeedGeneration;
  await act(async () => {
    result.current.acceptHostDocument(preview, generation, undefined, { preview: true });
    result.current.acceptHostDocument(full, generation);
    await new Promise((done) => setTimeout(done, 50));
  });
  expect(fontLoads).toEqual([full.referencedFonts, [...extractFontsFromDocument(full.document)]]);
  unmount();
  fontScope.dispose();
  previewSession.destroy();
  fullSession.destroy();
});
