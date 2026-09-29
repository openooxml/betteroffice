import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook, waitFor } = await import('@testing-library/react');
const { createFontLoadScope, isFontLoaded, registerDocumentFaces } = await import(
  '@betteroffice/docx/utils'
);
const { useDocumentLoader } = await import('./useDocumentLoader');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test("a parsed document's fonts load after the previous document's faces are released", async () => {
  const fontScope = createFontLoadScope();
  await registerDocumentFaces(
    [{ family: 'Previous Document Face', data: new Uint8Array([1, 2, 3]).buffer }],
    fontScope
  );
  expect(isFontLoaded('Previous Document Face')).toBe(true);
  const loadedAtLoad: boolean[] = [];
  fontScope.loadDocumentFonts = async () => {
    loadedAtLoad.push(isFontLoaded('Previous Document Face'));
  };
  const noop = () => {};
  const options = {
    documentBuffer: null,
    initialDocument: { package: { document: { content: [] } } } as never,
    externalContent: false,
    history: { reset: noop, state: null } as never,
    pagedEditorRef: { current: null },
    setLoadingState: noop,
    setComments: noop,
    setShowCommentsSidebar: noop,
    onError: undefined,
    resetForNewDocument: noop,
    commentsLoadedRef: { current: false },
    commentIdAllocator: { seedAbove: noop } as never,
    setDocumentFonts: noop,
    fontScope,
  };
  renderHook(() => useDocumentLoader(options));
  await waitFor(() => expect(loadedAtLoad).toEqual([false]));
  fontScope.dispose();
});
