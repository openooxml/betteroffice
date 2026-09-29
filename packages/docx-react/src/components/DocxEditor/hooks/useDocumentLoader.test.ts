import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const { createFontLoadScope, isFontLoaded, registerDocumentFaces } = await import(
  '@betteroffice/docx/utils'
);
const { useDocumentLoader } = await import('./useDocumentLoader');

afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const noop = () => {};
const loaderOptions = (fontScope: ReturnType<typeof createFontLoadScope>) => ({
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
  const options = loaderOptions(fontScope);
  renderHook(() => useDocumentLoader(options));
  await waitFor(() => expect(loadedAtLoad).toEqual([false]));
  fontScope.dispose();
});

async function fontsLoadedFor(unusedScriptFonts?: string[]): Promise<Set<string>> {
  const fontScope = createFontLoadScope();
  const asked = new Set<string>();
  fontScope.loadFontsWithMapping = async (families) => {
    for (const family of families) asked.add(family);
  };
  const options = loaderOptions(fontScope);
  const { result } = renderHook(() => useDocumentLoader(options));
  void result.current.loadBuffer(new ArrayBuffer(4));
  await waitFor(() => expect(result.current.yrsSeedBytes).not.toBeNull());
  const document = {
    package: {
      document: { content: [] },
      styles: { styles: [] },
      fontTable: {
        fonts: [
          { name: 'Calibri', panose1: '020F0502020204030204', charset: '00' },
          { name: 'Batang', altName: '바탕', charset: '81' },
        ],
      },
    },
  };
  const host = {
    document,
    referencedFonts: ['Batang', 'Calibri', '바탕'],
    embeddedFonts: new Map(),
    ...(unusedScriptFonts ? { unusedScriptFonts } : {}),
  };
  act(() => result.current.acceptHostDocument(host as never, result.current.yrsSeedGeneration));
  await waitFor(() => expect(asked.has('Calibri')).toBe(true));
  fontScope.dispose();
  return asked;
}

test('a seeded document loads none of the fonts it names only for script text it lacks', async () => {
  expect(await fontsLoadedFor(['Batang', '바탕'])).toEqual(new Set(['Calibri']));
  expect(await fontsLoadedFor()).toEqual(new Set(['Batang', 'Calibri', '바탕']));
});
