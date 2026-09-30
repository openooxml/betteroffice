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

function updateSource() {
  const listeners = new Set<() => void>();
  return {
    onUpdate: (listener: () => void) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    update: () => {
      for (const listener of [...listeners]) listener();
    },
  };
}

/** Every family the loader asks for, in order, after `act` on an accepted seeded document. */
async function fontsLoadedFor(
  unusedScriptFonts: string[] | undefined,
  act_: (
    accept: (session?: ReturnType<typeof updateSource>) => void,
    asked: string[]
  ) => void | Promise<void> = (accept) => accept(updateSource())
): Promise<string[]> {
  const fontScope = createFontLoadScope();
  const asked: string[] = [];
  fontScope.loadFontsWithMapping = async (families) => {
    asked.push(...new Set(families.map((family) => family.trim())));
  };
  const options = loaderOptions(fontScope);
  const { result } = renderHook(() => useDocumentLoader(options));
  act(() => {
    void result.current.loadBuffer(new ArrayBuffer(4));
  });
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
  const generation = result.current.yrsSeedGeneration;
  await act_(
    (session) =>
      act(() => result.current.acceptHostDocument(host as never, generation, session as never)),
    asked
  );
  await waitFor(() => expect(asked).toContain('Calibri'));
  fontScope.dispose();
  return asked;
}

const loadCount = (asked: string[], family: string) =>
  asked.filter((name) => name === family).length;

test('a seeded document loads none of the fonts it names only for script text it lacks', async () => {
  const skipped = ['Batang', '바탕'];
  expect(new Set(await fontsLoadedFor(skipped))).toEqual(new Set(['Calibri']));
  expect(new Set(await fontsLoadedFor(undefined))).toEqual(new Set(['Batang', 'Calibri', '바탕']));
  expect(new Set(await fontsLoadedFor(skipped, (accept) => accept()))).toEqual(
    new Set(['Batang', 'Calibri', '바탕'])
  );
});

test('the fonts skipped at open load once, after the document first changes', async () => {
  const skipped = ['Batang', '바탕'];
  const afterEdits = await fontsLoadedFor(skipped, async (accept, asked) => {
    const session = updateSource();
    accept(session);
    await waitFor(() => expect(asked).toContain('Calibri'));
    expect(asked).not.toContain('Batang');
    expect(asked).not.toContain('바탕');
    act(() => {
      session.update();
      session.update();
    });
    await waitFor(() => expect(asked).toContain('바탕'));
  });
  expect(loadCount(afterEdits, 'Batang')).toBe(1);
  expect(loadCount(afterEdits, '바탕')).toBe(1);

  const editedBeforeOpenLoads = await fontsLoadedFor(skipped, async (accept, asked) => {
    const session = updateSource();
    accept(session);
    act(() => session.update());
    expect(asked).toEqual([]);
    await waitFor(() => expect(asked).toContain('바탕'));
  });
  expect(loadCount(editedBeforeOpenLoads, 'Batang')).toBe(1);
  expect(loadCount(editedBeforeOpenLoads, '바탕')).toBe(1);
  expect(editedBeforeOpenLoads.indexOf('Batang')).toBeGreaterThan(
    editedBeforeOpenLoads.indexOf('Calibri')
  );

  const replaced = await fontsLoadedFor(skipped, async (accept, asked) => {
    accept(updateSource());
    accept(updateSource());
    await waitFor(() => expect(asked).toContain('바탕'));
  });
  expect(loadCount(replaced, 'Batang')).toBe(1);
  expect(loadCount(replaced, '바탕')).toBe(1);
});
