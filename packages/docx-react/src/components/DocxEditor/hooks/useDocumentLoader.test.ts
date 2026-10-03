import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { FontOption } from '@betteroffice/docx/utils/fontOptions';
import type { YrsDocxHost } from '@betteroffice/docx/yrs';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const { createFontLoadScope, isFontLoaded, registerDocumentFaces } = await import(
  '@betteroffice/docx/utils'
);
const { PICKER_FONTS_FALLBACK_MS, useDocumentLoader } = await import('./useDocumentLoader');
const fontLoader = await import('@betteroffice/docx/utils/fontLoader');
const { resolveFontFamily } = await import('@betteroffice/docx/utils/fontResolver');
const restorePicker: Array<() => void> = [];

afterEach(() => {
  cleanup();
  for (const restore of restorePicker.splice(0)) restore();
});
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
    accept: (session?: ReturnType<typeof updateSource>, options?: { preview: boolean }) => void,
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
    (session, acceptOptions) =>
      act(() =>
        result.current.acceptHostDocument(host as never, generation, session as never, acceptOptions)
      ),
    asked
  );
  await waitFor(() => expect(asked).toContain('Calibri'));
  fontScope.dispose();
  return asked;
}

const loadCount = (asked: string[], family: string) =>
  asked.filter((name) => name === family).length;

function pickerLoader() {
  let nextFrame = 0;
  const frames = new Map<number, FrameRequestCallback>();
  const requestFrame = spyOn(globalThis, 'requestAnimationFrame').mockImplementation((callback) => {
    const id = ++nextFrame;
    frames.set(id, callback);
    return id;
  });
  const cancelFrame = spyOn(globalThis, 'cancelAnimationFrame').mockImplementation((id) => {
    frames.delete(id);
  });
  const probe = spyOn(fontLoader, 'canRenderFont').mockImplementation(
    (family) => family !== 'Picker Missing'
  );
  const fontScope = createFontLoadScope();
  const loadFonts = mock(async (_families: string[]) => {});
  fontScope.loadFontsWithMapping = loadFonts;
  const fonts = mock((_fonts: FontOption[]) => {});
  const hook = renderHook(() => useDocumentLoader({
    ...loaderOptions(fontScope),
    initialDocument: null,
    setDocumentFonts: fonts,
  }));
  const load = () => {
    act(() => { void hook.result.current.loadBuffer(new ArrayBuffer(4)); });
    return hook.result.current.yrsSeedGeneration;
  };
  const frame = () => act(() => {
    const pending = [...frames];
    for (const [id, callback] of pending) {
      if (!frames.delete(id)) continue;
      callback(performance.now());
    }
  });
  const task = () => act(async () => {
    await new Promise((done) => setTimeout(done, 0));
  });
  restorePicker.push(() => {
    requestFrame.mockRestore();
    cancelFrame.mockRestore();
    probe.mockRestore();
    fontScope.dispose();
  });
  return { hook, load, frame, task, probe, fonts, loadFonts, frames };
}

function pickerHost(family: string): YrsDocxHost {
  return {
    document: {
      package: {
        document: { content: [] },
        theme: { fontScheme: { majorFont: { latin: family } } },
      },
    } as YrsDocxHost['document'],
    referencedFonts: [family],
    embeddedFonts: new Map(),
  };
}

test('host acceptance prepares layout fonts immediately and probes picker fonts after presentation', async () => {
  const h = pickerLoader();
  const generation = h.load();
  const session = updateSource();
  const host = pickerHost('Picker Available');
  host.document.package.theme!.fontScheme!.minorFont = { latin: 'Picker Missing' };
  host.document.package.fontTable = {
    fonts: [{ name: 'Picker Embedded', embedRegular: { relId: 'rId1' } }],
  };
  host.referencedFonts = [
    'picker available', 'Picker Embedded', 'Picker Missing', 'Picker Reference', 'sans-serif',
  ];

  act(() => h.hook.result.current.acceptHostDocument(host, generation, session as never));
  expect(h.probe).not.toHaveBeenCalled();
  expect(h.fonts).not.toHaveBeenCalled();
  await waitFor(() => expect(h.loadFonts).toHaveBeenCalledTimes(2));
  expect(h.loadFonts.mock.calls).toEqual([
    [host.referencedFonts],
    [['Picker Embedded']],
  ]);
  expect(h.probe).not.toHaveBeenCalled();

  act(() => h.hook.result.current.notifyDocumentFramePresented(updateSource()));
  expect(h.frames.size).toBe(0);
  act(() => {
    h.hook.result.current.notifyDocumentFramePresented(session);
    h.hook.result.current.notifyDocumentFramePresented(session);
  });
  expect(h.frames.size).toBe(1);
  expect(h.probe).not.toHaveBeenCalled();
  h.frame();
  await h.task();
  expect(h.probe).not.toHaveBeenCalled();
  h.frame();
  expect(h.probe).not.toHaveBeenCalled();
  await h.task();

  expect(h.fonts.mock.calls).toEqual([
    [['picker available', 'Picker Embedded', 'Picker Reference'].map((name) => ({
      name,
      fontFamily: resolveFontFamily(name).cssFallback,
      category: 'other',
    }))],
  ]);
  expect(h.probe).not.toHaveBeenCalledWith('Picker Embedded');
  act(() => h.hook.result.current.notifyDocumentFramePresented(session));
  expect(h.frames.size).toBe(0);
});

test('a new load cancels picker frames and timers, and publishing clears the fallback', async () => {
  const realSetTimeout = globalThis.setTimeout;
  const fallbacks: Array<() => void> = [];
  const tasks: Array<ReturnType<typeof setTimeout>> = [];
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((
    callback: () => void,
    ms?: number
  ) => {
    if (ms === PICKER_FONTS_FALLBACK_MS) {
      fallbacks.push(callback);
      return 0;
    }
    const id = realSetTimeout(callback, ms);
    if (ms === 0) tasks.push(id);
    return id;
  }) as unknown as typeof setTimeout);
  const clearTimer = spyOn(globalThis, 'clearTimeout');
  restorePicker.push(() => {
    timer.mockRestore();
    clearTimer.mockRestore();
  });
  const h = pickerLoader();
  let generation = h.load();
  for (const frameCount of [1, 2]) {
    const session = updateSource();
    act(() => {
      h.hook.result.current.acceptHostDocument(
        pickerHost('Picker Cancelled'), generation, session as never
      );
      h.hook.result.current.notifyDocumentFramePresented(session);
    });
    for (let frame = 0; frame < frameCount; frame++) h.frame();
    expect(h.frames.size).toBe(frameCount === 1 ? 1 : 0);
    const taskTimer = tasks.at(-1);
    clearTimer.mockClear();

    generation = h.load();
    expect(h.frames.size).toBe(0);
    expect(clearTimer).toHaveBeenCalledWith(0);
    if (frameCount === 2) expect(clearTimer).toHaveBeenCalledWith(taskTimer);
    await h.task();
    expect(h.probe).not.toHaveBeenCalled();
    expect(h.fonts).not.toHaveBeenCalled();
  }

  clearTimer.mockClear();
  const session = updateSource();
  act(() => {
    h.hook.result.current.acceptHostDocument(
      pickerHost('Picker Published'), generation, session as never
    );
    h.hook.result.current.notifyDocumentFramePresented(session);
  });
  h.frame();
  h.frame();
  await h.task();
  expect(h.fonts.mock.calls.map(([fonts]) => fonts.map((font) => font.name))).toEqual([
    ['Picker Published'],
  ]);
  expect(clearTimer).toHaveBeenCalledWith(0);
  expect(fallbacks).toHaveLength(3);
  act(() => fallbacks[2]!());
  expect(h.fonts).toHaveBeenCalledTimes(1);
});

test('a picker task queued for a replaced load neither probes nor publishes its fonts', async () => {
  const h = pickerLoader();
  const first = h.load();
  const firstSession = updateSource();
  act(() => {
    h.hook.result.current.acceptHostDocument(pickerHost('Picker Old'), first, firstSession as never);
    h.hook.result.current.notifyDocumentFramePresented(firstSession);
  });
  h.frame();
  h.frame();

  const next = h.load();
  const nextSession = updateSource();
  act(() =>
    h.hook.result.current.acceptHostDocument(pickerHost('Picker New'), next, nextSession as never)
  );
  await h.task();
  expect(h.probe).not.toHaveBeenCalled();
  expect(h.fonts).not.toHaveBeenCalled();

  act(() => h.hook.result.current.notifyDocumentFramePresented(firstSession));
  expect(h.frames.size).toBe(0);
  act(() => h.hook.result.current.notifyDocumentFramePresented(nextSession));
  h.frame();
  h.frame();
  await h.task();
  expect(h.fonts.mock.calls[0]![0].map((font) => font.name)).toEqual(['Picker New']);
  expect(h.probe).not.toHaveBeenCalledWith('Picker Old');
});

test("a full host replaces its preview's pending picker discovery within the same load", async () => {
  const h = pickerLoader();
  const generation = h.load();
  const previewSession = updateSource();
  act(() => {
    h.hook.result.current.acceptHostDocument(
      pickerHost('Picker Preview'), generation, previewSession as never, { preview: true }
    );
    h.hook.result.current.notifyDocumentFramePresented(previewSession);
  });
  h.frame();
  h.frame();

  const fullSession = updateSource();
  act(() =>
    h.hook.result.current.acceptHostDocument(pickerHost('Picker Full'), generation, fullSession as never)
  );
  await h.task();
  expect(h.probe).not.toHaveBeenCalled();
  expect(h.fonts).not.toHaveBeenCalled();
  act(() => h.hook.result.current.notifyDocumentFramePresented(fullSession));
  h.frame();
  h.frame();
  await h.task();
  expect(h.fonts.mock.calls[0]![0].map((font) => font.name)).toEqual(['Picker Full']);
  expect(h.probe).not.toHaveBeenCalledWith('Picker Preview');
});

test('picker fonts still arrive once when no frame is presented', () => {
  const realSetTimeout = globalThis.setTimeout;
  const fallbacks: Array<() => void> = [];
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((
    callback: () => void,
    ms?: number
  ) => {
    if (ms !== PICKER_FONTS_FALLBACK_MS) return realSetTimeout(callback, ms);
    fallbacks.push(callback);
    return 0;
  }) as unknown as typeof setTimeout);
  const clearTimer = spyOn(globalThis, 'clearTimeout');
  restorePicker.push(() => {
    timer.mockRestore();
    clearTimer.mockRestore();
  });
  const h = pickerLoader();
  const generation = h.load();
  act(() =>
    h.hook.result.current.acceptHostDocument(
      pickerHost('Picker Unshown'), generation, updateSource() as never
    )
  );
  expect(h.probe).not.toHaveBeenCalled();
  expect(fallbacks).toHaveLength(1);

  act(() => fallbacks[0]!());
  expect(h.fonts.mock.calls.map(([fonts]) => fonts.map((font) => font.name))).toEqual([
    ['Picker Unshown'],
  ]);
  expect(clearTimer).toHaveBeenCalledWith(0);
  act(() => fallbacks[0]!());
  expect(h.fonts).toHaveBeenCalledTimes(1);
});

test('a seeded document loads none of the fonts it names only for script text it lacks', async () => {
  const skipped = ['Batang', '바탕'];
  expect(new Set(await fontsLoadedFor(skipped))).toEqual(new Set(['Calibri']));
  expect(new Set(await fontsLoadedFor(undefined))).toEqual(new Set(['Batang', 'Calibri', '바탕']));
  expect(new Set(await fontsLoadedFor(skipped, (accept) => accept()))).toEqual(
    new Set(['Batang', 'Calibri', '바탕'])
  );
});

test("a preview loads none of the fonts its pages name only for script text they lack", async () => {
  const skipped = ['Batang', '바탕'];
  const preview = { preview: true };
  expect(new Set(await fontsLoadedFor(skipped, (accept) => accept(undefined, preview)))).toEqual(
    new Set(['Calibri'])
  );
  expect(new Set(await fontsLoadedFor(undefined, (accept) => accept(undefined, preview)))).toEqual(
    new Set(['Batang', 'Calibri', '바탕'])
  );
  // Nothing watches a preview's session for the skipped fonts.
  const touched = await fontsLoadedFor(skipped, async (accept, asked) => {
    const session = updateSource();
    accept(session, preview);
    await waitFor(() => expect(asked).toContain('Calibri'));
    act(() => session.update());
    await new Promise((done) => setTimeout(done, 50));
  });
  expect(touched).not.toContain('Batang');
  expect(touched).not.toContain('바탕');
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
