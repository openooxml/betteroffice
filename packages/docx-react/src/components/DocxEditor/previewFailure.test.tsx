import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';

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
let failPreviewLayout = false;
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
    if (created === 1 && failPreviewLayout) {
      session.layoutFontRequirementsJson = () => {
        throw new Error('preview layout failed');
      };
    }
    if (created === 2) fullSession = session;
    if (created === 2 && fullOpen === 'layout-fail') {
      session.layoutFontRequirementsJson = () => {
        throw new Error('layout failed');
      };
    }
    return session;
  },
}));
const layoutRender = await import('@betteroffice/docx/layout/render');
const { buildRustDisplayFrame } = layoutRender;
let holdFullDisplayList: Promise<void> | null = null;
mock.module('@betteroffice/docx/layout/render', () => ({
  ...layoutRender,
  buildRustDisplayFrame: async (...args: Parameters<typeof buildRustDisplayFrame>) => {
    const result = await buildRustDisplayFrame(...args);
    if (holdFullDisplayList && args[1] === fullSession) await holdFullDisplayList;
    return result;
  },
}));
const displayList = await import('./hooks/useDisplayList');
const { useCanvasRenderer } = displayList;
let renderer: ReturnType<typeof useCanvasRenderer> | null = null;
let holdCanvasReplay: 'full' | 'preview' | 'ordinary' | null = null;
interface PendingCanvasReplay {
  displayList: NonNullable<ReturnType<typeof useCanvasRenderer>['displayList']>;
  layoutEngine: unknown;
  isCurrent: () => boolean;
  resolve: () => void;
  reject: (error: Error) => void;
}
let pendingCanvasReplays: PendingCanvasReplay[] = [];
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
    if (error) return { ...renderer, error, status: 'error' as const };
    return holdCanvasReplay ? { ...renderer, offscreenReplay: null } : renderer;
  },
}));
const canvasReplay = await import('./canvasReplay');
const { presentCanvasReplay } = canvasReplay;
mock.module('./canvasReplay', () => ({
  ...canvasReplay,
  presentCanvasReplay: (...args: Parameters<typeof presentCanvasReplay>) => {
    const [preparations, isCurrent] = args;
    const shown = renderer;
    if (
      !holdCanvasReplay ||
      !shown?.displayList ||
      (holdCanvasReplay === 'full' &&
        (fullSession === null || shown.presentedEngine !== fullSession)) ||
      (holdCanvasReplay === 'preview' && shown.presentedEngine === fullSession)
    ) {
      return presentCanvasReplay(...args);
    }
    const list = shown.displayList;
    const ready = new Promise<void>((resolve, reject) => {
      pendingCanvasReplays.push({
        displayList: list,
        layoutEngine: shown.layoutEngine,
        isCurrent,
        resolve,
        reject,
      });
    });
    return presentCanvasReplay(
      [...preparations, { buffer: document.createElement('canvas'), ready, present: () => {} }],
      isCurrent
    );
  },
}));
const { isPresented, onReplayFailed } = await import('./internals/layoutProvenance');
const { DocxEditor } = await import('../../index');
type Editor = import('../../index').DocxEditorRef;

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const PAGES = resolve(
  import.meta.dir,
  '../../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx'
);
const quiet = { error: console.error, warn: console.warn };
const originalWorker = globalThis.Worker;

beforeAll(async () => {
  // Frames build on this thread, where the tests hold them; a resident worker's would not wait.
  globalThis.Worker = undefined as unknown as typeof Worker;
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
afterEach(() => {
  cleanup();
  holdCanvasReplay = null;
  for (const replay of pendingCanvasReplays) replay.resolve();
  pendingCanvasReplays = [];
});
afterAll(async () => {
  // React's scheduler still runs the last commit's passive effects, which read `window`.
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  globalThis.Worker = originalWorker;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const documentBuffer = () => {
  const bytes = readFileSync(PAGES);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};
const load = (buffer: ArrayBuffer, onError: (error: Error) => void, ref = createRef<Editor>()) => (
  <DocxEditor ref={ref} previewFirstPage documentBuffer={buffer} onError={onError} />
);

/** A wait begun after a load failed rejects with it. */
async function expectWaitRejects(ref: React.RefObject<Editor | null>) {
  let outcome = null as number | Error | null;
  ref.current!.whenLayoutComplete().then(
    (pages) => (outcome = pages),
    (error: Error) => (outcome = error)
  );
  await waitFor(() => expect(outcome).toBeInstanceOf(Error));
}

async function currentCanvasReplay() {
  await waitFor(() => expect(pendingCanvasReplays.some((replay) => replay.isCurrent())).toBe(true), {
    timeout: 20_000,
  });
  return pendingCanvasReplays.find((replay) => replay.isCurrent())!;
}

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
  await expectWaitRejects(ref);
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
  await expectWaitRejects(ref);
}, 30_000);

test('a full session whose canvas replay rejects during preview handover fails the load once', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  holdCanvasReplay = 'full';
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const replayError = new Error('full canvas replay failed');
  const view = render(load(documentBuffer(), (error) => errors.push(error), ref));
  const replay = await currentCanvasReplay();

  expect(created).toBe(2);
  expect(fullSession).not.toBeNull();
  expect(replay.layoutEngine).toBe(fullSession);
  expect(renderer!.presentedEngine).toBe(fullSession);
  expect(renderer!.displayList).toBe(replay.displayList);
  expect(isPresented(renderer!.canvasHostRef.current, replay.displayList)).toBe(false);
  expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
  expect(ref.current!.getDocument()).toBeNull();
  expect(errors).toEqual([]);

  await act(async () => {
    expect(replay.isCurrent()).toBe(true);
    replay.reject(replayError);
  });
  await waitFor(() => expect(errors).toEqual([replayError]), { timeout: 10_000 });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(errors).toEqual([replayError]);
  expect(view.container.querySelector('.docx-editor-error')?.textContent).toContain(
    replayError.message
  );
  expect(view.container.querySelector('.docx-editor-loading')).toBeNull();
  expect(view.queryByTestId('yrs-input')).toBeNull();
  expect(renderer!.displayList).toBeNull();
  expect(renderer!.presentedEngine).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
  await expectWaitRejects(ref);
}, 30_000);

test('a preview whose canvas replay rejects during full-session handover does not fail the load', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  holdCanvasReplay = 'preview';
  let releaseFullDisplayList = () => {};
  holdFullDisplayList = new Promise((done) => (releaseFullDisplayList = done));
  let offReplayFailed = () => {};
  try {
    const ref = createRef<Editor>();
    const errors: Error[] = [];
    const replayError = new Error('preview canvas replay failed');
    const view = render(load(documentBuffer(), (error) => errors.push(error), ref));
    await waitFor(
      () => {
        expect(fullSession).not.toBeNull();
        expect(renderer!.layoutEngine).toBe(fullSession);
      },
      { timeout: 20_000 }
    );
    const replay = await currentCanvasReplay();
    const previewEngine = renderer!.presentedEngine;

    expect(created).toBe(2);
    expect(previewEngine).not.toBeNull();
    expect(previewEngine).not.toBe(fullSession);
    expect((previewEngine as YrsSession).isDisplayOnly()).toBe(true);
    expect(renderer!.layoutEngine).toBe(fullSession);
    expect(renderer!.displayList).toBe(replay.displayList);
    expect(isPresented(renderer!.canvasHostRef.current, replay.displayList)).toBe(false);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();
    expect(errors).toEqual([]);

    const replayFailed = mock((_displayList: object, _error: unknown) => {});
    offReplayFailed = onReplayFailed(replayFailed);
    await act(async () => {
      expect(replay.isCurrent()).toBe(true);
      replay.reject(replayError);
    });
    await waitFor(() => expect(replayFailed).toHaveBeenCalledWith(replay.displayList, replayError));
    await act(async () => {});
    expect(errors).toEqual([]);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
    expect(renderer!.layoutEngine).toBe(fullSession);
    expect(renderer!.presentedEngine).toBe(previewEngine);
    expect(renderer!.displayList).toBe(replay.displayList);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();

    holdCanvasReplay = 'full';
    await act(async () => releaseFullDisplayList());
    await waitFor(() => expect(renderer!.presentedEngine).toBe(fullSession), { timeout: 10_000 });
    const fullReplay = await currentCanvasReplay();
    expect(fullReplay.layoutEngine).toBe(fullSession);
    expect(fullReplay.displayList).not.toBe(replay.displayList);
    expect(renderer!.displayList).toBe(fullReplay.displayList);
    expect(isPresented(renderer!.canvasHostRef.current, fullReplay.displayList)).toBe(false);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();

    await act(async () => {
      expect(fullReplay.isCurrent()).toBe(true);
      fullReplay.resolve();
    });
    await waitFor(() =>
      expect(isPresented(renderer!.canvasHostRef.current, fullReplay.displayList)).toBe(true)
    );
    await waitFor(() =>
      expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false)
    );
    expect(errors).toEqual([]);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
    expect(renderer!.layoutEngine).toBe(fullSession);
    expect(renderer!.presentedEngine).toBe(fullSession);
    expect(ref.current!.getEditorRef()!.getYrsSession()).toBe(fullSession as YrsSession);
    expect(ref.current!.getDocument()).not.toBeNull();
    expect(ref.current!.getTotalPages()).toBeGreaterThan(0);
    expect(created).toBe(2);
  } finally {
    offReplayFailed();
    releaseFullDisplayList();
    holdFullDisplayList = null;
  }
}, 30_000);

test('an ordinary editor logs a canvas replay rejection without failing the load', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  holdCanvasReplay = 'ordinary';
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const replayError = new Error('ordinary canvas replay failed');
  const previousError = console.error;
  const logged = mock((..._args: unknown[]) => {});
  console.error = logged;
  try {
    const view = render(
      <DocxEditor
        ref={ref}
        previewFirstPage={false}
        documentBuffer={documentBuffer()}
        onError={(error) => errors.push(error)}
      />
    );
    const replay = await currentCanvasReplay();
    expect(created).toBe(1);
    expect(renderer!.displayList).toBe(replay.displayList);
    expect(isPresented(renderer!.canvasHostRef.current, replay.displayList)).toBe(false);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false);
    expect(ref.current!.getDocument()).not.toBeNull();

    await act(async () => {
      expect(replay.isCurrent()).toBe(true);
      replay.reject(replayError);
    });
    await waitFor(() =>
      expect(logged).toHaveBeenCalledWith('[CanvasRenderer] Canvas replay failed', replayError)
    );
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(errors).toEqual([]);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false);
    expect(renderer!.status).toBe('ready');
    expect(renderer!.displayList).not.toBeNull();
    expect(ref.current!.getTotalPages()).toBeGreaterThan(0);
    expect(ref.current!.getDocument()).not.toBeNull();
  } finally {
    console.error = previousError;
  }
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
    let totalPages = null as number | null;
    void ref.current!.whenLayoutComplete().then((count) => {
      pages = count;
      totalPages = ref.current!.getTotalPages();
    });
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
    expect(totalPages).toBe(pages);
  } finally {
    release();
    holdFullOpen = null;
  }
}, 30_000);

test("a preview's layout error is reported and fails no wait for the document", async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  failPreviewLayout = true;
  try {
    const ref = createRef<Editor>();
    const errors: string[] = [];
    render(load(documentBuffer(), (error) => errors.push(error.message), ref));
    await waitFor(() => expect(ref.current).not.toBeNull());
    let outcome = null as number | Error | null;
    ref.current!.whenLayoutComplete().then(
      (pages) => (outcome = pages),
      (error: Error) => (outcome = error)
    );
    await waitFor(() => expect(errors).toContain('preview layout failed'), { timeout: 10_000 });
    await waitFor(() => expect(outcome).not.toBeNull(), { timeout: 20_000 });
    expect(renderer!.presentedEngine).toBe(fullSession);
    expect(outcome).toBe(renderer!.displayList!.pages.length);
  } finally {
    failPreviewLayout = false;
  }
}, 30_000);
