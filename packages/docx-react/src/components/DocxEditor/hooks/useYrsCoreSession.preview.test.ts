import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { saveYrsDocx, type YrsSession } from '@betteroffice/docx/yrs';
import { useYrsCoreSession } from './useYrsCoreSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');

const PAGES = new Uint8Array(
  readFileSync(
    resolve(
      import.meta.dir,
      '../../../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx'
    )
  )
);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);

// Its sections have columns, which a preview refuses.
const COLUMNS = new Uint8Array(
  readFileSync(
    resolve(
      import.meta.dir,
      '../../../../../../crates/betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx'
    )
  )
);

afterAll(async () => {
  cleanup();
  if (ownsDom) await GlobalRegistrator.unregister();
});

test.each([false, true])(
  'a first-page preview opens first and hands over after paint with workerOpen=%s',
  async (workerOpen) => {
    const hosts: Array<boolean> = [];
    const mainOpens: boolean[] = [];
    let workerOpens = 0;
    const openInWorker = async () => {
      workerOpens += 1;
      return null;
    };
    const { result, unmount } = renderHook(() =>
      useYrsCoreSession(
        true,
        null,
        null,
        PAGES,
        1,
        undefined,
        {
          isCurrentLoad: () => true,
          onSession: (session) => {
            const open = session.openDocx.bind(session);
            session.openDocx = (bytes, seed, options) => {
              mainOpens.push(seed);
              return open(bytes, seed, options);
            };
          },
          onHostDocument: (_host, _generation, _session, options) => hosts.push(options?.preview === true),
        },
        {
          previewFirstPage: true,
          workerOpen: workerOpen ? { openInWorker, renderedFrame: null } : undefined,
        }
      )
    );
    await waitFor(() => expect(result.current.previewing).toBe(true));
    const preview = result.current.session!;
    expect(result.current.opening).toBe(true);
    expect(hosts).toEqual([true]);
    expect(preview.materializeDocx()).toBeNull();
    await expect(saveYrsDocx(preview)).rejects.toThrow();
    expect(preview.isDisplayOnly()).toBe(true);
    expect(result.current.documentFromYrs(null)).toBeNull();
    const paragraph = preview.paragraphs('body')[0]!;
    expect(() =>
      preview.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'x')
    ).toThrow(/display-only/);

    // The worker opens the full document while the preview paints.
    await waitFor(() => expect(workerOpens).toBe(workerOpen ? 1 : 0));
    expect(mainOpens).toEqual([]);
    const requestFrame = globalThis.requestAnimationFrame;
    const frames: FrameRequestCallback[] = [];
    globalThis.requestAnimationFrame = (callback) => frames.push(callback);
    try {
      await act(async () => {
        result.current.notifyFramePresented(preview);
      });
      expect(workerOpens).toBe(workerOpen ? 1 : 0);
      expect(mainOpens).toEqual([]);
      expect(frames).toHaveLength(1);
      act(() => frames.shift()!(performance.now()));
      expect(workerOpens).toBe(workerOpen ? 1 : 0);
      expect(mainOpens).toEqual([]);
      act(() => frames.shift()!(performance.now()));
    } finally {
      globalThis.requestAnimationFrame = requestFrame;
    }
    await waitFor(() => expect(result.current.previewing).toBe(false));
    const full = result.current.session as YrsSession;
    expect(full).not.toBe(preview);
    expect(hosts).toEqual([true, false]);
    expect(result.current.handoffFrom).toBe(preview);
    // Still showing the preview's pages: nothing edits until the full session's are shown.
    expect(result.current.opening).toBe(true);
    expect(full.materializeDocx()).not.toBeNull();
    expect(full.isDisplayOnly()).toBe(false);
    expect(mainOpens).toEqual([true]);
    expect(workerOpens).toBe(workerOpen ? 1 : 0);
    expect(result.current.replicaReady).toBe(true);

    await act(async () => {
      result.current.notifyFramePresented(full);
    });
    expect(result.current.handoffFrom).toBeNull();
    expect(result.current.opening).toBe(false);
    unmount();
  }
);

test.each(['heldEngine', 'shownEngine'] as const)(
  "the renderer's %s keeps the preview alive past the handoff until it lets go",
  async (holder) => {
    const { result, rerender, unmount } = renderHook(
      ({ held }: { held: unknown }) =>
        useYrsCoreSession(
          true,
          null,
          null,
          PAGES,
          1,
          undefined,
          { isCurrentLoad: () => true },
          { previewFirstPage: true, [holder]: held }
        ),
      { initialProps: { held: null as unknown } }
    );
    await waitFor(() => expect(result.current.previewing).toBe(true));
    const preview = result.current.session!;
    let destroyed = false;
    const destroy = preview.destroy.bind(preview);
    preview.destroy = () => {
      destroyed = true;
      destroy();
    };
    rerender({ held: preview });
    await act(async () => {
      result.current.notifyFramePresented(preview);
    });
    await waitFor(() => expect(result.current.previewing).toBe(false));
    const full = result.current.session!;
    await act(async () => {
      result.current.notifyFramePresented(full);
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    expect(result.current.handoffFrom).toBeNull();
    expect(destroyed).toBe(false);
    expect(preview.paragraphs('body').length).toBeGreaterThan(0);

    rerender({ held: full });
    expect(destroyed).toBe(true);
    unmount();
  }
);

test('a tab that draws no frames still opens the full document', async () => {
  const requestFrame = globalThis.requestAnimationFrame;
  globalThis.requestAnimationFrame = () => 0;
  try {
    const { result, unmount } = renderHook(() =>
      useYrsCoreSession(
        true,
        null,
        null,
        PAGES,
        1,
        undefined,
        { isCurrentLoad: () => true },
        { previewFirstPage: true }
      )
    );
    await waitFor(() => expect(result.current.previewing).toBe(true));
    const preview = result.current.session!;
    await act(async () => {
      result.current.notifyFramePresented(preview);
    });
    await waitFor(() => expect(result.current.previewing).toBe(false));
    expect(result.current.session).not.toBe(preview);
    unmount();
  } finally {
    globalThis.requestAnimationFrame = requestFrame;
  }
});

test('the full open waits two frames past the painted preview', async () => {
  const requestFrame = globalThis.requestAnimationFrame;
  const frames: FrameRequestCallback[] = [];
  globalThis.requestAnimationFrame = (callback) => frames.push(callback);
  const runFrame = () => frames.shift()!(performance.now());
  try {
    const { result, unmount } = renderHook(() =>
      useYrsCoreSession(
        true,
        null,
        null,
        PAGES,
        1,
        undefined,
        { isCurrentLoad: () => true },
        { previewFirstPage: true }
      )
    );
    await waitFor(() => expect(result.current.previewing).toBe(true));
    const preview = result.current.session!;
    await act(async () => {
      result.current.notifyFramePresented(preview);
    });
    await waitFor(() => expect(frames).toHaveLength(1));
    runFrame();
    expect(frames).toHaveLength(1);
    expect(result.current.session).toBe(preview);
    runFrame();
    await waitFor(() => expect(result.current.previewing).toBe(false));
    expect(result.current.session).not.toBe(preview);
    unmount();
  } finally {
    globalThis.requestAnimationFrame = requestFrame;
  }
});

test('without the option the document opens in full at once', async () => {
  const hosts: Array<boolean> = [];
  const { result, unmount } = renderHook(() =>
    useYrsCoreSession(true, null, null, PAGES, 1, undefined, {
      isCurrentLoad: () => true,
      onHostDocument: (_host, _generation, _session, options) => hosts.push(options?.preview === true),
    })
  );
  await waitFor(() => expect(result.current.session).not.toBeNull());
  expect(result.current.previewing).toBe(false);
  expect(hosts).toEqual([false]);
  unmount();
});

test.each([false, true])(
  'a document the preview refuses opens in full on the main thread with workerOpen=%s',
  async (workerOpen) => {
    const hosts: Array<boolean> = [];
    let workerOpens = 0;
    const openInWorker = async () => {
      workerOpens += 1;
      return null;
    };
    const { result, unmount } = renderHook(() =>
      useYrsCoreSession(
        true,
        null,
        null,
        COLUMNS,
        1,
        undefined,
        {
          isCurrentLoad: () => true,
          onHostDocument: (_host, _generation, _session, options) => hosts.push(options?.preview === true),
        },
        {
          previewFirstPage: true,
          workerOpen: workerOpen ? { openInWorker, renderedFrame: null } : undefined,
        }
      )
    );
    await waitFor(() => expect(result.current.session).not.toBeNull());
    expect(result.current.previewing).toBe(false);
    expect(result.current.session!.isDisplayOnly()).toBe(false);
    expect(hosts).toEqual([false]);
    expect(workerOpens).toBe(workerOpen ? 1 : 0);
    unmount();
  }
);

test('changing the preview option keeps the open session', async () => {
  const { result, rerender, unmount } = renderHook(
    ({ preview }: { preview: boolean }) =>
      useYrsCoreSession(
        true,
        null,
        null,
        PAGES,
        1,
        undefined,
        { isCurrentLoad: () => true },
        { previewFirstPage: preview }
      ),
    { initialProps: { preview: false } }
  );
  await waitFor(() => expect(result.current.session).not.toBeNull());
  const session = result.current.session;
  rerender({ preview: true });
  await act(async () => {});
  expect(result.current.session).toBe(session);
  unmount();
});

test('collaboration attached during a preview gets only the full session as its replica', async () => {
  const replicas: unknown[] = [];
  const onReplica = (replica: unknown) => {
    replicas.push(replica);
  };
  const { result, rerender, unmount } = renderHook(
    ({ collaborate }: { collaborate: boolean }) =>
      useYrsCoreSession(
        true,
        null,
        null,
        PAGES,
        1,
        collaborate ? { onReplica } : undefined,
        { isCurrentLoad: () => true },
        { previewFirstPage: true }
      ),
    { initialProps: { collaborate: false } }
  );
  await waitFor(() => expect(result.current.previewing).toBe(true));
  const preview = result.current.session!;
  rerender({ collaborate: true });
  await act(async () => {});
  expect(replicas).toEqual([]);
  await act(async () => {
    result.current.notifyFramePresented(preview);
  });
  await waitFor(() => expect(result.current.previewing).toBe(false));
  const full = result.current.session;
  expect(full).not.toBe(preview);
  expect(replicas).toEqual([full]);
  unmount();
});
