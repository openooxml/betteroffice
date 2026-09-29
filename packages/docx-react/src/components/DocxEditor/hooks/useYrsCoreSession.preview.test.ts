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

test('a first-page preview opens first, cannot save, and hands over once it has painted', async () => {
  const hosts: Array<boolean> = [];
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
        onHostDocument: (_host, _generation, options) => hosts.push(options?.preview === true),
      },
      { previewFirstPage: true }
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

  await act(async () => {
    result.current.notifyFramePresented(preview);
  });
  await waitFor(() => expect(result.current.previewing).toBe(false));
  const full = result.current.session as YrsSession;
  expect(full).not.toBe(preview);
  expect(hosts).toEqual([true, false]);
  expect(result.current.handoffFrom).toBe(preview);
  // Still showing the preview's pages: nothing edits until the full session's are shown.
  expect(result.current.opening).toBe(true);
  expect(full.materializeDocx()).not.toBeNull();
  expect(full.isDisplayOnly()).toBe(false);

  await act(async () => {
    result.current.notifyFramePresented(full);
  });
  expect(result.current.handoffFrom).toBeNull();
  expect(result.current.opening).toBe(false);
  unmount();
});

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

test('without the option the document opens in full at once', async () => {
  const hosts: Array<boolean> = [];
  const { result, unmount } = renderHook(() =>
    useYrsCoreSession(true, null, null, PAGES, 1, undefined, {
      isCurrentLoad: () => true,
      onHostDocument: (_host, _generation, options) => hosts.push(options?.preview === true),
    })
  );
  await waitFor(() => expect(result.current.session).not.toBeNull());
  expect(result.current.previewing).toBe(false);
  expect(hosts).toEqual([false]);
  unmount();
});

test('a document the preview refuses opens in full at once', async () => {
  const hosts: Array<boolean> = [];
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
        onHostDocument: (_host, _generation, options) => hosts.push(options?.preview === true),
      },
      { previewFirstPage: true }
    )
  );
  await waitFor(() => expect(result.current.session).not.toBeNull());
  expect(result.current.previewing).toBe(false);
  expect(result.current.session!.isDisplayOnly()).toBe(false);
  expect(hosts).toEqual([false]);
  unmount();
});

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
