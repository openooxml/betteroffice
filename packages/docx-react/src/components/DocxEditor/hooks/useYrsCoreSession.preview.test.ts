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
  expect(hosts).toEqual([true]);
  expect(preview.materializeDocx()).toBeNull();
  await expect(saveYrsDocx(preview)).rejects.toThrow();

  await act(async () => {
    result.current.notifyFramePresented(preview);
  });
  await waitFor(() => expect(result.current.previewing).toBe(false));
  const full = result.current.session as YrsSession;
  expect(full).not.toBe(preview);
  expect(hosts).toEqual([true, false]);
  expect(result.current.handoffFrom).toBe(preview);
  expect(full.materializeDocx()).not.toBeNull();

  await act(async () => {
    result.current.notifyFramePresented(full);
  });
  expect(result.current.handoffFrom).toBeNull();
  unmount();
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
