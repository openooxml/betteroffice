import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');

// The second session a load creates is the full one, after its preview.
const real = await import('@betteroffice/docx/yrs');
// Mocking rebinds the module's live exports, `real`'s included.
const { createYrsSession } = real;
let fullOpen: 'open' | 'fail' | 'stall' = 'open';
let created = 0;
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: (options: Parameters<typeof createYrsSession>[0]) => {
    created += 1;
    if (created % 2 === 0 && fullOpen === 'fail') {
      return Promise.reject(new Error('full open failed'));
    }
    if (created % 2 === 0 && fullOpen === 'stall') return new Promise(() => {});
    return createYrsSession(options);
  },
}));
const { useYrsCoreSession } = await import('./useYrsCoreSession');

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

/** A previewing load, past the preview's paint, with the full open as `mode` says. */
async function previewThen(mode: typeof fullOpen, fullOpenTimeoutMs = 10_000) {
  fullOpen = mode;
  created = 0;
  const errors: Error[] = [];
  const hook = renderHook(() =>
    useYrsCoreSession(
      true,
      null,
      null,
      PAGES,
      1,
      undefined,
      { isCurrentLoad: () => true, onError: (error) => errors.push(error) },
      { previewFirstPage: true, fullOpenTimeoutMs }
    )
  );
  await waitFor(() => expect(hook.result.current.previewing).toBe(true));
  const preview = hook.result.current.session!;
  await act(async () => {
    hook.result.current.notifyFramePresented(preview);
  });
  return { ...hook, errors, preview };
}

test('a failed full open reports the error and leaves no preview behind', async () => {
  const { result, errors, unmount } = await previewThen('fail');
  await waitFor(() => expect(errors.map((error) => error.message)).toEqual(['full open failed']));
  expect(result.current.session).toBeNull();
  expect(result.current.previewing).toBe(false);
  expect(result.current.opening).toBe(false);
  unmount();
});

test('a full open that never finishes fails the load once its wait runs out', async () => {
  const { result, errors, unmount } = await previewThen('stall', 50);
  await waitFor(() => expect(errors).toHaveLength(1));
  expect(result.current.session).toBeNull();
  expect(result.current.opening).toBe(false);
  unmount();
});

test('a full session whose first frame never shows takes over once the wait runs out', async () => {
  const { result, errors, preview, unmount } = await previewThen('open', 50);
  await waitFor(() => expect(result.current.previewing).toBe(false));
  const full = result.current.session;
  expect(full).not.toBe(preview);
  await waitFor(() => expect(result.current.opening).toBe(false));
  expect(result.current.handoffFrom).toBeNull();
  expect(result.current.session).toBe(full);
  expect(errors).toEqual([]);
  unmount();
});
