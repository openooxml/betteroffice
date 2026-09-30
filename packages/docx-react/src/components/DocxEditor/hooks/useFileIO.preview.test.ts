import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { PagedEditorRef } from '../PagedEditor';
import { useFileIO } from './useFileIO';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('the editor refuses to save a first-page preview', async () => {
  let projected = 0;
  const session = { isDisplayOnly: () => true };
  const editor = {
    getYrsSession: () => session,
    flushPendingInput: async () => {},
    getDocument: () => {
      projected += 1;
      return null;
    },
  } as unknown as PagedEditorRef;
  const saved: ArrayBuffer[] = [];
  const errors: Error[] = [];
  const { result } = renderHook(() =>
    useFileIO({
      pagedEditorRef: { current: editor },
      resolveImage: () => null,
      comments: [],
      documentName: undefined,
      onSave: (buffer: ArrayBuffer) => saved.push(buffer),
      onOpen: undefined,
      onError: (error: Error) => errors.push(error),
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    } as unknown as Parameters<typeof useFileIO>[0])
  );
  let buffer: ArrayBuffer | null = new ArrayBuffer(1);
  await act(async () => {
    buffer = await result.current.handleSave();
  });
  expect(buffer).toBeNull();
  expect(saved).toEqual([]);
  expect(projected).toBe(0);
  expect(errors.map((error) => error.message).join()).toContain('still opening');
});
