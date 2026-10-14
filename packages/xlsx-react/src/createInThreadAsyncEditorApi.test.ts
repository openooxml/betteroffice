import { expect, mock, test } from 'bun:test';
import type { XlsxEditorApi } from './XlsxEditor';
import { createXlsxCommandController, XlsxCommandAdmissionError } from './commands/createXlsxCommandStore';
import { createInThreadAsyncEditorApi } from './createInThreadAsyncEditorApi';
import { XlsxAsyncApiSyncAccessError } from './worker/asyncApiErrors';

test('fallback exports wait for pending input and keep sync access unavailable', async () => {
  let release!: () => void;
  const input = new Promise<void>((resolve) => { release = resolve; });
  const exported = { ok: true as const, version: 'v1',
    content: { markdown: 'accepted cells', truncated: false, anchors: [], diagnostics: [] } };
  const exportMarkdown = mock(() => exported);
  const sync = {
    commands: createXlsxCommandController().store,
    flushPendingInput: () => input,
    handle: { exportMarkdown },
  } as unknown as XlsxEditorApi;
  const api = createInThreadAsyncEditorApi(sync, false, () => {});
  const result = api.exportMarkdown();
  await Promise.resolve();
  expect(exportMarkdown).not.toHaveBeenCalled();
  release();
  expect(await result).toBe(exported);
  expect(() => Reflect.get(api, 'handle')).toThrow(XlsxAsyncApiSyncAccessError);
  expect(() => Reflect.get(api, 'save')).toThrow(XlsxAsyncApiSyncAccessError);
  expect(api.collaboration).toBeNull();
});

test('fallback exports stop when input admission rejects a retired document', async () => {
  const failure = new XlsxCommandAdmissionError('document-replaced');
  const exportMarkdown = mock(() => ({}));
  const sync = {
    commands: createXlsxCommandController().store,
    flushPendingInput: async () => { throw failure; },
    handle: { exportMarkdown },
  } as unknown as XlsxEditorApi;
  const api = createInThreadAsyncEditorApi(sync, false, () => { throw failure; });
  await expect(api.exportMarkdown()).rejects.toBe(failure);
  expect(exportMarkdown).not.toHaveBeenCalled();
});
