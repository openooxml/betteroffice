import { expect, mock, test } from 'bun:test';
import { XlsxSaveRefusedError, type XlsxEditorApi } from './XlsxEditor';
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
  expect(Reflect.get(api, 'save')).toBeFunction();
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

for (const method of ['version', 'readCells', 'findText', 'validateEdits', 'applyEdits', 'save'] as const) {
  for (const outcome of ['accepted', 'input-failed', 'document-replaced'] as const) {
    test(`fallback host ${method} handles ${outcome} pending input`, async () => {
      let release!: () => void;
      const pending = new Promise<void>((resolve) => { release = resolve; });
      const events: string[] = [];
      const bytes = new Uint8Array([1, 2]);
      const operation = mock(async () => { await pending; events.push('draft', method); return 'v1'; });
      const saved = mock((_bytes: Uint8Array) => {});
      const sync = { handle: {}, commands: createXlsxCommandController().store,
        flushPendingInput: async () => {
          await pending;
          if (outcome === 'input-failed') throw new XlsxCommandAdmissionError(outcome);
          events.push('draft');
        },
        [method]: method === 'save' ? () => { events.push('save'); return bytes; } : operation,
      } as unknown as XlsxEditorApi;
      const api = createInThreadAsyncEditorApi(sync, false, () => {
        if (outcome !== 'accepted') throw new XlsxCommandAdmissionError(outcome);
      }, () => saved);
      const work = Reflect.get(api, method)(...(['version', 'save'].includes(method) ? [] : [{ steps: [] }]));
      expect(events).toEqual([]);
      release();
      if (outcome === 'accepted') {
        await work;
        expect(events).toEqual(['draft', method]);
        if (method === 'save') {
          expect(await work).toBe(bytes);
          expect(saved).toHaveBeenCalledExactlyOnceWith(bytes);
        }
      } else await expect(work).rejects.toMatchObject({ code: outcome });
    });
  }
}

for (const method of ['validateEdits', 'applyEdits'] as const) {
  test(`fallback host ${method} preserves read-only refusal`, async () => {
    const refusal = { ok: false, version: 'v1', failure: { code: 'read-only' } };
    const sync = { handle: {}, commands: createXlsxCommandController().store,
      [method]: async () => refusal } as unknown as XlsxEditorApi;
    const api = createInThreadAsyncEditorApi(sync, false, () => {});
    await expect(Reflect.get(api, method)({ steps: [] })).resolves.toBe(refusal);
  });
}

for (const code of ['input-pending', 'input-failed'] as const) {
  test(`fallback save handles synchronous ${code}`, async () => {
    const bytes = new Uint8Array([4]);
    const save = mock(() => bytes);
    save.mockImplementationOnce(() => { throw new XlsxSaveRefusedError(code); });
    const flushPendingInput = mock(async () => {});
    const sync = { handle: {}, commands: createXlsxCommandController().store, save, flushPendingInput } as unknown as XlsxEditorApi;
    const saved = mock((_bytes: Uint8Array) => {});
    const api = createInThreadAsyncEditorApi(sync, false, () => {}, () => saved);
    if (code === 'input-pending') {
      await expect(api.save()).resolves.toBe(bytes);
      expect(flushPendingInput).toHaveBeenCalledTimes(2);
      expect(saved).toHaveBeenCalledExactlyOnceWith(bytes);
    } else {
      await expect(api.save()).rejects.toMatchObject({ code: 'input-failed' });
      expect(saved).not.toHaveBeenCalled();
    }
  });
}
