import { expect, mock, test } from 'bun:test';
import { XlsxSaveRefusedError, type XlsxEditorApi } from './XlsxEditor';
import { createXlsxCommandController, XlsxCommandAdmissionError } from './commands/createXlsxCommandStore';
import { createInThreadAsyncEditorApi } from './createInThreadAsyncEditorApi';
import { XlsxAsyncApiSyncAccessError } from './worker/asyncApiErrors';

for (const method of ['readCells', 'findText', 'validateEdits', 'applyEdits'] as const) {
  test(`fallback host ${method} captures requests before delayed input`, async () => {
    let release!: () => void;
    const pending = new Promise<void>((resolve) => { release = resolve; });
    const request = { expectVersion: 'v1', query: 'original',
      ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }],
      steps: [{ op: 'setCellInputs', target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } },
        inputs: [['original']] }] };
    const original = structuredClone(request);
    const operation = mock(async (captured: typeof request) => {
      await pending;
      return structuredClone(captured);
    });
    const sync = { handle: {}, commands: createXlsxCommandController().store,
      [method]: operation } as unknown as XlsxEditorApi;
    const api = createInThreadAsyncEditorApi(sync, false, () => {});
    const work = Reflect.apply(api[method], api, [request]);
    request.expectVersion = 'v2';
    request.query = 'changed';
    request.ranges[0]!.range.a1 = 'B2';
    request.steps[0]!.target.range.a1 = 'B2';
    request.steps[0]!.inputs[0]![0] = 'changed';
    release();
    expect(await work).toEqual(original);
  });
}

for (const method of ['renderPng', 'renderRangePng', 'exportStructured', 'exportMarkdown', 'printDisplayList'] as const) {
  test(`fallback ${method} captures options before delayed input`, async () => {
    let release!: () => void;
    const pending = new Promise<void>((resolve) => { release = resolve; });
    const options = { sheet: 0, range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } };
    const markdownOptions = { includeEmptyCells: true };
    const args = method === 'exportMarkdown' ? [options, markdownOptions] : [options];
    const original = structuredClone(args);
    const operation = mock((...captured: unknown[]) => captured);
    const sync = { commands: createXlsxCommandController().store, flushPendingInput: () => pending,
      handle: { [method]: operation } } as unknown as XlsxEditorApi;
    const api = createInThreadAsyncEditorApi(sync, false, () => {});
    const work = Reflect.apply(api[method], api, args);
    expect(operation).not.toHaveBeenCalled();
    options.sheet = 1;
    options.range.start.row = 9;
    markdownOptions.includeEmptyCells = false;
    release();
    await work;
    expect(operation.mock.calls[0]).toEqual(original);
  });
}

for (const method of ['encodeStateAsUpdate', 'applyUpdate'] as const) {
  test(`fallback collaboration ${method} captures bytes before delayed input`, async () => {
    let release!: () => void;
    const pending = new Promise<void>((resolve) => { release = resolve; });
    const operation = mock((bytes: Uint8Array) => bytes.slice());
    const sync = { commands: createXlsxCommandController().store, flushPendingInput: () => pending,
      handle: { clientId: 1, [method]: operation } } as unknown as XlsxEditorApi;
    const api = createInThreadAsyncEditorApi(sync, true, () => {});
    const bytes = new Uint8Array([1, 2, 3]);
    const work = api.collaboration![method](bytes);
    expect(operation).not.toHaveBeenCalled();
    bytes.fill(0);
    release();
    await work;
    expect(operation.mock.calls[0]![0]).toEqual(new Uint8Array([1, 2, 3]));
  });
}

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
      let replaced = false;
      const operation = mock(async () => {
        await pending;
        if (outcome === 'input-failed') throw new XlsxCommandAdmissionError(outcome);
        events.push('draft', method);
        return 'v1';
      });
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
        if (replaced) throw new XlsxCommandAdmissionError('document-replaced');
      }, () => saved);
      const work = Reflect.apply(Reflect.get(api, method), api, ['version', 'save'].includes(method) ? [] : [{ steps: [] }]);
      expect(events).toEqual([]);
      if (outcome === 'document-replaced') replaced = true;
      release();
      if (outcome === 'accepted') {
        await work;
        expect(events).toEqual(['draft', method]);
        if (method === 'save') {
          expect(await work).toBe(bytes);
          expect(saved).toHaveBeenCalledTimes(1);
          expect(saved).toHaveBeenCalledWith(bytes);
        }
      } else await expect(work).rejects.toMatchObject({ code: outcome });
    });
  }
}

for (const method of ['validateEdits', 'applyEdits'] as const) {
  test(`fallback host ${method} preserves read-only refusal`, async () => {
    const refusal = { ok: false, version: 'v1', failure: { code: 'read-only', message: 'The editor is read-only' } } as const;
    const sync = { handle: {}, commands: createXlsxCommandController().store,
      [method]: async () => refusal } as unknown as XlsxEditorApi;
    const api = createInThreadAsyncEditorApi(sync, false, () => {});
    await expect(api[method]({ expectVersion: 'v1', steps: [] })).resolves.toBe(refusal);
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
      expect(saved).toHaveBeenCalledTimes(1);
      expect(saved).toHaveBeenCalledWith(bytes);
    } else {
      await expect(api.save()).rejects.toMatchObject({ code: 'input-failed' });
      expect(saved).not.toHaveBeenCalled();
    }
  });
}
