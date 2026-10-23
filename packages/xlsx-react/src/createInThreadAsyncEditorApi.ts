import { XlsxSaveRefusedError, type XlsxEditorApi } from './XlsxEditor';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import { guardAsyncEditorApi } from './worker/guardAsyncEditorApi';
import { XlsxCommandAdmissionError, xlsxCommandController } from './commands/createXlsxCommandStore';

export function createInThreadAsyncEditorApi(
  api: XlsxEditorApi, collaborative: boolean, assertCurrent: () => void,
  onSave: () => ((bytes: Uint8Array) => void) | undefined = () => undefined
): XlsxAsyncEditorApi {
  async function read<T>(operation: () => T): Promise<T> {
    await api.flushPendingInput();
    return operation();
  }
  const host = async <T>(operation: () => Promise<T>): Promise<T> => {
    assertCurrent();
    try {
      const value = await operation();
      assertCurrent();
      return value;
    } catch (error) { assertCurrent(); throw error; }
  };
  const handle = api.handle;
  return guardAsyncEditorApi({
    hydrated: true,
    failure: null,
    version: () => host(() => api.version()),
    readCells: (request) => host(() => api.readCells(request)),
    findText: (request) => host(() => api.findText(request)),
    validateEdits: (request) => host(() => api.validateEdits(request)),
    applyEdits: (request) => host(() => api.applyEdits(request)),
    save: () => host(async () => {
      let bytes: Uint8Array;
      try {
        await api.flushPendingInput();
        assertCurrent();
        try { bytes = api.save(); }
        catch (error) {
          if (!(error instanceof XlsxSaveRefusedError) || error.code !== 'input-pending') throw error;
          await api.flushPendingInput();
          assertCurrent();
          bytes = api.save();
        }
      } catch (error) {
        if (error instanceof XlsxSaveRefusedError) throw new XlsxCommandAdmissionError('input-failed');
        throw error;
      }
      assertCurrent();
      onSave()?.(bytes);
      return bytes;
    }),
    whenHydrated: api.flushPendingInput,
    recoverySave: () => read(() => ({ bytes: api.save(), recovery: true as const })),
    commands: xlsxCommandController(api.commands)!.guarded(assertCurrent),
    collaboration: collaborative ? {
      clientId: handle.clientId,
      encodeStateVector: () => read(() => handle.encodeStateVector()),
      encodeStateAsUpdate: (vector) => read(() => handle.encodeStateAsUpdate(vector)),
      applyUpdate: (update) => read(() => handle.applyUpdate(update)),
      onUpdate: (listener) => handle.onUpdate(listener),
    } : null,
    focus: api.focus,
    renderPng: (viewport) => read(() => handle.renderPng(viewport)),
    renderRangePng: (options) => read(() => handle.renderRangePng(options)),
    exportStructured: (options) => read(() => handle.exportStructured(options)),
    exportMarkdown: (options, markdownOptions) => read(() => handle.exportMarkdown(options, markdownOptions)),
    printDisplayList: (...args) => read(() => handle.printDisplayList(...args)),
  });
}
