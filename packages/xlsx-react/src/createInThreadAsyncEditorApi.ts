import { XlsxSaveRefusedError } from './XlsxSaveRefusedError';
import type { XlsxEditorApi } from './XlsxEditor.types';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import { guardAsyncEditorApi } from './worker/guardAsyncEditorApi';
import { XlsxCommandAdmissionError, xlsxCommandController } from './commands/createXlsxCommandStore';
import { commandReason } from './commands/evaluate';

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
    readCells(request) {
      const captured = structuredClone(request);
      return host(() => api.readCells(captured));
    },
    findText(request) {
      const captured = structuredClone(request);
      return host(() => api.findText(captured));
    },
    validateEdits(request) {
      const captured = structuredClone(request);
      return host(() => api.validateEdits(captured));
    },
    applyEdits(request) {
      const captured = structuredClone(request);
      return host(() => api.applyEdits(captured));
    },
    proposeChanges(request) {
      const captured = structuredClone(request);
      return host(() => api.proposeChanges ? api.proposeChanges(captured) : Promise.resolve({
        ok: false as const, failure: commandReason('unsupported-command', null),
      }));
    },
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
      encodeStateAsUpdate(vector) {
        const captured = vector?.slice();
        return read(() => handle.encodeStateAsUpdate(captured));
      },
      applyUpdate(update) {
        const captured = update.slice();
        return read(() => handle.applyUpdate(captured));
      },
      onUpdate: (listener) => handle.onUpdate(listener),
    } : null,
    focus: api.focus,
    async renderPng(target, options = {}) {
      const { signal, ...rest } = options;
      const captured = structuredClone(target);
      const wire = structuredClone(rest);
      return await read(() => handle.renderPng(captured, { ...wire, signal }));
    },
    exportStructured(options) {
      const captured = structuredClone(options);
      return read(() => handle.exportStructured(captured));
    },
    exportMarkdown(options, markdownOptions) {
      const captured = structuredClone([options, markdownOptions] as const);
      return read(() => handle.exportMarkdown(...captured));
    },
    printDisplayList(...args) {
      const captured = structuredClone(args);
      return read(() => handle.printDisplayList(...captured));
    },
  });
}
