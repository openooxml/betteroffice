import type { XlsxEditorApi } from './XlsxEditor';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import { guardAsyncEditorApi } from './worker/guardAsyncEditorApi';

export function createInThreadAsyncEditorApi(api: XlsxEditorApi, collaborative: boolean): XlsxAsyncEditorApi {
  async function read<T>(operation: () => T): Promise<T> {
    await api.flushPendingInput();
    return operation();
  }
  const handle = api.handle;
  return guardAsyncEditorApi({
    hydrated: true,
    failure: null,
    whenHydrated: api.flushPendingInput,
    recoverySave: () => read(() => ({ bytes: api.save(), recovery: true as const })),
    commands: api.commands,
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
