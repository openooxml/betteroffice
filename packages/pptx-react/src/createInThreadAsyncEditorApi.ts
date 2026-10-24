import { PRESENTATION_COMMAND_MUTATIONS } from '@betteroffice/pptx/worker';
import type { PptxFontFace } from '@betteroffice/pptx';
import type { PptxAsyncCollaborationOptions, PptxContentCollaborationOptions, PptxEditorApi } from './PptxEditor';
import { PptxCommandAdmissionError } from './commands/createPptxCommandStore';
import type { PptxAsyncEditorApi, PptxAsyncEditorAccess } from './worker/createPeerlessEditorApi';

/** Without a clientId, each opened handle generates its own. */
export function asyncCollaboration(options?: PptxAsyncCollaborationOptions): PptxContentCollaborationOptions | undefined {
  if (!options) return;
  const { onReplica, ...rest } = options;
  return { ...rest, onReplica: (replica) => onReplica?.(replica ? {
    clientId: replica.clientId,
    encodeStateVector: async () => replica.encodeStateVector(),
    encodeStateAsUpdate: async (vector) => replica.encodeStateAsUpdate(vector),
    applyUpdate: async (update) => replica.applyUpdate(update),
    onUpdate: (listener) => replica.onUpdate(listener),
  } : null) };
}

export function createInThreadAsyncEditorApi(api: PptxEditorApi, readOnly: () => boolean): PptxAsyncEditorApi {
  const handle = api.handle;
  async function read<T>(operation: () => T): Promise<T> {
    await api.flushPendingInput();
    return operation();
  }
  async function mutate<T>(operation: () => T): Promise<T> {
    return read(() => {
      if (readOnly()) throw new PptxCommandAdmissionError('read-only');
      const value = operation();
      api.refresh();
      return value;
    });
  }
  const accessMethods = ['snapshot', 'story', 'searchText', 'listProposals', 'comments', 'previewProposal',
    'layoutSlide', 'layoutProposalSlide', 'layoutProposalDiffSlide', 'encodeStateVector', 'encodeStateAsUpdate',
    'exportStructured', 'exportMarkdown'] as const;
  const operations = Object.fromEntries([...accessMethods, ...PRESENTATION_COMMAND_MUTATIONS].map((method) => [
    method, (...args: unknown[]) => {
      const invoke = () => Reflect.apply(handle[method], handle, args);
      return (PRESENTATION_COMMAND_MUTATIONS as readonly string[]).includes(method) ? mutate(invoke) : read(invoke);
    },
  ]));
  const register = (face: PptxFontFace, fallback: boolean) => mutate(() =>
    fallback ? handle.registerFallbackFont(face) : handle.registerFont(face));
  const access = Object.freeze({ ...operations,
    version: api.version, readContent: api.readContent, findText: api.findText,
    validateEdits: api.validateEdits, applyEdits: api.applyEdits,
    registerFont: (face: PptxFontFace) => register(face, false),
    registerFallbackFont: (face: PptxFontFace) => register(face, true),
  }) as PptxAsyncEditorAccess;
  const navigate = (slide: number) => read(() => api.goToSlide(slide));
  return {
    commands: api.commands,
    get hydrated() { return true; },
    failure: undefined,
    whenHydrated: api.flushPendingInput,
    handleAsync: () => read(() => access),
    saveAsync: () => read(api.save),
    recoverySave: () => read(() => ({ bytes: api.save(), recovery: true as const, throughSeq: 0, version: handle.version() })),
    flushPendingInput: api.flushPendingInput,
    version: api.version,
    readContent: api.readContent,
    findText: api.findText,
    exportStructured: access.exportStructured,
    exportMarkdown: access.exportMarkdown,
    validateEdits: api.validateEdits,
    applyEdits: api.applyEdits,
    getPositionAtPoint: (x, y) => read(() => api.getPositionAtPoint(x, y)),
    clearSelection: () => read(api.clearSelection),
    selectText: (target) => read(() => api.selectText(target)),
    goToSlide: navigate,
    goToSlideAsync: navigate,
    refresh: () => read(api.refresh),
    refreshProposals: () => read(api.refreshProposals),
    focus: api.focus,
  };
}
