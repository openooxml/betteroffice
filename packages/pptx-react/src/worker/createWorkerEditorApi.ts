import { PptxPeerNotReadyError, PptxWorkerEditorDisposedError } from '@betteroffice/pptx';
import type {
  PptxEditRefusal, PptxWorkerEditorAccess, PptxWorkerEditorFailedError, PptxWorkerEditorRecovery, PptxWorkerEditorSession,
} from '@betteroffice/pptx';
import type { PptxEditorApi } from '../PptxEditor';
import type { PptxCommandStore } from '../commands/types';
import { pptxCommandController } from '../commands/createPptxCommandStore';

/** @experimental */
export interface PptxWorkerEditorApi extends Omit<PptxEditorApi, 'handle' | 'save'> {
  readonly handle: null;
  save(): null;
  saveAsync(): Promise<Uint8Array>;
  goToSlideAsync(slide: number): Promise<boolean>;
  readonly hydrated: boolean;
  readonly failure: PptxWorkerEditorFailedError | undefined;
  whenHydrated(): Promise<void>;
  handleAsync(): Promise<PptxWorkerEditorAccess>;
  recoverySave(): Promise<PptxWorkerEditorRecovery>;
}

export function createWorkerEditorApi(
  session: PptxWorkerEditorSession,
  ui: Omit<PptxEditorApi, 'handle' | 'save'>,
  access: () => PptxWorkerEditorAccess | null,
  afterPaint: (slide: number) => Promise<boolean>,
  gestureActive: () => boolean = () => false,
  isCurrent: () => boolean = () => true,
  readOnly: () => boolean = () => false
): PptxWorkerEditorApi {
  const retired = () => {
    if (!isCurrent() || session.state.stage === 'disposed') throw new PptxWorkerEditorDisposedError();
  };
  const current = () => {
    retired();
    if (session.failure) throw session.failure;
  };
  const peer = () => {
    current();
    const value = access();
    if (!session.hydrated || !value) throw new PptxPeerNotReadyError();
    return value;
  };
  const admit = () => {
    const value = peer();
    if (gestureActive()) throw new Error('Finish the pointer gesture before flushing input');
    return value;
  };
  const refusal = (): PptxEditRefusal | null => {
    const value = peer();
    return readOnly() ? { ok: false, version: value.version(),
      failure: { code: 'read-only', message: 'The editor is read-only' } } : null;
  };
  const read = async <T,>(operation: (value: PptxWorkerEditorAccess) => T): Promise<T> => {
    const value = peer();
    await ui.flushPendingInput();
    current();
    return operation(value);
  };
  const commands = pptxCommandController(ui.commands)?.guarded(current) ?? Object.fromEntries(Object.keys(ui.commands).map((key) => [key, (...args: unknown[]) => {
    current();
    return Reflect.apply(Reflect.get(ui.commands, key), ui.commands, args);
  }])) as unknown as PptxCommandStore;
  const api = {
    get handle() { retired(); return null; },
    get commands() { current(); return commands; },
    get hydrated() { retired(); return session.hydrated; },
    get failure() {
      retired();
      return session.failure;
    },
    save: () => { retired(); return null; },
    saveAsync: async () => { admit(); const bytes = await session.saveAsync(); current(); return bytes; },
    whenHydrated: async () => { current(); await session.whenHydrated(); current(); },
    handleAsync: async () => { current(); await session.whenHydrated(); return peer(); },
    recoverySave: () => session.recoverySave(),
    goToSlideAsync: async (slide: number) => { current(); return afterPaint(slide); },
    version: () => read((value) => value.version()),
    readContent: (request) => read((value) => value.readContent(request)),
    findText: (request) => read((value) => value.findText(request)),
    validateEdits: async (request) => refusal() ?? read((value) => refusal() ?? value.validateEdits(request)),
    applyEdits: async (request) => {
      const early = refusal();
      if (early) return early;
      admit();
      const result = await ui.applyEdits(request);
      await ui.flushPendingInput();
      current();
      return result;
    },
    flushPendingInput: async () => { peer(); await ui.flushPendingInput(); current(); },
    getPositionAtPoint: (x, y) => { current(); return session.hydrated ? ui.getPositionAtPoint(x, y) : null; },
    selectText: (target) => { current(); return session.hydrated ? ui.selectText(target) : false; },
    goToSlide: (slide) => { current(); return ui.goToSlide(slide); },
    clearSelection: () => { current(); ui.clearSelection(); },
    focus: () => { current(); ui.focus(); },
    refresh: () => { current(); ui.refresh(); },
    refreshProposals: () => { current(); ui.refreshProposals(); },
  } satisfies PptxWorkerEditorApi;
  return api;
}
