import { PRESENTATION_COMMAND_MUTATIONS, hitGeometry, indexGeometry, PptxWorkerUnsupportedError } from '@betteroffice/pptx/worker';
import type { PptxEditRefusal, PptxFontFace } from '@betteroffice/pptx';
import type {
  PresentationCommandAccessMethods, PresentationCommandMutationMethods, PresentationCommandMutation,
  PresentationCommandAccessMethod,
} from '@betteroffice/pptx/worker';
import { invokePeerless } from './peerlessObjectCommands';
import { peerlessGestureActive } from './peerlessViewStore';
import { PptxCommandAdmissionError } from '../commands/createPptxCommandStore';
import type { PptxCommandStore } from '../commands/types';
import type { PptxPointPosition } from '../PptxEditor';
import type { PptxWorkerEditorApi } from './createWorkerEditorApi';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { LEGACY_SYNC_API } from '../runtimeOptions';

type AsyncOperations = Pick<PptxWorkerEditorApi,
  'getPositionAtPoint' | 'clearSelection' | 'selectText' | 'goToSlide' | 'refresh' | 'refreshProposals'>;

export type PptxAsyncEditorAccess = {
  [K in PresentationCommandAccessMethod]: (...args: Parameters<PresentationCommandAccessMethods[K]>) => Promise<ReturnType<PresentationCommandAccessMethods[K]>>;
} & { [K in keyof PresentationCommandMutationMethods]: (...args: Parameters<PresentationCommandMutationMethods[K]>) => Promise<ReturnType<PresentationCommandMutationMethods[K]>>
} & Pick<PptxWorkerEditorApi, 'version' | 'readContent' | 'findText' | 'validateEdits' | 'applyEdits'> & {
  registerFont(face: PptxFontFace): Promise<number>;
  registerFallbackFont(face: PptxFontFace): Promise<number>;
};

export type PptxAsyncEditorApi = Omit<PptxWorkerEditorApi,
  'handle' | 'save' | 'handleAsync' | 'recoverySave' | keyof AsyncOperations> & {
  [K in keyof AsyncOperations]: (...args: Parameters<AsyncOperations[K]>) => Promise<ReturnType<AsyncOperations[K]>>;
} & Pick<PptxAsyncEditorAccess, 'exportStructured' | 'exportMarkdown'> & {
  handleAsync(): Promise<PptxAsyncEditorAccess>;
  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true; throughSeq: number; version: string }>;
};

const caller = { reportRefusal: false } as const;

function unsupported(operation: string): never {
  throw new PptxWorkerUnsupportedError(operation);
}

export function peerlessPositionAtPoint(run: PeerlessPresentation, canvas: HTMLCanvasElement,
  clientX: number, clientY: number): PptxPointPosition | null {
  run.require();
  const bounds = canvas.getBoundingClientRect();
  if (!bounds.width || !bounds.height || clientX < bounds.left || clientX > bounds.right ||
    clientY < bounds.top || clientY > bounds.bottom) return null;
  const frame = run.getVisualSnapshot()?.view.confirmed;
  if (!frame) return null;
  const point = { x: (clientX - bounds.left) * frame.displayList.width / bounds.width,
    y: (clientY - bounds.top) * frame.displayList.height / bounds.height };
  const hit = hitGeometry(indexGeometry(frame.displayList), point);
  return hit ? { ...hit, slide: frame.frame.slideIndex + 1, slideId: frame.frame.slideId } : null;
}

export function createPeerlessEditorApi(run: PeerlessPresentation, canvas: HTMLCanvasElement,
  focus: () => void, commands?: PptxCommandStore): PptxAsyncEditorApi {
  function admit(): void {
    run.require();
    if (peerlessGestureActive(run.store)) throw new PptxCommandAdmissionError('gesture-active');
  }
  async function fenced<T>(operation: () => Promise<T>): Promise<T> {
    admit();
    await run.store.fence();
    admit();
    const result = await operation();
    run.require();
    return result;
  }
  async function installed(): Promise<void> { run.require(); await run.ready; run.require(); }
  const read = <K extends PresentationCommandAccessMethod>(method: K, ...args: Parameters<PresentationCommandAccessMethods[K]>) =>
    fenced(() => run.client.access?.(method, ...args) ?? Promise.reject(new PptxWorkerUnsupportedError(method)));
  const register = async (face: PptxFontFace, fallback: boolean): Promise<number> => {
    admit();
    const operation = run.store.dispatch({ kind: 'registerFont', payload: { font: { ...face, bytes: face.bytes.slice().buffer },
      fallback, inputTime: performance.now(), undoGroup: 'font' } }, caller);
    await run.installFont(face);
    const ack = await operation;
    run.require();
    if (ack.fontId === undefined) throw new Error(ack.outcomes[0]?.reason ?? 'Font registration refused');
    await run.client.flush();
    await installed();
    return ack.fontId;
  };
  const navigate = async (slide: number): Promise<boolean> => {
    admit();
    if (!Number.isSafeInteger(slide) || slide < 1 || slide > run.client.state.slides.length) return false;
    const slideId = run.client.state.slides[slide - 1]!.id;
    const ack = await run.store.dispatch({ kind: 'activateSlide', payload: { slideId } }, caller);
    await installed();
    return !ack.outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled') &&
      run.store.getSnapshot().confirmed?.frame.slideId === slideId;
  };
  const api: PptxAsyncEditorApi = {
    get commands() { run.require(); return commands ?? unsupported('commands'); },
    get hydrated() { run.require(); return run.hydrated; },
    get failure() { return run.failure; },
    saveAsync: () => fenced(() => run.client.save()),
    recoverySave: () => run.recover(),
    whenHydrated: installed,
    handleAsync: async () => { await installed(); return access; },
    flushPendingInput: () => fenced(async () => { await run.client.flush(); await installed(); }),
    version: () => fenced(async () => (await run.client.flush()).version),
    readContent: (request) => fenced(async () => (await run.client.read(request)).result),
    findText: (request) => read('findText', request),
    exportStructured: (options) => read('exportStructured', options),
    exportMarkdown: (options) => read('exportMarkdown', options),
    validateEdits: (request) => read('validateEdits', request),
    applyEdits: async (request) => {
      admit();
      if (run.store.getSnapshot().editState.readOnly) {
        const ack = await fenced(() => run.client.flush());
        return { ok: false, version: ack.version,
          failure: { code: 'read-only', message: 'The editor is read-only' } } satisfies PptxEditRefusal;
      }
      const ack = await run.store.dispatch({ kind: 'applyEdits', payload: { request, inputTime: performance.now(), undoGroup: 'batch' } }, caller);
      run.require();
      await installed();
      return ack.editResult ?? { ok: false, version: ack.version, failure: {
        code: ack.outcomes[0]?.reason === 'read-only' ? 'read-only' : 'invalid-step',
        message: ack.outcomes[0]?.reason ?? 'Edit batch refused' } } satisfies PptxEditRefusal;
    },
    selectText: async (target) => {
      admit();
      if (!Number.isSafeInteger(target.slide) || target.slide < 1 || target.slide > run.client.state.slides.length) return false;
      const slideId = run.client.state.slides[target.slide - 1]!.id;
      const ack = await run.store.dispatch({ kind: 'setSelection', payload: { selection: { slideId,
        shapeId: target.shapeId, storyId: target.storyId, anchor: target.start, focus: target.end, affinity: 'downstream' },
        inputTime: performance.now(), undoGroup: 'selection' } }, caller);
      await installed();
      return !ack.outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
    },
    goToSlide: navigate, goToSlideAsync: navigate,
    refresh: () => fenced(async () => { await run.client.flush(); await installed(); }),
    refreshProposals: async () => { await read('listProposals'); },
    clearSelection: async () => {
      admit();
      await run.store.dispatch({ kind: 'setSelection', payload: { selection: null,
        inputTime: performance.now(), undoGroup: 'selection' } });
      await installed();
    },
    focus() { run.require(); focus(); },
    getPositionAtPoint: (clientX, clientY) => fenced(async () => {
      await installed(); return peerlessPositionAtPoint(run, canvas, clientX, clientY);
    }),
  };
  const mutations = Object.fromEntries(PRESENTATION_COMMAND_MUTATIONS.map((method) => [method, async (...args: unknown[]) => {
    admit();
    const ack = await invokePeerless(run, { method, args } as PresentationCommandMutation, caller);
    run.require();
    const refused = ack.outcomes.find((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled');
    if (refused) {
      throw new PptxCommandAdmissionError(refused.reason === 'invalid-payload' ? 'invalid-arguments'
        : refused.reason === 'read-only' ? 'read-only' : 'target-changed');
    }
    await installed();
    return ack.mutationResult;
  }])) as Pick<PptxAsyncEditorAccess, keyof PresentationCommandMutationMethods>;
  const access: PptxAsyncEditorAccess = Object.freeze({
    ...mutations,
    encodeStateVector: () => read('encodeStateVector'),
    encodeStateAsUpdate: (vector) => read('encodeStateAsUpdate', vector),
    previewProposal: (id) => read('previewProposal', id), layoutSlide: (index) => read('layoutSlide', index),
    layoutProposalSlide: (id, index) => read('layoutProposalSlide', id, index), layoutProposalDiffSlide: (id, index) => read('layoutProposalDiffSlide', id, index),
    comments: () => read('comments'), snapshot: () => read('snapshot'), story: (id) => read('story', id), findText: api.findText,
    searchText: (query, options) => read('searchText', query, options), listProposals: () => read('listProposals'),
    exportStructured: api.exportStructured, exportMarkdown: api.exportMarkdown,
    version: api.version, readContent: api.readContent, validateEdits: api.validateEdits, applyEdits: api.applyEdits,
    registerFont: (face) => register(face, false), registerFallbackFont: (face) => register(face, true),
  });
  return Object.defineProperties(api, {
    handle: { get() { throw new PptxWorkerUnsupportedError('handle', `handleAsync() or ${LEGACY_SYNC_API}={true} for synchronous access`); } },
    save: { value() { throw new PptxWorkerUnsupportedError('save', `saveAsync() or ${LEGACY_SYNC_API}={true} for synchronous access`); } },
  });
}
