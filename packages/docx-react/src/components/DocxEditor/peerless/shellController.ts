import { createT, deepMerge, en, type LocaleStrings } from '@betteroffice/docx-i18n';
import { WASM32_MEMORY_LIMIT_BYTES, type ResidentEngineWorkerClient, type ResidentEngineWorkerFrame,
  type ResidentFontEnvironment, type ResidentFontEnvironmentEvent, type WasmModuleMemory } from '@betteroffice/docx/yrs';
import { normalizeFontFamilies } from '@betteroffice/docx/utils/fontOptions';
import { DEFAULT_FONTS } from '../../ui/FontPicker';
import { getPrimaryFontFamily } from '../../ui/fontPickerValue';
import { memoryPressureLevel, type DocxMemoryPressureLevel, type DocxMemoryStats } from '../memoryStats';
import type { DocxInput } from '@betteroffice/docx/utils';
import type { DocxCommandSelection, ResidentDocumentRead, ResidentDocumentReadValues } from '@betteroffice/docx/yrs';
import { createDocxCommandController, type DocxCommandBinding } from '../../../commands/createDocxCommandStore';
import type { DocxCommandEnvironment } from '../../../commands/evaluate';
import type { DocxCommandId, DocxCommandResult } from '../../../commands/types';
import type { DocxDocumentChange, DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { PeerlessViewSnapshot } from './viewStore';
import { readWorkerSelection } from './clipboard';
import type { PeerlessInputIntent } from './inputEvents';
import { flushAdmittedInput, type PeerlessAdmissionResult } from './inputAdmission';
import type { DocxHostSearch } from '../hooks/useHostSearch';
import type { SelectionState } from '../types';
import { yrsSelectionState } from '../hooks/useSelectionTracker';

import { PeerlessSyncApiError } from './contracts';
export { LEGACY_SYNC_API_PROP, PeerlessSyncApiError } from './contracts';

export interface PeerlessWorkerQuery<K extends ResidentDocumentRead['kind'] = ResidentDocumentRead['kind']> {
  version: string;
  value: ResidentDocumentReadValues[K];
  selection: DocxCommandSelection | null | undefined;
}

interface Options extends Pick<DocxEditorProps, 'i18n' | 'theme' | 'mode' | 'readOnly' | 'onDocumentChange' | 'onSelectionChange' | 'onError' | 'onSave' | 'showFileOpen' | 'author' | 'onModeChange' | 'fontFamilies' | 'onFontsLoaded' | 'onFirstPagePainted' | 'onMemoryPressure' | 'memoryBudget'> {
  zoom: number;
  isDark?: boolean;
  focus(): void;
  onZoom?(zoom: number): void;
  read: ResidentEngineWorkerClient['documentRead'];
  readAt?: ResidentEngineWorkerClient['documentReadAt'];
  search?: DocxHostSearch;
  save?(): Promise<ArrayBuffer>;
  loadBuffer?(buffer: DocxInput): Promise<void>;
  print?(preview?: boolean): Promise<void>;
  currentPage?(): number;
  scrollToPage?(number: number): void;
  scrollToParagraph?(paraId: string): Promise<boolean>;
  perform?(id: DocxCommandId, args: unknown): Promise<DocxCommandResult>;
  prepareFontFamily?(family: string): Promise<ResidentFontEnvironment | void>;
}

export function createPeerlessShellController(initial: Options) {
  let options = initial;
  let binding: PeerlessEditSessionBinding | null = null;
  let unsubscribe: (() => void) | undefined;
  let snapshot: { status: 'empty' | 'opening' | 'ready' | 'failed'; view: PeerlessViewSnapshot | null; error: Error | null; mode: NonNullable<DocxEditorProps['mode']>; preview: ResidentEngineWorkerFrame | null; previewPainted: boolean } =
    { status: 'empty', view: null, error: null, mode: initial.mode ?? 'editing', preview: null, previewPainted: false };
  const listeners = new Set<() => void>();
  const changes = new Set<(change: DocxDocumentChange) => void>();
  const selections = new Set<(selection: SelectionState | null) => void>();
  const commands = createDocxCommandController();
  let translate = translation(options);
  let internalMode: NonNullable<DocxEditorProps['mode']> = 'editing';
  let fontRevision = 0;
  let fontFamilies: string[] = [];
  let firstPainted = false;
  let workerMemory: WasmModuleMemory[] | null = null;
  let pressure: DocxMemoryPressureLevel = 'normal';
  const memoryStats = (): DocxMemoryStats => ({ limitBytes: WASM32_MEMORY_LIMIT_BYTES, main: [],
    worker: workerMemory?.map(sample => ({ ...sample })) ?? null });
  const mode = () => options.mode ?? internalMode;
  const updateAuthor = () => binding?.admission.setSuggesting?.(mode() === 'suggesting' ? options.author ?? 'User' : null);
  const emit = <T>(listener: ((value: T) => void) | undefined, value: T) => {
    try { listener?.(value); } catch (error) {
      try { options.onError?.(error instanceof Error ? error : new Error(String(error))); } catch {}
    }
  };
  const selectionChanged = (view: PeerlessViewSnapshot | null) => {
    const selection = view?.workerSelection;
    const toolbar = view?.editState?.toolbar;
    const state = !selection ? null : toolbar ? yrsSelectionState({ context: toolbar, startParagraphIndex: -1, endParagraphIndex: -1 }) : {
      hasSelection: selection.anchor.story !== selection.focus.story || selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset,
      isMultiParagraph: selection.anchor.story !== selection.focus.story || selection.anchor.paraId !== selection.focus.paraId,
      textFormatting: {}, paragraphFormatting: {}, styleId: null, startParagraphIndex: -1, endParagraphIndex: -1,
    };
    emit(options.onSelectionChange, state);
    for (const listener of selections) emit(listener, state);
  };
  const reset = (status: typeof snapshot.status, error: Error | null = null) => {
    const selected = snapshot.view?.workerSelection;
    fontRevision = 0; fontFamilies = []; firstPainted = false; workerMemory = null;
    unsubscribe?.(); binding = null; snapshot = { status, view: null, error, mode: mode(), preview: null, previewPainted: false }; publish();
    if (selected) selectionChanged(null);
  };
  const publish = () => { snapshot = { ...snapshot, mode: mode() }; updateAuthor(); commands.refresh(); for (const listener of listeners) listener(); };
  const painted = () => { if (firstPainted) return; firstPainted = true; emit(options.onFirstPagePainted, undefined); };
  const reportMemory = () => {
    const stats = memoryStats();
    const next = memoryPressureLevel(stats, options.memoryBudget);
    if (next === pressure) return;
    pressure = next; emit(options.onMemoryPressure, { level: next, stats });
  };
  const environment = (): DocxCommandEnvironment => {
    const view = snapshot.view;
    const toolbar = view?.editState?.toolbar;
    const choices = normalizeFontFamilies(options.fontFamilies) ?? DEFAULT_FONTS;
    return {
      status: snapshot.status === 'failed' ? 'unavailable' : snapshot.status === 'opening' ? 'loading' : snapshot.status,
      readOnly: options.readOnly === true || (mode() !== 'viewing' && view?.editState?.readOnly === true),
      mode: mode(), modeControl: options.mode ? options.onModeChange ? 'host' : 'fixed' : 'internal',
      sidebarOpen: false, sidebarControl: 'fixed', zoom: options.zoom, zoomWhileOpening: true, canOpen: options.showFileOpen !== false, canReportIssue: false,
      bodyStory: (view?.workerSelection?.focus.story ?? 'body') === 'body', pendingInput: false,
      supportedCommands: new Set(['zoom', 'open', 'save', 'print', 'undo', 'redo', 'find', 'replace', 'editingMode', 'bold', 'italic', 'underline', 'pageSetup', 'watermark',
        'fontFamily', 'fontSize', ...(view?.workerSelection && (view.workerSelection.anchor.paraId !== view.workerSelection.focus.paraId || view.workerSelection.anchor.offset !== view.workerSelection.focus.offset) ? ['insertLink' as const] : [])]),
      canUndo: view?.editState?.canUndo ?? false, canRedo: view?.editState?.canRedo ?? false,
      selection: toolbar ? { context: toolbar, fontFamily: toolbar.fontFamily ?? null,
        fontSize: view?.editState?.effectiveFormatting?.fontSizePt ?? null } : null,
      table: view?.editState?.tableContext ? { isInTable: true, ...view.editState.tableContext } : null,
      image: null, revisions: [], revisionIds: new Set(), currentRevisionId: null, styles: [], fonts: [
        ...choices.map(font => ({ args: { family: getPrimaryFontFamily(font.fontFamily) || font.name }, label: font.name,
          preview: { fontFamily: font.fontFamily, group: font.category ?? 'other' } })),
        ...fontFamilies.filter(family => !choices.some(font => font.name.toLowerCase() === family.toLowerCase()))
          .map(family => ({ args: { family }, label: family, preview: { group: 'document' } })),
      ], translate,
    };
  };
  const queryOrigins = new WeakMap<object, { binding: PeerlessEditSessionBinding; identity: PeerlessViewSnapshot['identity'] }>();
  const capture = () => binding;
  const ready = () => {
    if (!binding || snapshot.status !== 'ready') throw new Error('Peer-less editor is not ready');
    return binding;
  };
  const flush = async () => {
    const current = ready();
    await flushAdmittedInput(current);
    if (current !== binding) throw new Error('Document was replaced');
    return current;
  };
  const read = async <K extends ResidentDocumentRead['kind']>(request: ResidentDocumentRead & { kind: K }): Promise<ResidentDocumentReadValues[K]> => {
    const current = await flush();
    const version = current.viewStore.getSnapshot().version;
    const result = await options.read<K>(request);
    if (current !== binding) throw new Error('Document was replaced');
    if (version !== current.viewStore.getSnapshot().version) throw new Error('Worker read was superseded');
    return result.value;
  };
  const readAt = async <K extends ResidentDocumentRead['kind']>(request: ResidentDocumentRead & { kind: K }, token: string): Promise<ResidentDocumentReadValues[K]> => {
    const current = await flush();
    const view = current.viewStore.getSnapshot();
    if (!options.readAt) throw new Error('Versioned worker reads are unavailable');
    const result = await options.readAt<K>(request, token);
    const latest = current.viewStore.getSnapshot();
    if (current !== binding) throw new Error('Document was replaced');
    if (result.status !== 'ok' || result.version !== token || latest.version !== view.version || latest.identity !== view.identity) throw new Error('Worker read was superseded');
    return result.value;
  };
  const settleCommand = async (current: PeerlessEditSessionBinding, result: PeerlessAdmissionResult): Promise<DocxCommandResult> => {
    if (result.status !== 'submitted') return { ok: false, failure: { code: 'input-failed', message: 'Worker command was not admitted.' } };
    const outcome = await result.ticket.ack;
    if (current !== binding) return { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } };
    if (outcome.status === 'applied' || outcome.status === 'noop') return { ok: true, status: outcome.status === 'applied' ? 'executed' : 'noop' };
    return { ok: false, failure: { code: outcome.reason === 'read-only' ? 'read-only' : 'command-failed', message: outcome.reason ?? 'Worker command failed.' } };
  };
  const submit = (intent: PeerlessInputIntent): Promise<DocxCommandResult> => {
    const current = ready();
    return settleCommand(current, current.commands.submit(intent, { fence: true }));
  };
  const query = async <K extends ResidentDocumentRead['kind']>(request: ResidentDocumentRead & { kind: K }): Promise<PeerlessWorkerQuery<K>> => {
    const current = await flush();
    const view = current.viewStore.getSnapshot();
    const native = await options.read({ kind: 'readParagraphs', request: { view: 'accepted', paraIds: [] } });
    if (current !== binding || view.identity !== current.viewStore.getSnapshot().identity || view.throughSeq !== current.viewStore.getSnapshot().throughSeq) throw new Error('Worker read was superseded');
    const value = await readAt<K>(request, native.version);
    if (view.throughSeq !== current.viewStore.getSnapshot().throughSeq) throw new Error('Worker read was superseded');
    const result = { version: native.version, value, selection: structuredClone(view.workerSelection) };
    queryOrigins.set(result, { binding: current, identity: view.identity });
    return result;
  };
  const applyQuery = async (origin: PeerlessWorkerQuery, intent: PeerlessInputIntent): Promise<DocxCommandResult> => {
    const issued = queryOrigins.get(origin);
    if (!issued || issued.binding !== binding || JSON.stringify(issued.identity) !== JSON.stringify(binding.viewStore.getSnapshot().identity)) return { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } };
    const current = await flush();
    if (issued.binding !== current || JSON.stringify(issued.identity) !== JSON.stringify(current.viewStore.getSnapshot().identity)) return { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } };
    return submit(intent);
  };
  const navigateCell = (direction: 'forward' | 'backward') => {
    const current = ready();
    const env = environment();
    return current.commands.submit({ kind: 'navigateCell', payload: { direction, extend: false,
      createLastRow: direction === 'forward' && !env.readOnly && env.mode !== 'viewing' && !current.admission.getSnapshot().readOnly,
    } }, { source: 'selection', fence: true });
  };
  const commandBinding: DocxCommandBinding = {
    environment,
    ordered: (id) => !['zoom', 'open', 'print', 'bold', 'italic', 'underline', 'fontFamily', 'fontSize'].includes(id),
    admit: async (operation) => { await flush(); return operation(); },
    async perform(id, args) {
      if (id === 'editingMode') {
        const next = (args as { mode: NonNullable<DocxEditorProps['mode']> }).mode;
        if (options.mode) { emit(options.onModeChange, next); return { ok: true, status: 'requested' }; }
        internalMode = next; publish(); emit(options.onModeChange, next);
        return { ok: true, status: 'executed' };
      }
      if (id === 'undo' || id === 'redo') return submit({ kind: id, payload: {} });
      if (id === 'bold' || id === 'italic' || id === 'underline') {
        const current = ready();
        return settleCommand(current, current.commands.format({ kind: 'toggleMark', payload: { mark: { type: id } } }, { fence: true }));
      }
      if (id === 'fontFamily' || id === 'fontSize') {
        const current = ready();
        if (id === 'fontFamily') {
          const fonts = await options.prepareFontFamily?.((args as { family: string }).family);
          if (current !== binding) return { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } };
          if (fonts) {
            try {
              const prepared = await settleCommand(current, current.commands.submit({ kind: 'setFontEnvironment', payload: fonts }, { fence: true }));
              if (!prepared.ok) return prepared;
            } finally { fonts.faces.length = 0; }
          }
        }
        const delta = id === 'fontFamily' ? { fontFamily: { ascii: (args as { family: string }).family, hAnsi: (args as { family: string }).family } } :
          { fontSize: (args as { points: number }).points };
        return settleCommand(current, current.commands.format({ kind: 'formatRange', payload: { delta } }, { fence: true }));
      }
      if (id === 'zoom') {
        options = { ...options, zoom: (args as { scale: number }).scale };
        options.onZoom?.(options.zoom);
        publish();
        return { ok: true, status: 'executed' };
      }
      return options.perform?.(id, args) ?? { ok: false, failure: { code: 'unsupported-command', message: 'Worker command is unavailable.' } };
    },
    capture: () => { const current = capture(); return current ? { document: current } : null; },
    resume: (origin) => origin.document === binding ? null : 'document-replaced',
    chrome: () => ({ i18n: options.i18n, isDark: options.isDark ?? false, theme: options.theme ?? null }),
    focusEditor: () => options.focus(),
  };
  commands.attach(commandBinding);
  const implemented = {
    commands: commands.store,
    flushPendingInput: async () => { await flush(); },
    readParagraphs: (request) => read({ kind: 'readParagraphs', request }),
    getParagraphIdentities: () => read({ kind: 'paragraphIdentities' }),
    resolveParagraphAnchors: async (anchors) => {
      const current = await flush();
      const result = await options.read({ kind: 'resolveParagraphAnchors', anchors: [...anchors] });
      if (current !== binding) throw new Error('Document was replaced');
      return { version: result.version, ...result.value };
    },
    search: (query, config) => options.search?.search(query, config) ?? Promise.reject(new Error('Worker search is unavailable')),
    searchNext: () => options.search?.searchNext() ?? null,
    searchPrevious: () => options.search?.searchPrevious() ?? null,
    searchGoTo: (index) => options.search?.searchGoTo(index) ?? null,
    clearSearch: () => options.search?.clearSearch(),
    getSearchState: () => options.search?.getSearchState() ?? null,
    onSearchChange: (listener) => options.search?.onSearchChange(listener) ?? (() => {}),
    findText: (request) => read({ kind: 'findText', request }),
    readSelectionInfo: async () => {
      const current = await flush();
      const result = await readWorkerSelection({ binding: current, isCurrent: () => binding === current,
        readSelection: (selection) => read({ kind: 'rawSelectionInfo', selection }) });
      return result?.info ?? null;
    },
    findParagraphs: (query, config) => read({ kind: 'findParagraphs', query, ...config }),
    listContentControls: (config) => read({ kind: 'listContentControls', options: config ?? {} }),
    findContentControls: (query, config) => read({ kind: 'findContentControls', query, options: config ?? {} }),
    save: async () => { const current = await flush(); const result = await options.save?.();
      if (current !== binding) throw new Error('Document was replaced'); if (result) options.onSave?.(result); return result ?? null; },
    loadDocumentBuffer: (buffer) => options.loadBuffer?.(buffer) ?? Promise.reject(new Error('Worker loading is unavailable')),
    print: () => { void options.print?.().catch(() => {}); },
    openPrintPreview: () => { void options.print?.(true).catch(() => {}); },
    getCurrentPage: () => options.currentPage?.() ?? 1,
    getTotalPages: () => snapshot.view?.geometry?.pages.length ?? 0,
    scrollToPage: (number) => options.scrollToPage?.(number),
    scrollToParagraph: (paraId) => options.scrollToParagraph?.(paraId) ?? Promise.resolve(false),
    focus: () => options.focus(),
    setZoom: (zoom) => { if (!Number.isFinite(zoom) || zoom <= 0) return;
      options = { ...options, zoom }; options.onZoom?.(zoom); publish(); },
    getZoom: () => options.zoom,
    getMemoryStats: memoryStats,
    onDocumentChange: (listener) => { changes.add(listener); return () => { changes.delete(listener); }; },
    onSelectionChange: (listener) => { selections.add(listener); return () => { selections.delete(listener); }; },
  } satisfies Partial<DocxEditorRef>;
  const ref = new Proxy(implemented as DocxEditorRef, {
    get(target, member) {
      if (member in target) return Reflect.get(target, member);
      return () => { throw new PeerlessSyncApiError(String(member)); };
    },
  });
  return {
    commands, ref, read, readAt, query, applyQuery, submit, navigateCell, environment,
    getSnapshot: () => snapshot,
    subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    update(next: Options) { options = next; translate = translation(next);
      if (snapshot.mode !== mode()) publish(); else { updateAuthor(); commands.refresh(); }
      reportMemory();
    },
    opening() { reset('opening'); },
    preview(frame: ResidentEngineWorkerFrame) {
      if (snapshot.status !== 'opening') return;
      snapshot = { ...snapshot, preview: frame, previewPainted: false }; publish();
    },
    previewPresented(frame: ResidentEngineWorkerFrame) {
      if (snapshot.preview !== frame || snapshot.status !== 'opening') return;
      snapshot = { ...snapshot, previewPainted: true }; publish();
    },
    painted,
    fontChoices(families: string[]) { fontFamilies = [...families]; publish(); },
    fontEnvironment(event: ResidentFontEnvironmentEvent) {
      if (event.phase !== 'ready' || event.revision <= fontRevision || snapshot.status === 'empty' || snapshot.status === 'failed') return;
      fontRevision = event.revision; fontFamilies = [...event.families]; publish(); emit(options.onFontsLoaded, undefined);
    },
    memory(samples: WasmModuleMemory[]) { workerMemory = samples.map(sample => ({ ...sample })); reportMemory(); },
    empty() { reset('empty'); },
    fail(error: Error) { reset('failed', error); },
    attach(next: PeerlessEditSessionBinding) {
      if (next === binding) return;
      unsubscribe?.(); binding = next;
      const accept = () => {
        if (binding !== next) return;
        const view = next.viewStore.getSnapshot();
        const previous = snapshot.view;
        const ready = view.presentedThroughSeq >= 0;
        snapshot = { ...snapshot, status: ready ? 'ready' : 'opening', view, error: null, mode: mode(),
          preview: ready ? null : snapshot.preview };
        publish();
        if (ready && view.geometry?.pages.some(page => page.displayPage?.unbuilt !== true)) painted();
        if (previous && view.version !== previous.version) {
          const change = { version: String(view.version) };
          emit(options.onDocumentChange, change);
          for (const listener of changes) emit(listener, change);
        }
        if ((previous || view.workerSelection) &&
          (JSON.stringify(previous?.workerSelection) !== JSON.stringify(view.workerSelection) ||
            JSON.stringify(previous?.editState?.toolbar) !== JSON.stringify(view.editState?.toolbar))) selectionChanged(view);
      };
      unsubscribe = next.viewStore.subscribe(accept);
      accept();
    },
    detach(previous: PeerlessEditSessionBinding) {
      if (binding !== previous) return;
      reset('empty');
    },
    dispose() { unsubscribe?.(); binding = null; commands.detach(commandBinding); listeners.clear(); changes.clear(); selections.clear(); },
  };
}

function translation(options: Options) {
  const strings = deepMerge(en as Record<string, unknown>, options.i18n as Record<string, unknown> | undefined) as LocaleStrings;
  return createT(strings, typeof options.i18n?._lang === 'string' ? options.i18n._lang : 'en');
}
