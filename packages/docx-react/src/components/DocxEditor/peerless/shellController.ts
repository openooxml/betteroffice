import { createT, deepMerge, en, type LocaleStrings } from '@betteroffice/docx-i18n';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxInput } from '@betteroffice/docx/utils';
import type { ResidentDocumentRead, ResidentDocumentReadValues } from '@betteroffice/docx/yrs';
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

interface Options extends Pick<DocxEditorProps, 'i18n' | 'theme' | 'mode' | 'readOnly' | 'onDocumentChange' | 'onSelectionChange' | 'onError' | 'onSave' | 'showFileOpen' | 'author' | 'onModeChange'> {
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
}

export function createPeerlessShellController(initial: Options) {
  let options = initial;
  let binding: PeerlessEditSessionBinding | null = null;
  let unsubscribe: (() => void) | undefined;
  let snapshot: { status: 'empty' | 'opening' | 'ready' | 'failed'; view: PeerlessViewSnapshot | null; error: Error | null; mode: NonNullable<DocxEditorProps['mode']> } =
    { status: 'empty', view: null, error: null, mode: initial.mode ?? 'editing' };
  const listeners = new Set<() => void>();
  const changes = new Set<(change: DocxDocumentChange) => void>();
  const selections = new Set<(selection: SelectionState | null) => void>();
  const commands = createDocxCommandController();
  let translate = translation(options);
  let internalMode: NonNullable<DocxEditorProps['mode']> = 'editing';
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
    unsubscribe?.(); binding = null; snapshot = { status, view: null, error, mode: mode() }; publish();
    if (selected) selectionChanged(null);
  };
  const publish = () => { snapshot = { ...snapshot, mode: mode() }; updateAuthor(); commands.refresh(); for (const listener of listeners) listener(); };
  const environment = (): DocxCommandEnvironment => {
    const view = snapshot.view;
    const toolbar = view?.editState?.toolbar;
    return {
      status: snapshot.status === 'failed' ? 'unavailable' : snapshot.status === 'opening' ? 'loading' : snapshot.status,
      readOnly: options.readOnly === true || (mode() !== 'viewing' && view?.editState?.readOnly === true),
      mode: mode(), modeControl: options.mode ? options.onModeChange ? 'host' : 'fixed' : 'internal',
      sidebarOpen: false, sidebarControl: 'fixed', zoom: options.zoom, zoomWhileOpening: true, canOpen: options.showFileOpen !== false, canReportIssue: false,
      bodyStory: (view?.workerSelection?.focus.story ?? 'body') === 'body', pendingInput: false,
      supportedCommands: new Set(['zoom', 'open', 'save', 'print', 'undo', 'redo', 'find', 'replace', 'editingMode', 'bold', 'italic', 'underline']),
      canUndo: view?.editState?.canUndo ?? false, canRedo: view?.editState?.canRedo ?? false,
      selection: toolbar ? { context: toolbar, fontFamily: toolbar.fontFamily ?? null,
        fontSize: view?.editState?.effectiveFormatting?.fontSizePt ?? null } : null,
      table: view?.editState?.tableContext ? { isInTable: true, ...view.editState.tableContext } : null,
      image: null, revisions: [], revisionIds: new Set(), currentRevisionId: null, styles: [], fonts: [], translate,
    };
  };
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
  const navigateCell = (direction: 'forward' | 'backward') => {
    const current = ready();
    const env = environment();
    return current.commands.submit({ kind: 'navigateCell', payload: { direction, extend: false,
      createLastRow: direction === 'forward' && !env.readOnly && env.mode !== 'viewing' && !current.admission.getSnapshot().readOnly,
    } }, { source: 'selection', fence: true });
  };
  const commandBinding: DocxCommandBinding = {
    environment,
    ordered: (id) => !['zoom', 'open', 'print', 'bold', 'italic', 'underline'].includes(id),
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
    commands, ref, read, readAt, submit, navigateCell, environment,
    getSnapshot: () => snapshot,
    subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    update(next: Options) { options = next; translate = translation(next);
      if (snapshot.mode !== mode()) publish(); else { updateAuthor(); commands.refresh(); }
    },
    opening() { reset('opening'); },
    empty() { reset('empty'); },
    fail(error: Error) { reset('failed', error); },
    attach(next: PeerlessEditSessionBinding) {
      if (next === binding) return;
      unsubscribe?.(); binding = next;
      const accept = () => {
        if (binding !== next) return;
        const view = next.viewStore.getSnapshot();
        const previous = snapshot.view;
        snapshot = { status: view.presentedThroughSeq >= 0 ? 'ready' : 'opening', view, error: null, mode: mode() };
        publish();
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
