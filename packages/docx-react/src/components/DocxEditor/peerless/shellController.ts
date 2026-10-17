import { createT, deepMerge, en, type LocaleStrings } from '@betteroffice/docx-i18n';
import { WASM32_MEMORY_LIMIT_BYTES, type DocxEditRequest, type DocxEditResult, type ResidentEngineWorkerClient, type ResidentEngineWorkerFrame,
  type ResidentFontEnvironment, type ResidentFontEnvironmentEvent, type WasmModuleMemory } from '@betteroffice/docx/yrs';
import { computeSplitDialogDefaults, resolveColorToHex } from '@betteroffice/docx/utils';
import type { YrsCellBorders } from '@betteroffice/docx/yrs';
import { getBuiltinTableStyle } from '../../ui/TableStyleGallery';
import { normalizeFontFamilies } from '@betteroffice/docx/utils/fontOptions';
import { DEFAULT_FONTS } from '../../ui/FontPicker';
import { DEFAULT_STYLES } from '../../ui/StylePicker';
import { getPrimaryFontFamily } from '../../ui/fontPickerValue';
import { memoryPressureLevel, type DocxMemoryPressureLevel, type DocxMemoryStats } from '../memoryStats';
import type { DocxInput } from '@betteroffice/docx/utils';
import type { ImageLayoutTarget } from '@betteroffice/docx/docx';
import type { DocxCommandSelection, ResidentDocumentRead, ResidentDocumentReadValues } from '@betteroffice/docx/yrs';
import { createDocxCommandController, DocxCommandAdmissionError, type DocxCommandBinding } from '../../../commands/createDocxCommandStore';
import type { DocxCommandEnvironment } from '../../../commands/evaluate';
import type { DocxCommandArgs, DocxCommandId, DocxCommandResult, DocxImageTransform } from '../../../commands/types';
import type { DocxCommentInsertion, DocxDocumentChange, DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { PeerlessViewSnapshot } from './viewStore';
import { readWorkerSelection } from './clipboard';
import type { PeerlessInputIntent } from './inputEvents';
import { flushAdmittedInput, type PeerlessAdmissionResult, type PeerlessFormattingIntent } from './inputAdmission';
import type { DocxHostSearch } from '../hooks/useHostSearch';
import type { SelectionState } from '../types';
import { yrsSelectionState } from '../hooks/useSelectionTracker';
import { readPresentedReview, reviewDescribes, reviewMatches } from './workerReview';
import { workerRevisions, type WorkerRevisionState } from './workerRevisions';
import type { DocxProposalResult, DocxProposalSnapshot, ResidentReviewSnapshot, ResidentReviewTarget } from '@betteroffice/docx/yrs';
import type { Comment } from '@betteroffice/docx/types/content';

import { LAYOUT_POLL_MS, LAYOUT_REFUSALS, LAYOUT_WAIT_MS, VIEWER_LAYOUT_WAIT_MS, type WorkerPageExportResult } from '../internals/workerOpenExport';
import { PeerlessSyncApiError } from './contracts';
export { LEGACY_SYNC_API_PROP, PeerlessSyncApiError } from './contracts';

export class WorkerReadSupersededError extends Error {}

const OPEN_WAIT_MS = 120_000;
const MAX_WAITING_ROUNDS = 64;
const PROPOSAL_FAILURES = {
  'stale-version': true, 'missing-target': true, 'ambiguous-target': true, 'content-mismatch': true, 'overlapping-steps': true,
  'locked-target': true, 'read-only': true, 'tracked-revision-conflict': true, unsupported: true, 'invalid-step': true,
  'limit-exceeded': true, 'stale-preview': true, 'unknown-proposal': true, 'proposal-id-conflict': true,
} satisfies Record<Extract<DocxProposalResult, { ok: false }>['failure']['code'], true>;

const FORMATTING_COMMANDS = new Set<DocxCommandId>(['bold', 'italic', 'underline', 'fontFamily', 'fontSize', 'textColor', 'highlightColor',
  'strikethrough', 'superscript', 'subscript', 'clearFormatting', 'alignment', 'lineSpacing', 'setLtr', 'setRtl',
  'bulletList', 'numberedList', 'indent', 'outdent', 'paragraphStyle']);

export interface PeerlessWorkerQuery<K extends ResidentDocumentRead['kind'] = ResidentDocumentRead['kind']> {
  version: string;
  value: ResidentDocumentReadValues[K];
  selection: DocxCommandSelection | null | undefined;
}

interface Options extends Pick<DocxEditorProps, 'i18n' | 'theme' | 'mode' | 'readOnly' | 'allowHostProposals' | 'onDocumentChange' | 'onSelectionChange' | 'onCommentsChange' | 'onError' | 'onSave' | 'showFileOpen' | 'author' | 'onModeChange' | 'fontFamilies' | 'onFontsLoaded' | 'onFirstPagePainted' | 'onMemoryPressure' | 'memoryBudget'> {
  reviewTimeoutMs?: number;
  openTimeoutMs?: number;
  imagePropertiesAvailable?: boolean;
  image?: { wrap: ImageLayoutTarget | null } | null;
  onImageWrap?: (wrap: ImageLayoutTarget) => void | boolean | Promise<boolean | DocxCommandResult>;
  onImageTransform?: (action: DocxImageTransform) => void | boolean | Promise<boolean | DocxCommandResult>;
  zoom: number;
  tableGridlines?: boolean;
  isDark?: boolean;
  focus(): void;
  onZoom?(zoom: number): void;
  onTableGridlines?(visible: boolean): void;
  read: ResidentEngineWorkerClient['documentRead'];
  readAt?: ResidentEngineWorkerClient['documentReadAt'];
  layoutRequest?(): string | null;
  search?: DocxHostSearch;
  save?(): Promise<ArrayBuffer>;
  loadBuffer?(buffer: DocxInput): Promise<void>;
  loadDocument?(document: NonNullable<DocxEditorProps['document']>): Promise<void>;
  print?(preview?: boolean): Promise<void>;
  currentPage?(): number;
  scrollToPage?(number: number): void;
  scrollToParagraph?(paraId: string): Promise<boolean>;
  scrollToReviewTarget?(target: ResidentReviewTarget): Promise<boolean>;
  comments?(): Comment[];
  readPositionAtPoint?(x: number, y: number): Promise<import('../types').DocxPointPosition | null>;
  perform?(id: DocxCommandId, args: unknown): Promise<DocxCommandResult>;
  prepareFontFamily?(family: string): Promise<ResidentFontEnvironment | void>;
  pickImage?(insert: (image: Readonly<Record<string, unknown>>) => Promise<DocxCommandResult>): void;
  openSplitCell?(defaults: { initialRows: number; initialCols: number; minRows: number; minCols: number },
    apply: (rows: number, columns: number) => Promise<DocxCommandResult>): void;
  openTableProperties?(initialProps: { width?: number | null; widthType?: string | null; justification?: 'left' | 'center' | 'right' | null },
    apply: (props: { width?: number | null; widthType?: string | null; justification?: 'left' | 'center' | 'right' | null }) => Promise<DocxCommandResult>): void;
}

export function createPeerlessShellController(initial: Options) {
  let options = initial;
  let binding: PeerlessEditSessionBinding | null = null;
  let unsubscribe: (() => void) | undefined;
  let snapshot: { status: 'empty' | 'opening' | 'ready' | 'failed'; view: PeerlessViewSnapshot | null; error: Error | null; mode: NonNullable<DocxEditorProps['mode']>; preview: ResidentEngineWorkerFrame | null; previewPainted: boolean } =
    { status: 'empty', view: null, error: null, mode: initial.mode ?? 'editing', preview: null, previewPainted: false };
  const listeners = new Set<() => void>();
  const waits = new Set<() => void>();
  let closed = false;
  const changes = new Set<(change: DocxDocumentChange) => void>();
  const selections = new Set<(selection: SelectionState | null) => void>();
  const commands = createDocxCommandController();
  let translate = translation(options);
  let internalMode: NonNullable<DocxEditorProps['mode']> = 'editing';
  let fontRevision = 0;
  let fontFamilies: string[] = [];
  let metadata: { binding: PeerlessEditSessionBinding; version: number; value: ResidentDocumentReadValues['toolbarMetadata'] } | null = null;
  let firstPainted = false;
  let workerMemory: WasmModuleMemory[] | null = null;
  let pressure: DocxMemoryPressureLevel = 'normal';
  let reviewValue: ResidentReviewSnapshot | null = null;
  let reviewBinding: PeerlessEditSessionBinding | null = null;
  let reviewIdentity: PeerlessViewSnapshot['identity'] | null = null;
  let proposalTail: Promise<unknown> = Promise.resolve();
  let revisions: { binding: PeerlessEditSessionBinding; key: string; value: WorkerRevisionState } | null = null;
  let documentKey: unknown;
  let generation = 0;
  let borderPen: { style: string; size: number; color: { rgb: string } } = { style: 'single', size: 4, color: { rgb: '000000' } };
  const syncBorderPen = () => {
    const cellColor = snapshot.view?.editState?.tableContext?.cellBorderColor;
    const resolved = cellColor ? resolveColorToHex(cellColor, options.theme ?? undefined) : undefined;
    if (resolved) borderPen = { ...borderPen, color: { rgb: resolved } };
  };
  const waiting = new Set<(error: Error) => void>();
  const reviewReads = new WeakMap<PeerlessEditSessionBinding, { identity: PeerlessViewSnapshot['identity']; pending: Promise<ResidentReviewSnapshot> }>();
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
    fontRevision = 0; fontFamilies = []; metadata = null; revisions = null; firstPainted = false; workerMemory = null;
    unsubscribe?.(); binding = null; reviewValue = null; reviewBinding = null; reviewIdentity = null;
    batchResults.clear(); commentResults.clear(); appliedBatches.clear(); commentList = null; snapshot = { status, view: null, error, mode: mode(), preview: null, previewPainted: false }; publish();
    if (selected) selectionChanged(null);
  };
  const publish = () => { snapshot = { ...snapshot, mode: mode() }; updateAuthor(); commands.refresh(); for (const listener of listeners) listener();
    for (const wait of [...waits]) wait(); };
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
      sidebarOpen: false, sidebarControl: 'fixed', zoom: options.zoom, zoomWhileOpening: true, tableGridlines: options.tableGridlines === true, canOpen: options.showFileOpen !== false, canReportIssue: false,
      bodyStory: (view?.workerSelection?.focus.story ?? 'body') === 'body' || view?.workerSelection?.focus.story.startsWith('body:') === true, pendingInput: false,
      supportedCommands: new Set(['zoom', 'tableGridlines', 'open', 'save', 'print', 'undo', 'redo', 'find', 'replace', 'editingMode', 'pageSetup', 'watermark', 'imageProperties', 'imageWrap', 'imageTransform', 'insertTable', 'insertImage', 'tableAction',
        'insertPageBreak', 'insertSectionBreakNextPage', 'insertSectionBreakContinuous', 'insertLink', 'reviewAccept', 'reviewReject', 'reviewNext', 'reviewPrevious', ...FORMATTING_COMMANDS]),
      tableActions: new Set(['selectTable', 'selectRow', 'selectColumn', 'addRowAbove', 'addRowBelow', 'addColumnLeft',
        'addColumnRight', 'deleteRow', 'deleteColumn', 'deleteTable', 'mergeCells', 'splitCell', 'cellFillColor',
        'borderAll', 'borderOutside', 'borderInside', 'borderNone', 'borderTop', 'borderBottom', 'borderLeft', 'borderRight',
        'borderColor', 'borderWidth', 'cellBorder', 'applyTableStyle', 'tableProperties', 'openTableProperties']),
      canUndo: view?.editState?.canUndo ?? false, canRedo: view?.editState?.canRedo ?? false,
      selection: toolbar ? { context: toolbar, fontFamily: toolbar.fontFamily ?? null,
        fontSize: view?.editState?.effectiveFormatting?.fontSizePt ?? null } : null,
      table: view?.editState?.tableContext ? { isInTable: true, ...view.editState.tableContext,
        table: { attrs: { justification: view.editState.tableContext.justification } } } : null,
      image: options.image !== undefined ? options.image : options.imagePropertiesAvailable ? { wrap: null } : null, ...revisionEnvironment(view),
      styles: workerStyleOptions(metadata && metadata.binding === binding && metadata.version === view?.version ? metadata.value.styles : [], translate), fonts: [
        ...choices.map(font => ({ args: { family: getPrimaryFontFamily(font.fontFamily) || font.name }, label: font.name,
          preview: { fontFamily: font.fontFamily, group: font.category ?? 'other' } })),
        ...fontFamilies.filter(family => !choices.some(font => font.name.toLowerCase() === family.toLowerCase()))
          .map(family => ({ args: { family }, label: family, preview: { group: 'document' } })),
      ], translate,
    };
  };
  const revisionEnvironment = (view: PeerlessViewSnapshot | null | undefined) => {
    const state = revisions?.binding === binding ? revisions?.value : null;
    return { revisions: state?.navigation ?? [], revisionIds: state?.ids ?? new Set<string>(), currentRevisionId: view?.editState?.currentRevisionId ?? null };
  };
  const queryOrigins = new WeakMap<object, { binding: PeerlessEditSessionBinding; identity: PeerlessViewSnapshot['identity']; version: number }>();
  const capture = () => binding;
  const currentQuery = (result: object) => {
    const origin = queryOrigins.get(result);
    const view = binding?.viewStore.getSnapshot();
    return !!origin && origin.binding === binding && !!view && sameIdentity(origin.identity, view.identity) && origin.version === view.version;
  };
  const ready = () => {
    if (!binding || snapshot.status !== 'ready') throw new Error('Peer-less editor is not ready');
    return binding;
  };
  const replace = () => { generation += 1; proposalTail = Promise.resolve(); };
  const opened = (): Promise<PeerlessEditSessionBinding> => {
    if (binding && snapshot.status === 'ready') return Promise.resolve(binding);
    if (closed || snapshot.status !== 'opening') return Promise.reject(new Error('Peer-less editor is not ready'));
    if (waiting.size >= MAX_WAITING_ROUNDS) return Promise.reject(new Error('Too many calls are waiting for the document'));
    const awaited = generation;
    return new Promise((resolve, reject) => {
      const settle = (error: Error | null) => {
        clearTimeout(timer); waits.delete(check); waiting.delete(fail);
        if (error) reject(error); else resolve(binding!);
      };
      const fail = (error: Error) => settle(error);
      const check = () => {
        if (closed) settle(new Error('Editor was closed'));
        else if (snapshot.error) settle(snapshot.error);
        else if (generation !== awaited) settle(new Error('Document was replaced'));
        else if (binding && snapshot.status === 'ready') settle(null);
      };
      const timer = setTimeout(() => settle(new Error('The document did not open in time')), options.openTimeoutMs ?? OPEN_WAIT_MS);
      waits.add(check); waiting.add(fail); check();
    });
  };
  const flush = async () => {
    const current = ready();
    await flushAdmittedInput(current);
    if (current !== binding) throw new WorkerReadSupersededError('Document was replaced');
    return current;
  };
  const read = async <K extends ResidentDocumentRead['kind']>(request: ResidentDocumentRead & { kind: K }): Promise<ResidentDocumentReadValues[K]> => {
    const current = await flush();
    const view = current.viewStore.getSnapshot();
    const result = await options.read<K>(request);
    if (current !== binding) throw new WorkerReadSupersededError('Document was replaced');
    if (view.version !== current.viewStore.getSnapshot().version || !sameIdentity(view.identity, current.viewStore.getSnapshot().identity)) throw new WorkerReadSupersededError('Worker read was superseded');
    return result.value;
  };
  const readAt = async <K extends ResidentDocumentRead['kind']>(request: ResidentDocumentRead & { kind: K }, token: string): Promise<ResidentDocumentReadValues[K]> => {
    const current = await flush();
    const view = current.viewStore.getSnapshot();
    if (!options.readAt) throw new Error('Versioned worker reads are unavailable');
    const result = await options.readAt<K>(request, token);
    const latest = current.viewStore.getSnapshot();
    if (current !== binding) throw new WorkerReadSupersededError('Document was replaced');
    if (result.status !== 'ok' || result.version !== token || latest.version !== view.version || !sameIdentity(latest.identity, view.identity)) throw new WorkerReadSupersededError('Worker read was superseded');
    return result.value;
  };
  const subscribe = (listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener); }; };
  const currentReview = (matches: typeof reviewMatches) => binding && binding === reviewBinding && reviewIdentity &&
    sameIdentity(reviewIdentity, binding.viewStore.getSnapshot().identity) && reviewValue && matches(binding.viewStore.getSnapshot(), reviewValue) ? reviewValue : null;
  const reviewSnapshot = () => currentReview(reviewMatches);
  /** The last review read while it still describes the presented document and preview, whichever frame shows them. */
  const describedReview = () => currentReview(reviewDescribes);
  const readReview = (current: PeerlessEditSessionBinding, timeoutMs = options.reviewTimeoutMs ?? 30000): Promise<ResidentReviewSnapshot> => {
    const identity = { ...current.viewStore.getSnapshot().identity };
    const pending = reviewReads.get(current);
    if (pending && sameIdentity(pending.identity, identity)) return pending.pending;
    const deadline = Date.now() + timeoutMs;
    const request = (async () => {
      while (Date.now() < deadline) {
        const value = await readPresentedReview({ binding: current, current: () => binding, subscribe,
          read: options.read, readAt: options.readAt, timeoutMs: Math.max(0, deadline - Date.now()) });
        if (binding !== current || !sameIdentity(identity, current.viewStore.getSnapshot().identity)) throw new Error('Document was replaced');
        if (!reviewMatches(current.viewStore.getSnapshot(), value)) continue;
        reviewBinding = current; reviewIdentity = { ...current.viewStore.getSnapshot().identity }; reviewValue = value;
        rememberComments(current, current.viewStore.getSnapshot(), value.uiComments); publish();
        if (reviewMatches(current.viewStore.getSnapshot(), value)) return value;
      }
      throw new Error('Review snapshot timed out');
    })().finally(() => { if (reviewReads.get(current)?.pending === request) reviewReads.delete(current); });
    reviewReads.set(current, { identity, pending: request });
    return request;
  };
  const presentedProposals = (current: PeerlessEditSessionBinding): Promise<DocxProposalSnapshot> => {
    const deadline = Date.now() + (options.reviewTimeoutMs ?? 30000);
    let timer: ReturnType<typeof setTimeout>;
    const expired = new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new Error('Review snapshot timed out')), Math.max(0, deadline - Date.now())); });
    const select = async () => {
      while (Date.now() < deadline) {
        const value = await readReview(current, Math.max(0, deadline - Date.now()));
        if (binding !== current) throw new Error('Document was replaced');
        if (reviewMatches(current.viewStore.getSnapshot(), value)) return structuredClone(value.proposals);
      }
      throw new Error('Review snapshot timed out');
    };
    return Promise.race([select(), expired]).finally(() => { clearTimeout(timer!); });
  };
  const released = (result: PeerlessAdmissionResult) => result.status === 'held' ? result.released.catch(() => null) : Promise.resolve(result);
  type ProposalIntent = Extract<PeerlessInputIntent, { kind: 'propose' | 'setProposalStates' | 'withdrawProposals' }>;
  const proposalFailure = (current: PeerlessEditSessionBinding, intent: ProposalIntent, code: string): DocxProposalResult => {
    const view = current.viewStore.getSnapshot();
    const version = view.editState?.review?.version ?? view.editState?.documentVersion ?? view.geometry?.documentVersion ?? intent.payload.expectVersion;
    const ids = intent.kind === 'propose' ? intent.payload.proposals.map(input => input.id) : intent.kind === 'setProposalStates' ? intent.payload.changes.map(change => change.id) : intent.payload.ids;
    const known = Object.hasOwn(PROPOSAL_FAILURES, code);
    const message = code === 'read-only' ? 'The editor is read-only' : code === 'stale-version' ? 'the document changed since the expected version was read' : known ? undefined : code;
    const failure = (known ? code : code === 'invalid-payload' ? 'invalid-step' : 'unsupported') as keyof typeof PROPOSAL_FAILURES;
    return { ok: false, version, failure: { code: failure, ...(ids.length === 1 ? { proposalId: ids[0] } : {}), ...(message ? { message } : {}) } as Extract<DocxProposalResult, { ok: false }>['failure'] };
  };
  const proposalRound = <T>(operation: (current: PeerlessEditSessionBinding) => Promise<T>): Promise<T> => {
    const pending = Promise.all([opened(), proposalTail]).then(async ([current]) => {
      if (binding !== current) throw new Error('Document was replaced');
      await flush();
      return operation(current);
    });
    proposalTail = pending.catch(() => {});
    return pending;
  };
  const mutateProposal = (intent: ProposalIntent): Promise<DocxProposalResult> => {
    const captured = structuredClone(intent);
    return proposalRound(async current => {
      const env = environment();
      if ((env.readOnly || env.mode === 'viewing') && !options.allowHostProposals) return proposalFailure(current, captured, 'read-only');
      let admitted = current.commands.submit(captured, { fence: true, hostProposal: options.allowHostProposals === true });
      try {
        while (admitted.status === 'held') admitted = await admitted.released;
      } catch (error) {
        if (binding !== current) throw new Error('Document was replaced');
        throw error;
      }
      if (binding !== current) throw new Error('Document was replaced');
      if (admitted.status === 'ignored') return proposalFailure(current, captured, admitted.reason);
      const outcome = await admitted.ticket.ack;
      if (binding !== current) throw new Error('Document was replaced');
      if (outcome.status !== 'applied' && outcome.status !== 'noop') {
        if (!outcome.reason) throw new Error('The worker did not confirm the proposal command');
        return proposalFailure(current, captured, outcome.reason);
      }
      return { ok: true, snapshot: await presentedProposals(current) };
    });
  };
  const settleTicket = async (current: PeerlessEditSessionBinding, held: PeerlessAdmissionResult): Promise<{ result: DocxCommandResult; seq?: number }> => {
    const replaced = { result: { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } } } as const;
    const admitted = await released(held);
    if (current !== binding) return replaced;
    if (admitted?.status !== 'submitted') return { result: { ok: false, failure: { code: 'input-failed', message: 'Worker command was not admitted.' } } };
    const outcome = await admitted.ticket.ack;
    if (current !== binding) return replaced;
    if (outcome.status === 'applied' || outcome.status === 'noop') return { result: { ok: true, status: outcome.status === 'applied' ? 'executed' : 'noop' }, seq: outcome.seq };
    const code = outcome.reason === 'read-only' ? 'read-only' : outcome.reason === 'revision-required' ? 'revision-required'
      : outcome.reason === 'unknown-revision' ? 'revision-not-found' : 'command-failed';
    return { result: { ok: false, failure: { code, message: outcome.reason ?? 'Worker command failed.' } }, seq: outcome.seq };
  };
  const settleCommand = async (current: PeerlessEditSessionBinding, held: PeerlessAdmissionResult): Promise<DocxCommandResult> =>
    (await settleTicket(current, held)).result;
  const submit = (intent: PeerlessInputIntent): Promise<DocxCommandResult> => {
    const current = ready();
    return settleCommand(current, current.commands.submit(intent, { fence: true }));
  };
  const format = (intent: PeerlessFormattingIntent): Promise<DocxCommandResult> => {
    const current = ready();
    return settleCommand(current, current.commands.format(intent, { fence: true }));
  };
  const query = async <K extends ResidentDocumentRead['kind']>(request: ResidentDocumentRead & { kind: K }): Promise<PeerlessWorkerQuery<K>> => {
    const current = await flush();
    const view = current.viewStore.getSnapshot();
    const native = await options.read({ kind: 'readParagraphs', request: { view: 'accepted', paraIds: [] } });
    if (current !== binding || !sameIdentity(view.identity, current.viewStore.getSnapshot().identity) || view.throughSeq !== current.viewStore.getSnapshot().throughSeq) throw new Error('Worker read was superseded');
    const value = await readAt<K>(request, native.version);
    if (view.throughSeq !== current.viewStore.getSnapshot().throughSeq) throw new Error('Worker read was superseded');
    if (request.kind === 'reviewSnapshot') rememberComments(current, view, (value as ResidentDocumentReadValues['reviewSnapshot'])?.uiComments);
    const result = { version: native.version, value, selection: structuredClone(view.workerSelection) };
    queryOrigins.set(result, { binding: current, identity: view.identity, version: view.version });
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
    ordered: (id) => !FORMATTING_COMMANDS.has(id) && !['zoom', 'tableGridlines', 'open', 'print', 'insertTable', 'reviewAccept', 'reviewReject'].includes(id),
    admit: async (operation) => {
      const current = binding;
      if (current && snapshot.status === 'ready') {
        await flushAdmittedInput(current);
        if (current !== binding) throw new DocxCommandAdmissionError('document-replaced');
      }
      return operation();
    },
    async perform(id, args) {
      if (id === 'editingMode') {
        const next = (args as { mode: NonNullable<DocxEditorProps['mode']> }).mode;
        if (options.mode) { emit(options.onModeChange, next); return { ok: true, status: 'requested' }; }
        internalMode = next; publish(); emit(options.onModeChange, next);
        return { ok: true, status: 'executed' };
      }
      if (id === 'undo' || id === 'redo') return submit({ kind: id, payload: {} });
      if (id === 'reviewAccept' || id === 'reviewReject') {
        const revisionId = (args as DocxCommandArgs['reviewAccept'])?.revisionId;
        return submit({ kind: id === 'reviewAccept' ? 'acceptRevision' : 'rejectRevision', payload: revisionId ? { revisionId } : {} });
      }
      if (id === 'reviewNext' || id === 'reviewPrevious') {
        const target = await query({ kind: 'adjacentRevision', selection: ready().viewStore.getSnapshot().workerSelection ?? null, direction: id === 'reviewNext' ? 'next' : 'previous' });
        return target.value && await revealReviewTarget({ kind: 'revision', revisionId: target.value.revisionId }) ? { ok: true, status: 'executed' }
          : { ok: false, failure: { code: 'command-failed', message: 'Worker command failed.' } };
      }
      if (id === 'strikethrough') return format({ kind: 'toggleStrike', payload: {} });
      if (id === 'superscript' || id === 'subscript') return format({ kind: 'toggleMark', payload: { mark: { type: id } } });
      if (id === 'clearFormatting') return format({ kind: 'clearFormatting', payload: {} });
      if (id === 'alignment') return format({ kind: 'setParagraphAttrs', payload: { attrs: { alignment: (args as DocxCommandArgs['alignment']).value } } });
      if (id === 'lineSpacing') return format({ kind: 'setParagraphAttrs', payload: { attrs: { lineSpacing: (args as DocxCommandArgs['lineSpacing']).value, lineSpacingRule: 'auto' } } });
      if (id === 'setLtr' || id === 'setRtl') return format({ kind: 'setParagraphAttrs', payload: { attrs: { bidi: id === 'setRtl' } } });
      if (id === 'bulletList' || id === 'numberedList') return format({ kind: 'toggleList', payload: { kind: id === 'bulletList' ? 'bullet' : 'numbered' } });
      if (id === 'indent' || id === 'outdent') return format({ kind: 'changeIndent', payload: { direction: id === 'indent' ? 'increase' : 'decrease' } });
      if (id === 'paragraphStyle') return format({ kind: 'applyParagraphStyle', payload: { styleId: (args as DocxCommandArgs['paragraphStyle']).styleId } });
      if (id === 'insertTable') return submit({ kind: 'insertTable', payload: args as DocxCommandArgs['insertTable'] });
      if (id === 'tableAction') {
        const key = typeof args === 'string' ? args : (args as { type?: string } | undefined)?.type;
        const target: 'table' | 'row' | 'column' | null = key === 'selectTable' ? 'table' : key === 'selectRow' ? 'row' : key === 'selectColumn' ? 'column' : null;
        if (target) {
          const current = ready();
          return settleCommand(current, current.commands.submit({ kind: 'selectTableCells', payload: { target } },
            { source: 'selection', fence: true }));
        }
        switch (key) {
          case 'addRowAbove': return submit({ kind: 'insertTableRow', payload: { side: 'above' } });
          case 'addRowBelow': return submit({ kind: 'insertTableRow', payload: { side: 'below' } });
          case 'addColumnLeft': return submit({ kind: 'insertTableColumn', payload: { side: 'left' } });
          case 'addColumnRight': return submit({ kind: 'insertTableColumn', payload: { side: 'right' } });
          case 'deleteRow': return submit({ kind: 'deleteTableRow', payload: {} });
          case 'deleteColumn': return submit({ kind: 'deleteTableColumn', payload: {} });
          case 'deleteTable': return submit({ kind: 'deleteTable', payload: {} });
          case 'mergeCells': return submit({ kind: 'mergeTableCells', payload: {} });
          case 'cellFillColor': {
            const color = (args as { color?: string | null }).color;
            const normalized = color === null || color === undefined || color === '' ? null : color.replace(/^#/, '');
            return submit({ kind: 'setTableCellShading', payload: { color: normalized } });
          }
          case 'borderAll': case 'borderOutside': case 'borderInside': case 'borderNone':
          case 'borderTop': case 'borderBottom': case 'borderLeft': case 'borderRight': {
            const pen = borderPen;
            const all = (border: typeof pen): YrsCellBorders => ({ top: border, bottom: border, left: border, right: border, insideH: border, insideV: border });
            const borders: YrsCellBorders = key === 'borderAll' ? all(pen)
              : key === 'borderOutside' ? { top: pen, bottom: pen, left: pen, right: pen }
              : key === 'borderInside' ? { insideH: pen, insideV: pen }
              : key === 'borderNone' ? all({ style: 'none', size: 0, color: { rgb: '000000' } })
              : { [key.slice('border'.length).toLowerCase()]: pen } as YrsCellBorders;
            return submit({ kind: 'setTableBorders', payload: { borders } });
          }
          case 'borderColor': {
            borderPen = { ...borderPen, color: { rgb: (args as { color: string }).color.replace(/^#/, '') } };
            const pen = borderPen;
            return submit({ kind: 'setTableBorders', payload: { borders: { top: pen, bottom: pen, left: pen, right: pen, insideH: pen, insideV: pen } } });
          }
          case 'borderWidth': {
            borderPen = { ...borderPen, size: (args as { size: number }).size };
            const pen = borderPen;
            return submit({ kind: 'setTableBorders', payload: { borders: { top: pen, bottom: pen, left: pen, right: pen, insideH: pen, insideV: pen } } });
          }
          case 'cellBorder': {
            const action = args as { side: string; style: string; size: number; color: string };
            const border = { style: action.style, size: action.size, color: { rgb: action.color.replace(/^#/, '') } };
            const borders: YrsCellBorders = action.side === 'all'
              ? { top: border, bottom: border, left: border, right: border, insideH: border, insideV: border }
              : { [action.side]: border } as YrsCellBorders;
            return submit({ kind: 'setTableBorders', payload: { borders } });
          }
          case 'applyTableStyle': {
            const preset = getBuiltinTableStyle((args as { styleId: string }).styleId);
            if (!preset?.tableBorders) return { ok: true, status: 'noop' };
            return submit({ kind: 'setTableBorders', payload: { borders: preset.tableBorders } });
          }
          case 'splitCell': {
            if (!options.openSplitCell) break;
            const tableContext = snapshot.view?.editState?.tableContext;
            const defaults = computeSplitDialogDefaults(tableContext?.rowSpan ?? 1, tableContext?.columnSpan ?? 1);
            const deferred = commands.defer('tableAction', args as DocxCommandArgs['tableAction'], 'selection');
            options.openSplitCell(defaults, (rows, columns) =>
              deferred.complete(() => submit({ kind: 'splitTableCell', payload: { rows, columns } })));
            return { ok: true, status: 'opened' };
          }
          case 'tableProperties': {
            const props = (args as { props: { width?: number | null; widthType?: string | null; justification?: 'left' | 'center' | 'right' | null } }).props;
            return submit({ kind: 'setTableProperties', payload: { op: 'patch', patch: props } });
          }
          case 'openTableProperties': {
            if (!options.openTableProperties) break;
            const deferred = commands.defer('tableAction', args as DocxCommandArgs['tableAction'], 'selection');
            const selected = snapshot.view?.workerSelection;
            if (!selected) return { ok: false, failure: { code: 'command-failed', message: 'Table selection is unavailable.' } };
            const originBinding = binding;
            const originMode = mode();
            const opened = await query({ kind: 'properties', target: { kind: 'tableSelection', selection: selected } });
            if (binding !== originBinding) return { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } };
            if (mode() !== originMode || !sameSelection(selected, snapshot.view?.workerSelection)) {
              return { ok: false, failure: { code: 'target-changed', message: 'Table selection changed.' } };
            }
            if (opened.value?.kind !== 'table') return { ok: false, failure: { code: 'command-failed', message: 'Table properties are unavailable.' } };
            const captured = opened.selection ?? selected;
            options.openTableProperties(opened.value.properties, (props) =>
              deferred.complete(() => applyQuery(opened, { kind: 'setTableProperties',
                payload: { op: 'patch', patch: props, expectVersion: opened.version, selection: captured } })));
            return { ok: true, status: 'opened' };
          }
          default: break;
        }
      }
      if (id === 'insertImage' && options.pickImage) {
        const picker = commands.defer('insertImage', null, 'selection');
        options.pickImage(image => picker.complete(() => submit({ kind: 'insertImage', payload: { image } })));
        return { ok: true, status: 'opened' };
      }
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
      if (id === 'highlightColor') {
        const color = (args as DocxCommandArgs['highlightColor']).color;
        return format({ kind: 'formatRange', payload: { delta: { highlight: color && color !== 'none' ? color : null } } });
      }
      if (id === 'insertPageBreak') return submit({ kind: 'insertPageBreak', payload: {} });
      if (id === 'insertSectionBreakNextPage' || id === 'insertSectionBreakContinuous') {
        return submit({ kind: 'insertSectionBreak', payload: { type: id === 'insertSectionBreakNextPage' ? 'nextPage' : 'continuous' } });
      }
      if (id === 'textColor') {
        const current = ready();
        const color = (args as DocxCommandArgs['textColor']).color;
        const value = typeof color === 'object' ? { ...color } : null;
        const nativeColor = value && ('rgb' in value ? { rgb: value.rgb } : { themeColor: value.themeColor });
        const delta = { color: nativeColor, other: { textColor: value } };
        return settleCommand(current, current.commands.format({ kind: 'formatRange', payload: { delta } }, { fence: true }));
      }
      if (id === 'zoom') {
        options = { ...options, zoom: (args as { scale: number }).scale };
        options.onZoom?.(options.zoom);
        publish();
        return { ok: true, status: 'executed' };
      }
      if (id === 'tableGridlines') {
        options = { ...options, tableGridlines: options.tableGridlines !== true };
        options.onTableGridlines?.(options.tableGridlines === true);
        publish();
        return { ok: true, status: 'executed' };
      }
      if (id === 'imageWrap') {
        const result = await options.onImageWrap?.((args as DocxCommandArgs['imageWrap']).wrap);
        if (result && typeof result === 'object' && 'ok' in result) return result;
        return result === false ? { ok: false, failure: { code: 'command-failed', message: 'Image wrap is unavailable.' } } : { ok: true, status: 'executed' };
      }
      if (id === 'imageTransform') {
        const result = await options.onImageTransform?.((args as DocxCommandArgs['imageTransform']).action);
        if (result && typeof result === 'object' && 'ok' in result) return result;
        return result === false ? { ok: false, failure: { code: 'command-failed', message: 'Image transform is unavailable.' } } : { ok: true, status: 'executed' };
      }
      return options.perform?.(id, args) ?? { ok: false, failure: { code: 'unsupported-command', message: 'Worker command is unavailable.' } };
    },
    capture: (target) => {
      const current = capture();
      return current ? { document: current, identity: { ...current.viewStore.getSnapshot().identity },
        selection: target === 'selection' ? structuredClone(current.viewStore.getSnapshot().workerSelection) : undefined } : null;
    },
    resume: (origin) => {
      const captured = origin as typeof origin & { identity: PeerlessViewSnapshot['identity']; selection?: DocxCommandSelection | null };
      if (captured.document !== binding || !binding || !sameIdentity(captured.identity, binding.viewStore.getSnapshot().identity)) return 'document-replaced';
      return captured.selection && !sameSelection(captured.selection, binding.viewStore.getSnapshot().workerSelection) ? 'target-changed' : null;
    },
    chrome: () => ({ i18n: options.i18n, isDark: options.isDark ?? false, theme: options.theme ?? null }),
    focusEditor: () => options.focus(),
  };
  commands.attach(commandBinding);
  const batchResults = new Map<number, DocxEditResult>();
  const appliedBatches = new Map<{ readonly seq: number | undefined }, { binding: PeerlessEditSessionBinding; applied: (view: PeerlessViewSnapshot) => void }>();
  const commentResults = new Map<number, { id?: number; comments: Comment[] }>();
  let commentTail: Promise<unknown> = Promise.resolve();
  let commentList: { binding: PeerlessEditSessionBinding; identity: PeerlessViewSnapshot['identity']; version: number; comments: Comment[] } | null = null;
  const commentsMatch = (view: PeerlessViewSnapshot) => commentList?.binding === binding &&
    sameIdentity(commentList.identity, view.identity) && commentList.version <= view.version;
  const rememberComments = (current: PeerlessEditSessionBinding, view: PeerlessViewSnapshot, comments: Comment[] | undefined) => {
    const latest = current.viewStore.getSnapshot();
    if (!comments || current !== binding || !sameIdentity(view.identity, latest.identity) || view.version !== latest.version) return;
    if (commentList?.binding === current && sameIdentity(commentList.identity, view.identity) && commentList.version > view.version) return;
    commentList = { binding: current, identity: { ...view.identity }, version: view.version, comments };
  };
  const flushComments = async () => {
    const current = ready();
    await commentTail;
    if (current !== binding) throw new WorkerReadSupersededError('Document was replaced');
    return flush();
  };
  const commentRound = <T>(refused: () => T, operation: (current: PeerlessEditSessionBinding) => Promise<T>, failed: (error: Error) => T = refused): Promise<T> => {
    const origin = binding;
    const pending = commentTail.then(async () => {
      const env = environment();
      if (!origin || origin !== binding || env.readOnly || env.mode === 'viewing') return refused();
      try {
        const current = await flush();
        return origin === current ? await operation(current) : refused();
      } catch (error) {
        const reported = error instanceof Error ? error : new Error(String(error));
        emit(options.onError, reported);
        return failed(reported);
      }
    });
    commentTail = pending.catch(() => null);
    return pending;
  };
  const commentCommand = async (current: PeerlessEditSessionBinding, intent: PeerlessInputIntent) => {
    const { result, seq } = await settleTicket(current, current.commands.submit(intent, { fence: true }));
    const comment = seq === undefined ? undefined : commentResults.get(seq);
    if (seq !== undefined) commentResults.delete(seq);
    if (result.ok && result.status === 'executed' && comment && binding === current) emit(options.onCommentsChange, comment.comments);
    return { result, comment };
  };
  const addComment = async (current: PeerlessEditSessionBinding, payload: Extract<PeerlessInputIntent, { kind: 'addComment' }>['payload']) => {
    const { result, comment } = await commentCommand(current, { kind: 'addComment', payload });
    return result.ok && result.status === 'executed' && binding === current ? comment?.id ?? null : null;
  };
  const insertComment = (request: DocxCommentInsertion): Promise<number | null> => commentRound(() => null, async current => {
    const target = await query({ kind: 'commentInsertionRange', paraId: request.paraId,
      ...(request.search === undefined ? {} : { search: request.search }) });
    if (!target.value || 'status' in target.value || current !== binding) return null;
    return addComment(current, { ranges: [target.value], author: request.author, date: new Date().toISOString(), body: request.text, expectVersion: target.version });
  });
  /** Adds a reply anchored where its parent is; null when the parent does not exist or the editor is read-only. */
  const replyToComment = (commentId: number, text: string, author: string): Promise<number | null> => commentRound(() => null, async current => {
    const target = await query({ kind: 'reviewTarget', target: { kind: 'comment', id: commentId } });
    if (!target.value?.ranges.length || current !== binding) return null;
    return addComment(current, { ranges: target.value.ranges, author, date: new Date().toISOString(), body: text,
      parentId: target.value.id, expectVersion: target.version });
  });
  const changeComment = (commentId: number, intent: (id: string, expectVersion: string) => PeerlessInputIntent): Promise<DocxCommandResult> => {
    const refused = (): DocxCommandResult => {
      const env = environment();
      return { ok: false, failure: env.mode === 'viewing' ? { code: 'viewing-mode', message: 'The editor is viewing.' }
        : env.readOnly ? { code: 'read-only', message: 'The editor is read-only.' } : { code: 'document-replaced', message: 'Document was replaced.' } };
    };
    return commentRound(refused, async current => {
      const target = await query({ kind: 'reviewTarget', target: { kind: 'comment', id: commentId } });
      if (current !== binding) return { ok: false, failure: { code: 'document-replaced', message: 'Document was replaced.' } };
      if (!target.value) return { ok: false, failure: { code: 'command-failed', message: 'unknown-comment' } };
      return (await commentCommand(current, intent(target.value.id, target.version))).result;
    }, error => ({ ok: false, failure: { code: 'command-failed', message: error.message } }));
  };
  /** Marks a comment done or open again through the worker command queue. */
  const setCommentDone = (commentId: number, done: boolean) =>
    changeComment(commentId, (id, expectVersion) => ({ kind: 'updateComment', payload: { id, patch: { done }, expectVersion } }));
  /** Removes a comment through the worker command queue. */
  const removeComment = (commentId: number) =>
    changeComment(commentId, (id, expectVersion) => ({ kind: 'removeComment', payload: { id, expectVersion } }));
  const getComments = (): Comment[] => snapshot.status !== 'ready' ? [] :
    structuredClone(snapshot.view && commentsMatch(snapshot.view) ? commentList!.comments : options.comments?.() ?? []);
  /** Selects a comment's or revision's first anchored range and scrolls it into view; false when it does not exist. */
  const revealReviewTarget = (target: ResidentReviewTarget): Promise<boolean> =>
    options.scrollToReviewTarget?.(target) ?? Promise.resolve(false);
  const batchPolicy = (request: DocxEditRequest) => {
    const env = environment();
    if (env.readOnly || env.mode === 'viewing') return { ok: false as const, version: request.expectVersion, failure: { code: 'read-only' as const, message: 'The editor is read-only' } };
    const stepIndex = request.steps.findIndex(step => !('suggest' in step && step.suggest));
    if (env.mode === 'suggesting' && stepIndex >= 0) return { ok: false as const, version: request.expectVersion, failure: { code: 'invalid-step' as const, stepIndex, message: 'Suggesting mode requires suggest metadata' } };
    return null;
  };
  /** `applied` runs with the view that first shows the batch applied, before that view is published. */
  const applyEdits = async (request: DocxEditRequest, authorize: () => unknown = () => null, applied?: (view: PeerlessViewSnapshot) => void) => {
    const current = await flush();
    const denied = authorize();
    if (denied) return denied as DocxEditResult;
    const policy = batchPolicy(request);
    if (policy) return policy;
    const admitted = current.commands.submit({ kind: 'applyEdits', payload: { request } }, { fence: true });
    if (admitted.status !== 'submitted') return { ok: false as const, version: request.expectVersion, failure: { code: 'read-only' as const, message: 'Worker batch was not admitted' } };
    if (applied) appliedBatches.set(admitted.ticket, { binding: current, applied });
    const outcome = await admitted.ticket.ack.finally(() => appliedBatches.delete(admitted.ticket));
    if (binding !== current) throw new Error('Document was replaced');
    const receipt = outcome.seq === undefined ? undefined : batchResults.get(outcome.seq);
    if (outcome.seq !== undefined) batchResults.delete(outcome.seq);
    if (receipt) return receipt;
    return { ok: false as const, version: request.expectVersion, failure: { code: (outcome.reason ?? 'unsupported') as import('@betteroffice/docx/yrs').DocxEditFailureCode, message: outcome.reason ?? 'Worker batch failed' } };
  };
  const whenLayoutComplete: DocxEditorRef['whenLayoutComplete'] = ({ timeoutMs = 30000 } = {}) => new Promise<number>((resolve, reject) => {
    let current: PeerlessEditSessionBinding | null = null;
    let flushed = false;
    let stop: (() => void) | undefined;
    const settle = (error: Error | null) => {
      clearTimeout(timeout); waits.delete(check); stop?.(); stop = undefined;
      if (error) reject(error); else resolve(current?.viewStore.getSnapshot().geometry?.pages.length ?? 0);
    };
    const check = () => {
      if (closed) return settle(new Error('Editor was closed'));
      if (snapshot.error) return settle(snapshot.error);
      if (current) {
        if (binding !== current) return settle(new Error('Document was replaced'));
        const view = current.viewStore.getSnapshot();
        if (flushed && view.geometry && view.presentedThroughSeq >= view.throughSeq && !view.geometry.layoutProvisional) settle(null);
        return;
      }
      if (snapshot.status !== 'ready' || !binding) return;
      const opened = binding;
      current = opened;
      void flushAdmittedInput(opened).then(() => {
        if (!waits.has(check)) return;
        flushed = true; stop = opened.viewStore.subscribe(check); check();
      }, (error: unknown) => { if (waits.has(check)) settle(error instanceof Error ? error : new Error(String(error))); });
    };
    const timeout = setTimeout(() => settle(new Error('Layout timed out')), timeoutMs);
    waits.add(check);
    check();
  });
  const exportStructuredWithPages: DocxEditorRef['exportStructuredWithPages'] = async (config) => {
    const current = await opened();
    await flush();
    const view = current.viewStore.getSnapshot();
    let request: string | null = null;
    let superseded = false;
    const assertCurrent = () => { if (binding !== current) throw new WorkerReadSupersededError('Document was replaced'); };
    const moved = (): WorkerPageExportResult => {
      superseded = true;
      return { ok: false, version: view.editState?.documentVersion ?? view.geometry?.documentVersion ?? '',
        failure: { code: 'stale-document', target: null, message: 'The document changed while exporting.' } };
    };
    const attempt = async (): Promise<WorkerPageExportResult> => {
      assertCurrent();
      request = options.layoutRequest?.() ?? null;
      if (request === null) return { ok: false, version: view.editState?.documentVersion ?? '',
        failure: { code: 'layout-unavailable', target: null, message: 'The fonts this document uses are not loaded yet.' } };
      try {
        const value = await read({ kind: 'exportStructuredWithPages', options: config, currentRequest: request, workerFonts: true });
        assertCurrent();
        return JSON.parse(value) as WorkerPageExportResult;
      } catch (error) {
        assertCurrent();
        if (error instanceof WorkerReadSupersededError) return moved();
        throw error;
      }
    };
    const showsMarkup = () => {
      if (request === null) return false;
      const preview = (JSON.parse(request) as { renderEnv?: { revisionPreview?: object } }).renderEnv?.revisionPreview;
      return !preview || Object.keys(preview).length === 0;
    };
    let result = await attempt();
    if (config.expectLayoutVersion !== undefined || superseded) return result;
    if (!result.ok && (LAYOUT_REFUSALS.has(result.failure.code) ||
      (result.failure.code === 'unsupported-revision-layout' && showsMarkup()))) {
      await whenLayoutComplete({ timeoutMs: VIEWER_LAYOUT_WAIT_MS }).catch(() => {});
      assertCurrent();
      result = await attempt();
      if (superseded) return result;
    }
    const deadline = Date.now() + LAYOUT_WAIT_MS;
    while (!result.ok && LAYOUT_REFUSALS.has(result.failure.code) && Date.now() < deadline) {
      await new Promise(resolve => setTimeout(resolve, LAYOUT_POLL_MS));
      result = await attempt();
      if (superseded) return result;
    }
    return result;
  };
  const implemented = {
    commands: commands.store,
    insertComment,
    insertCommentReply: replyToComment,
    resolveComment: (commentId) => { void setCommentDone(commentId, true); },
    getComments,
    proposeChanges: request => mutateProposal({ kind: 'propose', payload: request }),
    setProposalStates: request => mutateProposal({ kind: 'setProposalStates', payload: request }),
    withdrawProposals: request => mutateProposal({ kind: 'withdrawProposals', payload: request }),
    getProposals: () => proposalRound(current => presentedProposals(current)),
    validateEdits: async (request) => { await flush(); return batchPolicy(request) ?? read({ kind: 'validateEdits', request }); },
    applyEdits: (request) => applyEdits(request),
    readPositionAtPoint: (x, y) => options.readPositionAtPoint?.(x, y) ?? Promise.resolve(null),
    whenLayoutComplete,
    exportStructuredWithPages,
    flushPendingInput: async () => { await flushComments(); },
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
    save: async () => { const current = await flushComments(); const result = await options.save?.();
      if (current !== binding) throw new Error('Document was replaced'); if (result) options.onSave?.(result); return result ?? null; },
    loadDocumentBuffer: (buffer) => options.loadBuffer?.(buffer) ?? Promise.reject(new Error('Worker loading is unavailable')),
    loadDocument: (document) => { void options.loadDocument?.(document).catch(() => {}); },
    print: () => { void options.print?.().catch(() => {}); },
    openPrintPreview: () => { void options.print?.(true).catch(() => {}); },
    getCurrentPage: () => options.currentPage?.() ?? 1,
    getTotalPages: () => snapshot.view?.geometry?.pages.length ?? 0,
    scrollToPage: (number) => options.scrollToPage?.(number),
    scrollToParagraph: (paraId) => options.scrollToParagraph?.(paraId) ?? Promise.resolve(false),
    scrollToComment: (commentId) => Number.isSafeInteger(commentId) ? revealReviewTarget({ kind: 'comment', id: commentId }) : Promise.resolve(false),
    scrollToChange: (revisionId) => revealReviewTarget({ kind: 'revision', revisionId: String(revisionId) }),
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
    commands, ref, read, readAt, query, applyQuery, submit, navigateCell, environment, applyEdits, reviewSnapshot, describedReview, revealReviewTarget,
    replyToComment, setCommentDone, removeComment,
    review: async () => readReview(await flush()),
    getSnapshot: () => snapshot,
    subscribe,
    update(next: Options) { options = next; translate = translation(next); syncBorderPen();
      if (snapshot.mode !== mode()) publish(); else { updateAuthor(); commands.refresh(); }
      reportMemory();
    },
    opening(document?: unknown) {
      if (document === undefined || document !== documentKey) replace();
      documentKey = document;
      reset('opening');
    },
    preview(frame: ResidentEngineWorkerFrame) {
      if (snapshot.status !== 'opening') return;
      snapshot = { ...snapshot, preview: frame, previewPainted: false }; publish();
    },
    previewPresented(frame: ResidentEngineWorkerFrame) {
      if (snapshot.preview !== frame || snapshot.status !== 'opening') return;
      snapshot = { ...snapshot, previewPainted: true }; publish();
    },
    painted,
    stalled(error: Error) {
      if (snapshot.status === 'opening') for (const fail of [...waiting]) fail(error);
    },
    failOpening(error: Error) {
      if (snapshot.status === 'opening') { snapshot = { ...snapshot, error }; publish(); }
    },
    toolbarMetadata(result: PeerlessWorkerQuery<'toolbarMetadata'> | null) {
      const origin = result && queryOrigins.get(result);
      const view = binding?.viewStore.getSnapshot();
      if (!result) metadata = null;
      else if (origin && origin.binding === binding && view && sameIdentity(origin.identity, view.identity) && origin.version === view.version) {
        metadata = { binding: origin.binding, version: origin.version, value: result.value };
      }
      commands.refresh();
    },
    revisions(review: PeerlessWorkerQuery<'reviewSnapshot'> | null, sidebar: PeerlessWorkerQuery<'sidebar'> | null) {
      if (!review?.value || !currentQuery(review) || (sidebar && (!currentQuery(sidebar) || sidebar.version !== review.version))) return;
      const value = workerRevisions(review.value, sidebar?.value ?? null);
      if (!value) return;
      const key = JSON.stringify([[...value.ids], value.navigation, value.entries, [...value.anchors]]);
      if (revisions?.binding === binding && revisions.key === key) return;
      revisions = { binding: binding!, key, value };
      publish();
    },
    trackedChanges: () => revisions?.binding === binding ? revisions?.value ?? null : null,
    fontChoices(families: string[]) { fontFamilies = [...families]; publish(); },
    fontEnvironment(event: ResidentFontEnvironmentEvent) {
      if (event.phase !== 'ready' || event.revision <= fontRevision || snapshot.status === 'empty' || snapshot.status === 'failed') return;
      fontRevision = event.revision; fontFamilies = [...event.families]; publish(); emit(options.onFontsLoaded, undefined);
    },
    memory(samples: WasmModuleMemory[]) { workerMemory = samples.map(sample => ({ ...sample })); reportMemory(); },
    empty() { replace(); documentKey = undefined; reset('empty'); },
    fail(error: Error) { replace(); reset('failed', error); },
    attach(next: PeerlessEditSessionBinding) {
      if (next === binding) return;
      unsubscribe?.(); binding = next;
      const accept = () => {
        if (binding !== next) return;
        const view = next.viewStore.getSnapshot();
        if (!commentsMatch(view)) commentList = null;
        if (view.editState?.batch) {
          batchResults.set(view.editState.batch.seq, view.editState.batch.result);
          if (batchResults.size > 256) batchResults.delete(batchResults.keys().next().value!);
        }
        for (const comment of [view.editState?.comment, view.editState?.commentChange]) {
          if (!comment || comment === snapshot.view?.editState?.comment || comment === snapshot.view?.editState?.commentChange) continue;
          rememberComments(next, view, comment.comments);
          commentResults.set(comment.seq, comment);
          if (commentResults.size > 256) commentResults.delete(commentResults.keys().next().value!);
        }
        const previous = snapshot.view;
        const ready = view.presentedThroughSeq >= 0;
        snapshot = { ...snapshot, status: ready ? 'ready' : 'opening', view, error: null, mode: mode(),
          preview: ready ? null : snapshot.preview };
        syncBorderPen();
        const batch = view.editState?.batch;
        if (batch?.result.ok) {
          for (const [ticket, own] of appliedBatches) {
            if (own.binding !== next || ticket.seq !== batch.seq) continue;
            appliedBatches.delete(ticket);
            emit(own.applied, view);
          }
        }
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
      replace();
      reset('empty');
    },
    dispose() { unsubscribe?.(); binding = null; batchResults.clear(); commentResults.clear(); appliedBatches.clear(); closed = true; publish(); commands.detach(commandBinding); listeners.clear(); changes.clear(); selections.clear(); },
  };
}

function sameIdentity(first: PeerlessViewSnapshot['identity'], second: PeerlessViewSnapshot['identity']) {
  return first.session === second.session && first.incarnation === second.incarnation && first.epoch === second.epoch;
}

function sameSelection(first: DocxCommandSelection, second: DocxCommandSelection | null | undefined) {
  return second && (['anchor', 'focus'] as const).every(key => first[key].story === second[key].story &&
    first[key].paraId === second[key].paraId && first[key].offset === second[key].offset);
}

function workerStyleOptions(styles: ResidentDocumentReadValues['toolbarMetadata']['styles'], translate: DocxCommandEnvironment['translate']): DocxCommandEnvironment['styles'] {
  return (styles.length ? styles : DEFAULT_STYLES).map(style => {
    const nameKey = DEFAULT_STYLES.find(option => option.styleId === style.styleId)?.nameKey;
    const preview: NonNullable<DocxCommandEnvironment['styles'][number]['preview']> = {};
    if (style.fontSize != null) preview.fontSize = style.fontSize / 2;
    if (style.bold != null) preview.bold = style.bold;
    if (style.italic != null) preview.italic = style.italic;
    if (style.color != null) preview.color = style.color;
    return { args: { styleId: style.styleId }, label: nameKey ? translate(nameKey) : style.name,
      ...(Object.keys(preview).length ? { preview } : {}) };
  });
}

function translation(options: Options) {
  const strings = deepMerge(en as Record<string, unknown>, options.i18n as Record<string, unknown> | undefined) as LocaleStrings;
  return createT(strings, typeof options.i18n?._lang === 'string' ? options.i18n._lang : 'en');
}
