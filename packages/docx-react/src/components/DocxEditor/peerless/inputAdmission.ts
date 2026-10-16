import type { CommandTicket, PresentationStatus, SubmitOptions, TicketOutcome } from '../../../../../../shared/office-session/commands/types';
import type {
  DocxCommandSelection,
  DocxCommandPayloads,
  DocxWorkerEditSession,
} from '@betteroffice/docx/peerless';
import type { YrsAuthor } from '@betteroffice/docx/yrs';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import { hasImplicitVerticalGoal, type PeerlessInputIntent, type PeerlessProposalIntent, type PeerlessTextIntent } from './inputEvents';
import type { PeerlessViewStore } from './viewStore';

/** @internal */
export const DOCX_ECHO_MAX_UNACKED_INTENTS = 256;

/** @internal Formatting operations that resolve from the causal predecessor selection. */
export type PeerlessFormattingIntent = {
  [K in 'toggleMark' | 'formatRange' | 'clearFormatting' | 'setParagraphAttrs' | 'applyParagraphStyle' | 'formatSelection' | 'toggleStrike' | 'toggleList' | 'changeIndent']:
    { kind: K; payload: Omit<DocxCommandPayloads[K], 'selection'> };
}['toggleMark' | 'formatRange' | 'clearFormatting' | 'setParagraphAttrs' | 'applyParagraphStyle' | 'formatSelection' | 'toggleStrike' | 'toggleList' | 'changeIndent'];

/** @internal Original-intent accounting happens before the shared coalescing hook. */
export interface PeerlessAdmissionSnapshot {
  readOnly: boolean;
  tail: CommandTicket | undefined;
  originalUnackedIntents: number;
  echoEnabled: boolean;
  echoDisabledReason: 'context-cap' | 'intent-cap' | 'fence' | 'unsupported-shaping' | null;
  compositionSlot: number | null;
  heldIntents: number;
}

/** @internal Composition targets are worker snapshots or predecessor tickets. */
export type PeerlessAdmissionTarget =
  | { kind: 'predecessor'; ticket: CommandTicket; suggesting?: YrsAuthor }
  | { kind: 'worker-selection'; snapshot: ReturnType<PeerlessViewStore['getSnapshot']>; suggesting?: YrsAuthor };

/** @internal */
export interface PeerlessAdmissionOptions extends Omit<SubmitOptions<PeerlessEchoIntent>, 'after' | 'echo'> {
  source?: 'typing' | 'paste' | 'formatting' | 'selection' | 'composition';
  onSubmitted?(ticket: CommandTicket): void;
  hostProposal?: boolean;
}

/** @internal Held intents remain admitted and `released` settles with their admission when the composition ends; cap overflow only suppresses their echo. */
export type PeerlessAdmissionResult =
  | { status: 'submitted'; ticket: CommandTicket; echoEnabled: boolean }
  | { status: 'held'; slot: number; echoEnabled: false; released: Promise<PeerlessAdmissionResult> }
  | { status: 'ignored'; reason: 'read-only' | 'empty-composition' };

/** @internal */
export interface PeerlessInputAdmission {
  session: DocxWorkerEditSession<PeerlessEchoIntent>;
  viewStore: PeerlessViewStore;
  echoPainter: PeerlessEchoPainter;
  getSnapshot(): PeerlessAdmissionSnapshot;
  whenCompositionSettled?(): Promise<void>;
  setReadOnly(readOnly: boolean): CommandTicket;
  setSuggesting?(author: string | null): void;
  subscribeReadOnly(listener: () => void): () => void;
  reserveComposition(): { slot: number; target: PeerlessAdmissionTarget };
  settleComposition(
    slot: number,
    commit: { text: string; target: PeerlessAdmissionTarget } | null
  ): PeerlessAdmissionResult;
  dispose(): void;
}

/** @internal */
export interface PeerlessInputAdmissionOptions {
  session: DocxWorkerEditSession<PeerlessEchoIntent>;
  viewStore: PeerlessViewStore;
  echoPainter: PeerlessEchoPainter;
}

export async function flushAdmittedInput(binding: {
  admission?: PeerlessInputAdmission;
  session: Pick<DocxWorkerEditSession<PeerlessEchoIntent>, 'flush'>;
}): Promise<void> {
  await binding.admission?.whenCompositionSettled?.();
  await binding.admission?.getSnapshot?.().tail?.ack;
  await binding.session.flush();
}

/** @internal */
export function createInputAdmission(_options: PeerlessInputAdmissionOptions): PeerlessInputAdmission {
  const { session, viewStore, echoPainter } = _options;
  const pending = new Map<CommandTicket, number>();
  const prefixes = new WeakMap<CommandTicket, Promise<boolean>>();
  const deferred = new WeakMap<CommandTicket, Promise<{ ticket?: CommandTicket; blocked?: boolean }>>();
  const cancellations = new Set<() => void>();
  const held: Array<{ intent: AdmissionIntent; options: PeerlessAdmissionOptions; deletionExtension: boolean;
    release(result: PeerlessAdmissionResult): void; drop(error: Error): void }> = [];
  let tail: CommandTicket | undefined;
  let originalUnackedIntents = 0;
  let echoDisabledReason: PeerlessAdmissionSnapshot['echoDisabledReason'] = null;
  let compositionSlot: number | null = null;
  let nextSlot = 0;
  let nextTicket = 0;
  let disposed = false;
  const compositionWaiters = new Set<{ resolve(): void; reject(error: Error): void }>();
  let verticalGoal: number | undefined;
  let suggesting: string | null = null;
  const stamped = new WeakSet<AdmissionIntent>();
  let requestedReadOnly: boolean | undefined;
  let modeBarrier: CommandTicket | undefined;
  const modeListeners = new Set<() => void>();

  const disableEcho = (reason: NonNullable<PeerlessAdmissionSnapshot['echoDisabledReason']>) => {
    if (echoDisabledReason === reason) return;
    echoDisabledReason = reason;
    echoPainter.disable(reason);
  };
  const count = () => {
    originalUnackedIntents += 1;
    if (originalUnackedIntents > DOCX_ECHO_MAX_UNACKED_INTENTS) disableEcho('intent-cap');
  };
  const resetIdle = () => {
    if (originalUnackedIntents !== 0) return;
    if (tail !== modeBarrier) tail = undefined;
    if (compositionSlot === null) echoDisabledReason = null;
  };
  const releaseCount = (count = 1) => {
    originalUnackedIntents -= count;
    resetIdle();
  };
  const readOnly = () => requestedReadOnly ?? viewStore.getSnapshot().editState?.readOnly ?? session.state.editState?.readOnly ?? false;
  const workerSelection = () => latestSelection(session, viewStore);
  const track = (ticket: CommandTicket) => {
    const prefix = tail ? prefixes.get(tail) ?? Promise.resolve(true) : Promise.resolve(true);
    prefixes.set(ticket, Promise.all([prefix, ticket.ack]).then(([ok, outcome]) =>
      ok && outcome.status !== 'refused' && outcome.status !== 'cancelled', () => false));
    tail = ticket;
    pending.set(ticket, 1);
    const settle = () => {
      if (disposed) return;
      const count = pending.get(ticket) ?? 0;
      pending.delete(ticket);
      releaseCount(count);
    };
    void ticket.ack.then(settle, settle);
  };
  const geometryGoal = (predecessor: CommandTicket | undefined, direction: 'up' | 'down'): Promise<number> => new Promise((resolve, reject) => {
    let unsubscribe = () => {};
    const finish = () => {
      unsubscribe();
      cancellations.delete(cancel);
    };
    const cancel = () => { finish(); reject(new Error('Input admission is disposed')); };
    const check = () => {
      const snapshot = viewStore.getSnapshot();
      const geometry = snapshot.geometry;
      const throughSeq = Math.max(predecessor?.seq ?? 0, snapshot.throughSeq, session.state.throughSeq);
      if (!geometry || geometry.coversThroughSeq < throughSeq) return;
      const rect = direction === 'down' ? geometry.selectionRects.at(-1) : geometry.selectionRects[0];
      const goalX = geometry.caret ? verticalGoal ?? geometry.caret.x
        : rect ? rect.x + (direction === 'down' ? rect.width : 0)
          : verticalGoal ?? geometry.pages[0]?.bounds.x ?? 0;
      finish();
      resolve(goalX);
    };
    cancellations.add(cancel);
    unsubscribe = viewStore.subscribe(check);
    check();
  });
  const defer = (
    after: CommandTicket | undefined,
    waitForSelection: boolean,
    send: (predecessor: CommandTicket | undefined, goalX?: number) => CommandTicket | undefined,
    verticalDirection?: 'up' | 'down',
    barrier = false
  ): CommandTicket => {
    let backing: CommandTicket | undefined;
    let resolveAck!: (outcome: TicketOutcome) => void;
    let rejectAck!: (error: unknown) => void;
    let resolvePresented!: (status: PresentationStatus) => void;
    let resolveReady!: (value: { ticket?: CommandTicket; blocked?: boolean }) => void;
    const ticket: CommandTicket = {
      id: --nextTicket,
      get seq() { return backing?.seq; },
      ack: new Promise((resolve, reject) => { resolveAck = resolve; rejectAck = reject; }),
      presented: new Promise((resolve) => { resolvePresented = resolve; }),
    };
    const ready = new Promise<{ ticket?: CommandTicket; blocked?: boolean }>((resolve) => { resolveReady = resolve; });
    deferred.set(ticket, ready);
    const cancel = (reason = 'disposed') => {
      resolveAck({ status: 'cancelled', reason });
      resolvePresented('dropped');
      resolveReady({ blocked: true });
    };
    cancellations.add(cancel);
    const dependency = after ? deferred.get(after) ?? Promise.resolve({ ticket: after }) : Promise.resolve({});
    void dependency.then(async (value: { ticket?: CommandTicket; blocked?: boolean }) => {
      if (disposed) return cancel();
      if (value.blocked && !barrier) return cancel('dependency-refused');
      if (waitForSelection && value.ticket) {
        const ok = after && prefixes.has(after) ? await prefixes.get(after)! :
          await value.ticket.ack.then((outcome) => outcome.status !== 'refused' && outcome.status !== 'cancelled');
        if (disposed) return cancel();
        if (!ok) return cancel('dependency-refused');
      }
      const goalX = verticalDirection ? await geometryGoal(value.ticket, verticalDirection) : undefined;
      if (disposed) return cancel();
      backing = send(value.ticket, goalX);
      resolveReady({ ticket: backing });
      if (backing) {
        void backing.ack.then(resolveAck, rejectAck);
        void backing.presented.then(resolvePresented, () => resolvePresented('dropped'));
      } else {
        resolveAck({ status: 'noop' });
        resolvePresented('dropped');
      }
    }).catch((error: unknown) => {
      rejectAck(error);
      resolvePresented('dropped');
      resolveReady({ blocked: true });
    }).finally(() => { cancellations.delete(cancel); });
    return ticket;
  };
  const submit = (
    intent: AdmissionIntent,
    options: PeerlessAdmissionOptions = {},
    counted = false,
    target?: PeerlessAdmissionTarget,
    deletionExtension = false
  ): PeerlessAdmissionResult => {
    if (disposed) throw new Error('Input admission is disposed');
    const environment = intent.kind === 'setRenderEnvironment' || intent.kind === 'setFontEnvironment';
    const { source, onSubmitted, hostProposal, ...metadata } = options;
    const granted = isProposal(intent) && hostProposal === true;
    if (readOnly() && !environment && (!isSelection(intent) || deletionExtension) && !granted) {
      if (counted) releaseCount();
      return { status: 'ignored', reason: 'read-only' };
    }
    if (isProposal(intent)) {
      const { hostProposal: _grant, ...payload } = intent.payload;
      intent = { ...intent, payload: granted ? { ...payload, hostProposal: true } : payload } as AdmissionIntent;
    }
    const authored = SUGGESTED[intent.kind] && (intent.kind !== 'navigateCell' || intent.payload.createLastRow === true);
    if (authored && !stamped.has(intent)) {
      const author = target ? target.suggesting : suggesting === null ? undefined : { name: suggesting, date: new Date().toISOString() };
      intent = { ...intent, payload: { ...intent.payload, ...(authored && author ? { suggesting: author } : {}) } } as AdmissionIntent;
      stamped.add(intent);
    }
    try {
      if (!counted) count();
      if (compositionSlot !== null && !target) {
        let release!: (result: PeerlessAdmissionResult) => void;
        let drop!: (error: Error) => void;
        const released = new Promise<PeerlessAdmissionResult>((resolve, reject) => { release = resolve; drop = reject; });
        released.catch(() => {});
        held.push({ intent, options, deletionExtension, release, drop });
        return { status: 'held', slot: compositionSlot, echoEnabled: false, released };
      }
      const fence = 'suggesting' in intent.payload || metadata.fence || isSelection(intent) || !isText(intent) ||
        intent.kind === 'paragraphBreak' || intent.kind === 'replaceSelection' ||
        source === 'paste' || source === 'composition';
      if (fence && echoDisabledReason !== 'intent-cap') disableEcho('fence');
      const after = target?.kind === 'predecessor' ? target.ticket : tail;
      const preview = !echoDisabledReason && isText(intent) ? echoPainter.preview(intent) : undefined;
      const echo = preview ? { ...preview, originalIntentCount: 1 } : undefined;
      const vertical = intent.kind === 'navigate' && (intent.payload.direction === 'up' || intent.payload.direction === 'down')
        ? intent.payload.direction : undefined;
      const send = (predecessor: CommandTicket | undefined, goalX?: number) => {
        const selection = workerSelection();
        if (deletionExtension && !isCollapsed(selection)) return predecessor;
        if (intent.kind === 'navigate' && 'goalX' in intent.payload) {
          verticalGoal = hasImplicitVerticalGoal(intent) ? goalX : intent.payload.goalX;
          goalX = verticalGoal;
        } else verticalGoal = undefined;
        const payload = intent.kind === 'formatRange'
          ? { ...intent.payload, selection: requireSelection(selection) }
          : vertical && intent.kind === 'navigate' && goalX !== undefined
            ? { ...intent.payload, goalX }
            : target?.kind === 'worker-selection' && target.snapshot.workerSelection && isText(intent)
              ? { ...intent.payload, selection: target.snapshot.workerSelection } : intent.payload;
        const ticket = session.submit(intent.kind, payload, {
          ...metadata, after: predecessor, fence: !!fence,
          coalesceKey: fence ? undefined : metadata.coalesceKey ?? intent.kind,
          echo,
        });
        if (onSubmitted) {
          try { onSubmitted(ticket); } catch (error) { console.error('[PeerlessInput] submission observer failed', error); }
        }
        return ticket;
      };
      const waitForSelection = intent.kind === 'formatRange' || deletionExtension || vertical !== undefined;
      const ticket = deletionExtension || vertical || (after && (waitForSelection || deferred.has(after))) ||
        (intent.kind === 'formatRange' && !workerSelection())
        ? defer(after, waitForSelection, send, vertical) : send(after)!;
      track(ticket);
      return { status: 'submitted', ticket, echoEnabled: echo !== undefined };
    } catch (error) {
      if (!counted) releaseCount();
      throw error;
    }
  };
  const admission: PeerlessInputAdmission = {
    session, viewStore, echoPainter,
    whenCompositionSettled() {
      if (disposed) return Promise.reject(new Error('Input admission is disposed'));
      if (compositionSlot === null) return Promise.resolve();
      return new Promise<void>((resolve, reject) => { compositionWaiters.add({ resolve, reject }); });
    },
    getSnapshot: () => ({
      readOnly: readOnly(),
      tail, originalUnackedIntents, echoEnabled: echoDisabledReason === null,
      echoDisabledReason, compositionSlot, heldIntents: held.length,
    }),
    setSuggesting(author) { suggesting = author; },
    setReadOnly(value) {
      if (disposed) throw new Error('Input admission is disposed');
      const send = () => session.submit('setReadOnly', { readOnly: value }, { fence: true });
      const ticket = tail && deferred.has(tail) ? defer(tail, false, send, undefined, true) : send();
      requestedReadOnly = value;
      modeBarrier = tail = ticket;
      prefixes.set(ticket, ticket.ack.then((outcome) =>
        outcome.status !== 'refused' && outcome.status !== 'cancelled', () => false));
      disableEcho('fence');
      modeListeners.forEach((listener) => listener());
      const settle = (failed: boolean) => {
        if (disposed || modeBarrier !== ticket) return;
        modeBarrier = undefined;
        if (failed) requestedReadOnly = undefined;
        if (tail === ticket && originalUnackedIntents === 0) tail = undefined;
        modeListeners.forEach((listener) => listener());
      };
      void ticket.ack.then((outcome) => settle(outcome.status === 'refused' || outcome.status === 'cancelled'), () => settle(true));
      return ticket;
    },
    subscribeReadOnly(listener) {
      modeListeners.add(listener);
      return () => { modeListeners.delete(listener); };
    },
    reserveComposition() {
      if (disposed) throw new Error('Input admission is disposed');
      if (compositionSlot !== null) throw new Error('Composition is already active');
      compositionSlot = ++nextSlot;
      const target: PeerlessAdmissionTarget = tail
        ? { kind: 'predecessor', ticket: tail, suggesting: suggesting === null ? undefined : { name: suggesting, date: new Date().toISOString() } }
        : { kind: 'worker-selection', snapshot: viewStore.getSnapshot(), suggesting: suggesting === null ? undefined : { name: suggesting, date: new Date().toISOString() } };
      const fence = () => { void session.flush().catch(() => {}); return undefined; };
      if (tail && deferred.has(tail)) defer(tail, false, fence, undefined, true);
      else fence();
      if (echoDisabledReason !== 'intent-cap') disableEcho('fence');
      return { slot: compositionSlot, target };
    },
    settleComposition(slot, commit) {
      if (disposed) throw new Error('Input admission is disposed');
      if (slot !== compositionSlot) throw new Error('Composition slot is not active');
      compositionSlot = null;
      let failed = false;
      let failure: unknown;
      let result: PeerlessAdmissionResult = { status: 'ignored', reason: 'empty-composition' };
      try {
        if (commit?.text) result = submit({ kind: 'replaceSelection', payload: { text: commit.text } },
          { source: 'composition', fence: true }, false, commit.target);
      } catch (error) { failed = true; failure = error; }
      const compositionFailed = failed;
      for (const entry of held.splice(0)) {
        if (!failed) {
          try { entry.release(submit(entry.intent, entry.options, true, undefined, entry.deletionExtension)); continue; } catch (error) { failed = true; failure = error; }
        }
        releaseCount();
        entry.drop(failure instanceof Error ? failure : new Error(String(failure)));
      }
      for (const waiter of compositionWaiters) waiter.resolve();
      compositionWaiters.clear();
      resetIdle();
      if (compositionFailed) throw failure;
      return result;
    },
    dispose() {
      disposed = true;
      for (const waiter of compositionWaiters) waiter.reject(new Error('Input admission is disposed'));
      compositionWaiters.clear();
      for (const cancel of cancellations) cancel();
      cancellations.clear();
      for (const entry of held.splice(0)) entry.drop(new Error('Input admission is disposed'));
      pending.clear();
      modeListeners.clear();
      tail = undefined;
      compositionSlot = null;
      originalUnackedIntents = 0;
      admissions.delete(admission);
    },
  };
  admissions.set(admission, (intent, options, deletionExtension) => submit(intent, options, false, undefined, deletionExtension));
  return admission;
}

/** @internal */
export function admitIntent(
  _admission: PeerlessInputAdmission,
  _intent: PeerlessInputIntent,
  _options?: PeerlessAdmissionOptions
): PeerlessAdmissionResult {
  return submitAdmission(_admission, _intent, _options);
}

/** @internal */
export function admitInputIntents(
  admission: PeerlessInputAdmission,
  intents: readonly PeerlessInputIntent[],
  options: PeerlessAdmissionOptions
): void {
  for (const [index, intent] of intents.entries()) {
    const next = intents[index + 1];
    const deletionExtension = intent.kind === 'navigate' && intent.payload.extend &&
      (next?.kind === 'deleteBackward' || next?.kind === 'deleteForward');
    const snapshot = admission.getSnapshot();
    const waiting = snapshot.tail !== undefined || snapshot.compositionSlot !== null;
    const readOnly = admission.getSnapshot().readOnly;
    if (deletionExtension && readOnly) continue;
    if (deletionExtension && !waiting && !isCollapsed(latestSelection(admission.session, admission.viewStore))) continue;
    submitAdmission(admission, intent, {
      ...options, source: intent.kind === 'navigate' ? 'selection' : options.source,
    }, deletionExtension && waiting);
  }
}

/** @internal */
export function submitFormattingIntent(
  _admission: PeerlessInputAdmission,
  _intent: PeerlessFormattingIntent,
  _options?: PeerlessAdmissionOptions
): PeerlessAdmissionResult {
  return submitAdmission(_admission, _intent, { ..._options, source: 'formatting', fence: true });
}

type AdmissionIntent = PeerlessInputIntent | PeerlessFormattingIntent;

/** Kinds the worker records as suggestions; suggesting mode stamps its author on exactly these. */
const SUGGESTED = {
  insertText: true, replaceSelection: true, deleteBackward: true, deleteForward: true, paragraphBreak: true,
  pasteFragment: true,
  applyEdits: true, replaceMatches: true, insertImage: true, insertTable: true, insertPageBreak: true, insertSectionBreak: true,
  toggleMark: true, formatRange: true, clearFormatting: true, setParagraphAttrs: true, applyParagraphStyle: true,
  formatSelection: true, toggleStrike: true, toggleList: true, changeIndent: true, navigateCell: true,
  insertTableRow: true, insertTableColumn: true, deleteTableRow: true, deleteTableColumn: true,
  deleteTable: true, mergeTableCells: true, splitTableCell: true, setTableCellShading: true, setTableBorders: true, setTableProperties: true,
  selectTableCells: false,
  addComment: false, resizeObject: false, moveObject: false,
  setHyperlink: false, setObjectProperties: false, setSectionProperties: false, setNoteProperties: false, setWatermark: false,
  setSectionMargins: false, setTabStops: false, removeTabStop: false,
  setContentControlValue: false, addRepeatingItem: false, removeRepeatingItem: false,
  openPart: false, ensurePart: false, removePart: false,
  select: false, navigate: false, selectPoint: false, extendSelection: false,
  undo: false, redo: false, setRenderEnvironment: false, setFontEnvironment: false,
  propose: false, setProposalStates: false, withdrawProposals: false,
} satisfies Record<AdmissionIntent['kind'], boolean>;
const admissions = new WeakMap<PeerlessInputAdmission,
  (intent: AdmissionIntent, options?: PeerlessAdmissionOptions, deletionExtension?: boolean) => PeerlessAdmissionResult>();

function submitAdmission(
  admission: PeerlessInputAdmission, intent: AdmissionIntent, options?: PeerlessAdmissionOptions,
  deletionExtension = false
): PeerlessAdmissionResult {
  const submit = admissions.get(admission);
  if (!submit) throw new Error('Input admission is unavailable');
  return submit(intent, options, deletionExtension);
}

function latestSelection(session: DocxWorkerEditSession<PeerlessEchoIntent>, viewStore: PeerlessViewStore) {
  const snapshot = viewStore.getSnapshot();
  return session.state.throughSeq >= snapshot.throughSeq && session.state.selection !== undefined
    ? session.state.selection : snapshot.workerSelection;
}

function isCollapsed(selection: DocxCommandSelection | null | undefined): boolean {
  return !!selection && selection.anchor.story === selection.focus.story &&
    selection.anchor.paraId === selection.focus.paraId && selection.anchor.offset === selection.focus.offset;
}

function requireSelection(selection: DocxCommandSelection | null | undefined): DocxCommandSelection {
  if (!selection) throw new Error('Worker selection is unavailable');
  return selection;
}

function isProposal(intent: AdmissionIntent): intent is PeerlessProposalIntent {
  return intent.kind === 'propose' || intent.kind === 'setProposalStates' || intent.kind === 'withdrawProposals';
}

function isSelection(intent: AdmissionIntent): boolean {
  return intent.kind === 'openPart' || intent.kind === 'select' || intent.kind === 'navigate' || intent.kind === 'navigateCell' ||
    intent.kind === 'selectTableCells' || intent.kind === 'selectPoint' || intent.kind === 'extendSelection';
}

function isText(intent: AdmissionIntent): intent is PeerlessTextIntent {
  return intent.kind === 'insertText' || intent.kind === 'replaceSelection' ||
    intent.kind === 'deleteBackward' || intent.kind === 'deleteForward' || intent.kind === 'paragraphBreak';
}
