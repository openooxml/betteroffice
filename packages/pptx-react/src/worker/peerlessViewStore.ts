import {
  decodeCommandDisplayList, findLine, frameAssetIds, hitGeometry, indexGeometry, navigateGeometry,
  presentationSharedInputs, PptxWorkerEditorDisposedError, selectionGeometry, transformedRect,
} from '@betteroffice/pptx/worker';
import type { SlideDisplayList, PositionedTextRun } from '@betteroffice/pptx';
import type {
  PresentationCommandAck, PresentationCommandBasis, PresentationCommandClient, PresentationCommandEditState,
  PresentationCommandFrame, PresentationCommandFailure, PresentationCommandIntent, PresentationCommandMotion,
  PresentationCommandOutcome, PresentationCommandSelected, PresentationCommandSelection,
  PresentationCommandSelectionReference, CommandGeometryLine, CommandRect,
} from '@betteroffice/pptx/worker';
import { measureEchoPosition } from './localEchoPainter';
import type { CommandTicket } from '../../../../shared/office-session/commands/types';

export const PEERLESS_ACTIVE_CONTEXT_BYTES = 16 * 1024;
export const PEERLESS_PENDING_INTENTS = 256;

export interface PeerlessRect extends CommandRect {}
export interface PeerlessConfirmedFrame { frame: PresentationCommandFrame; displayList: SlideDisplayList }

export interface PeerlessPendingIntent {
  seq: number;
  basisThroughSeq: number;
  afterSeq?: number;
  intent: PresentationCommandIntent;
  status: 'pending' | 'applied';
  echoed: boolean;
}

export interface PeerlessComposition {
  id: string;
  selection: PresentationCommandSelection | null;
  text: string;
}

export interface PeerlessEcho {
  text: string;
  sequences: readonly number[];
  bounds: PeerlessRect;
  caretPosition: number;
  insertionOffset: number;
  insertionStyle: PositionedTextRun;
}

export interface PeerlessActiveLine extends CommandGeometryLine { text: string }

export interface PeerlessViewSnapshot {
  confirmed: PeerlessConfirmedFrame | null;
  selection: PresentationCommandSelected | null;
  focused: boolean;
  caret: PeerlessRect | null;
  selectionRects: readonly PeerlessRect[];
  composition: PeerlessComposition | null;
  inputFailure?: PptxCommandRefusedError;
  echo: PeerlessEcho | null;
  activeLine: PeerlessActiveLine | null;
  activeContextBytes: number;
  pending: readonly PeerlessPendingIntent[];
  waitingInputCount: number;
  appliedThroughSeq: number;
  presentedThroughSeq: number;
  editState: PresentationCommandEditState;
  thumbnails: ReadonlyMap<string, PeerlessConfirmedFrame>;
  dirtyThumbnailIds: ReadonlySet<string>;
}

export type PeerlessPaintCheckpoint = (() => void) & { commit?(): void };

export interface PeerlessViewPainters {
  checkpoint?(pixels?: boolean): PeerlessPaintCheckpoint;
  paintFrame(frame: PeerlessConfirmedFrame, view: PeerlessViewSnapshot,
    media?: PresentationCommandFrame['media']): void;
  paintEcho(view: PeerlessViewSnapshot): void;
  restoreLine(view: PeerlessViewSnapshot): void;
}

export interface PeerlessViewStore {
  getSnapshot(): PeerlessViewSnapshot;
  subscribePresentation(listener: () => void): () => void;
  subscribeStructure(listener: () => void): () => void;
  subscribeEditState(listener: () => void): () => void;
  dispatch(intent: PresentationCommandIntent, options?: { reportRefusal?: boolean }): Promise<PresentationCommandAck>;
  fence(): Promise<void>;
  setSelection(selection: PresentationCommandSelection | null): void;
  setFocused(focused: boolean): void;
  beginComposition(id: string, inputTime: number): void;
  updateComposition(text: string): void;
  installMedia(media: PresentationCommandFrame['media']): void;
  installThumbnail(frame: PresentationCommandFrame): boolean;
  dispose(): void;
}

interface WaitingInput {
  seq: number;
  ticket?: CommandTicket;
  after?: WaitingInput;
  basis: PresentationCommandBasis;
  basisThroughSeq: number;
  intent: PresentationCommandIntent;
  chunks: string[];
  echoed: boolean;
  quiet: boolean;
  operation: Promise<PresentationCommandAck>;
  resolve(ack: PresentationCommandAck): void;
  reject(error: unknown): void;
}

interface ViewController {
  install(frame: PresentationCommandFrame): boolean;
  hit(point: { x: number; y: number }, borderScale?: number): ReturnType<typeof hitGeometry>;
  navigate(motion: PresentationCommandMotion, extend: boolean): PresentationCommandSelection | null;
  inTableCell(): boolean;
  known(): boolean;
  reference(predecessorSelection?: boolean): PresentationCommandSelectionReference | null;
  finishComposition(): (PeerlessComposition & { at: PresentationCommandSelectionReference }) | null;
}

const gestures = new WeakMap<PeerlessViewStore, Set<string>>();

export function setPeerlessGestureActive(store: PeerlessViewStore, kind: string, active: boolean): void {
  let current = gestures.get(store);
  if (!current) { current = new Set(); gestures.set(store, current); }
  if (active) current.add(kind); else current.delete(kind);
}

export function peerlessGestureActive(store: PeerlessViewStore): boolean { return Boolean(gestures.get(store)?.size); }

const controllers = new WeakMap<PeerlessViewStore, ViewController>();

export class PptxFrameDecodeError extends Error {
  override readonly name = 'PptxFrameDecodeError';
  constructor(readonly seq: number, readonly cause: unknown) {
    super('Could not decode the covering presentation frame');
  }
}

export class PptxCommandRefusedError extends Error {
  override readonly name = 'PptxCommandRefusedError';
  constructor(readonly operation: PresentationCommandIntent['kind'], readonly outcome: PresentationCommandOutcome) {
    super(`Presentation ${operation} refused: ${outcome.reason}`);
  }
}

export function createPeerlessViewStore(
  client: PresentationCommandClient, painters: PeerlessViewPainters,
  { echoEnabled = false, onError }: { echoEnabled?: boolean; onError?(error: PptxFrameDecodeError | PptxCommandRefusedError): void } = {},
): PeerlessViewStore {
  const initial = client.state;
  const identity = { session: initial.session, incarnation: initial.incarnation, epoch: initial.epoch };
  const media = new Map<string, ArrayBuffer>((initial.frame?.media ?? []).map((asset) => [asset.assetId, asset.bytes]));
  const presentation = new Set<() => void>();
  const structure = new Set<() => void>();
  const editState = new Set<() => void>();
  const waiting: WaitingInput[] = [];
  const submitted = new Map<WaitingInput, PresentationCommandIntent>();
  const replies = new Map<number, PresentationCommandSelected | null>();
  const navigationTargets = new Map<number, string>();
  let navigationReply: { seq: number; slideId: string } | undefined;
  const captured = new Map<Promise<PresentationCommandAck>, (error: unknown) => void>();
  const quietSeqs = new Set<number>();
  const sharedInput = presentationSharedInputs.get(client);
  const resultReferences = new WeakMap<PresentationCommandSelectionReference, WaitingInput>();
  let compatibilityNextSeq = initial.throughSeq + 1;
  let lastSentSeq = initial.throughSeq;
  let selectionTailSeq = initial.throughSeq;
  let chainTail = initial.throughSeq;
  let chainTailInput: WaitingInput | undefined;
  let unresolvedThroughSeq = initial.throughSeq;
  let localSelectionSeq = initial.throughSeq;
  let localOnlySelection = false;
  let disposed = false;
  let failure: unknown;
  let draining = false;
  let compositionVisible = false;
  let pendingSelection: { inTableCell: boolean } | undefined;
  let geometry: ReturnType<typeof indexGeometry> = initial.frame
    ? indexGeometry(decodeCommandDisplayList(initial.frame)) : { hits: [], lines: new Map() };
  let echoCharacters: Map<number, Set<string>> | null = null;
  let echoContext = initial.frame?.coversThroughSeq === initial.throughSeq
    ? { selection: initial.selection, run: initial.editState.toolbar.echoRun } : null;
  const echoContexts = new Map<number, NonNullable<typeof echoContext>>();
  let view: PeerlessViewSnapshot = {
    confirmed: initial.frame ? decodeFrame(initial.frame) : null, selection: initial.selection, focused: false,
    caret: null, selectionRects: [], composition: null, echo: null, activeLine: null,
    activeContextBytes: 0, pending: [], waitingInputCount: 0,
    appliedThroughSeq: initial.throughSeq, presentedThroughSeq: initial.frame?.coversThroughSeq ?? initial.throughSeq,
    editState: initial.editState, thumbnails: new Map(), dirtyThumbnailIds: new Set(),
  };

  function withMedia(frame: PeerlessConfirmedFrame): PeerlessConfirmedFrame {
    const additions: PresentationCommandFrame['media'] = [];
    for (const assetId of frameAssetIds(frame.displayList)) {
      const bytes = media.get(assetId);
      if (bytes && !frame.frame.media.some((asset) => asset.assetId === assetId)) additions.push({ assetId, bytes });
    }
    return additions.length ? { ...frame, frame: { ...frame.frame, media: [...frame.frame.media, ...additions] } } : frame;
  }

  function decodeFrame(frame: PresentationCommandFrame): PeerlessConfirmedFrame {
    return withMedia({ frame, displayList: decodeCommandDisplayList(frame) });
  }

  function known(): boolean {
    return waiting.length === 0 && submitted.size === 0 && unresolvedThroughSeq <= view.presentedThroughSeq;
  }

  function reference(predecessorSelection = false): PresentationCommandSelectionReference | null {
    if (!known() || predecessorSelection && !localOnlySelection && localSelectionSeq > view.presentedThroughSeq) {
      const at: PresentationCommandSelectionReference = { source: 'afterSeq',
        seq: sharedInput ? chainTailInput?.ticket?.seq ?? selectionTailSeq : chainTail };
      if (sharedInput && chainTailInput) resultReferences.set(at, chainTailInput);
      return at;
    }
    const selected = view.selection;
    if (selected?.kind === 'object') {
      if (!selected.storyId || !view.editState.toolbar.canEditText) return null;
      return { source: 'confirmed', selection: { slideId: selected.slideId, shapeId: selected.shapeId,
        storyId: selected.storyId, kind: 'shape', anchor: 0, focus: 0, affinity: 'downstream' } };
    }
    return selected ? { source: 'confirmed', selection: selected } : null;
  }

  function notify(listeners: Set<() => void>): void {
    for (const listener of [...listeners]) if (!disposed && listeners.has(listener)) listener();
  }

  function echoStyle(): PositionedTextRun | null {
    if (view.editState.readOnly || !echoContext?.selection || !echoContext.run || !view.selection ||
      view.selection.kind === 'object' || !sameSelection(echoContext.selection, view.selection)) return null;
    const entry = findLine(geometry, view.selection);
    const position = view.selection.focus;
    const neighbour = echoContext.run === 'after'
      ? entry?.line.runs.find((run) => run.start <= position && position < run.end) ?? null
      : [...(entry?.line.runs ?? [])].reverse().find((run) => run.start < position && position <= run.end) ?? null;
    if (!neighbour || !echoScript(neighbour.text) || !/[a-z]/.test(neighbour.text) ||
      (neighbour.letterSpacingPx ?? 0) !== 0 || (neighbour.baselineOffsetPx ?? 0) !== 0 ||
      entry?.line.runs.some((run) => run.start <= position && position <= run.end &&
        (run.fontId !== neighbour.fontId || !echoScript(run.text)))) return null;
    return neighbour;
  }

  function canEchoText(text: string): boolean {
    if (!echoEnabled) return false;
    const neighbour = echoStyle();
    if (!neighbour || !echoScript(text)) return false;
    if (!echoCharacters) {
      echoCharacters = new Map();
      for (const lines of geometry.lines.values()) for (const { line } of lines) for (const run of line.runs) {
        let characters = echoCharacters.get(run.fontId);
        if (!characters) { characters = new Set(); echoCharacters.set(run.fontId, characters); }
        for (const character of run.text) characters.add(character);
      }
    }
    const characters = echoCharacters.get(neighbour.fontId);
    return [...text].every((character) => characters?.has(character));
  }

  function rebuild(): void {
    const entry = findLine(geometry, view.selection);
    const original = entry?.line.runs.map((run) => run.text).join('') ?? '';
    const activeLine = entry && original.length * 2 <= PEERLESS_ACTIVE_CONTEXT_BYTES
      ? { ...entry, text: original } : null;
    let insertion = '';
    const sequences: number[] = [];
    for (const record of view.pending) {
      if (isLocalSelection(record.intent) || record.intent.kind === 'setReadOnly') continue;
      if (!record.echoed || record.intent.kind !== 'insertText') break;
      insertion += record.intent.payload.text;
      sequences.push(record.seq);
    }
    const insertionStyle = echoStyle();
    const composition = compositionVisible && waiting.length === 0 &&
      view.pending.every((record) => record.echoed && record.intent.kind === 'insertText') &&
      canEchoText(view.composition?.text ?? '')
      ? view.composition : null;
    if (composition) insertion += composition.text;
    const position = view.selection?.focus ?? 0;
    const offset = position - (entry?.line.start ?? 0);
    const text = original.slice(0, offset) + insertion + original.slice(offset);
    const bounds = entry ? transformedRect(entry.transforms, {
      x: entry.box.x, y: entry.line.y, width: entry.box.w, height: entry.line.height,
    }) : null;
    const echo = activeLine && bounds && insertionStyle && canEchoText(insertion) &&
      view.selection?.anchor === position && offset >= 0 &&
      offset <= original.length && text.length * 2 <= PEERLESS_ACTIVE_CONTEXT_BYTES &&
      (sequences.length > 0 || composition !== null)
      ? { text, sequences, bounds, caretPosition: position + insertion.length, insertionOffset: offset, insertionStyle } : null;
    const projection = selectionGeometry(geometry, view.selection);
    if (echo && activeLine) {
      const x = measureEchoPosition(activeLine, echo, echo.caretPosition);
      if (x !== null) projection.caret = transformedRect(activeLine.transforms, {
        x, y: activeLine.line.y, width: 1, height: activeLine.line.height,
      });
    }
    view = { ...view, activeLine, echo, activeContextBytes: activeLine ? Math.max(original.length, echo?.text.length ?? 0) * 2 : 0,
      ...projection };
  }

  function paintLocal(): void {
    rebuild();
    if (view.confirmed) { painters.restoreLine(view); painters.paintEcho(view); }
    notify(presentation);
  }

  function matches(ack: PresentationCommandAck): boolean {
    return ack.session === identity.session && ack.incarnation === identity.incarnation && ack.epoch === identity.epoch;
  }

  function receiveAck(ack: PresentationCommandAck): void {
    if (disposed || failure || !matches(ack) || ack.throughSeq < view.appliedThroughSeq) return;
    if (ack.throughSeq === view.appliedThroughSeq) {
      if (ack.frame) install(ack.frame);
      return;
    }
    const changedState = JSON.stringify(view.editState) !== JSON.stringify(ack.editState);
    if (view.pending.some((record) => preservesSelection(record.intent) &&
      ack.outcomes.some((outcome) => outcome.seq === record.seq && outcome.committed))) {
      for (const input of waiting) {
        const at = 'at' in input.intent.payload ? input.intent.payload.at : undefined;
        const dependency = input.after?.ticket?.seq ?? (at?.source === 'afterSeq' ? at.seq : undefined);
        if (input.after && input.after.ticket?.seq === undefined || dependency !== undefined && dependency > ack.throughSeq) continue;
        if (at?.source === 'afterSeq') input.intent = chainedIntent(input.intent, ack.throughSeq);
        input.after = undefined;
        input.basis = { epoch: ack.epoch, version: ack.version };
        input.basisThroughSeq = ack.throughSeq;
      }
    }
    const remoteHistory = view.pending.some((record) => record.seq === ack.historyThroughSeq && (record.intent.kind === 'applyRemoteUpdate' || record.intent.kind === 'collabPipe'));
    const refusals = view.pending.flatMap((record) => {
      const outcome = ack.outcomes.find((entry) => entry.seq === record.seq);
      return outcome?.reason && !quietSeqs.has(record.seq) && (outcome.status === 'refused' || outcome.status === 'cancelled')
        ? [new PptxCommandRefusedError(record.intent.kind, outcome)] : [];
    });
    for (const outcome of ack.outcomes) quietSeqs.delete(outcome.seq);
    if (ack.historyThroughSeq === ack.throughSeq) {
      localOnlySelection = false;
      localSelectionSeq = ack.throughSeq;
      const composition = remoteHistory && view.composition && ack.selection && ack.selection.kind !== 'object' &&
        ack.selection.storyId === view.composition.selection?.storyId
        ? { ...view.composition, selection: ack.selection } : null;
      if (!composition) compositionVisible = false;
      view = { ...view, selection: ack.selection, composition };
      if (ack.activeSlideId !== null && ack.activeSlideId !== undefined &&
        (!remoteHistory || ack.activeSlideId !== view.confirmed?.frame.slideId)) {
        navigationReply = { seq: ack.throughSeq, slideId: ack.activeSlideId };
      }
    }
    for (const [seq, slideId] of navigationTargets) {
      const outcome = ack.outcomes.find((entry) => entry.seq === seq);
      if (seq <= ack.throughSeq && ack.activeSlideId === slideId &&
        outcome?.status !== 'refused' && outcome?.status !== 'cancelled' &&
        (!navigationReply || ack.throughSeq > navigationReply.seq)) navigationReply = { seq: ack.throughSeq, slideId };
    }
    replies.set(ack.throughSeq, ack.selection);
    echoContexts.set(ack.throughSeq, { selection: ack.selection, run: ack.editState.toolbar.echoRun });
    const rejected = new Set(ack.outcomes.filter((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled')
      .map((outcome) => outcome.seq));
    for (const record of view.pending) {
      if (record.afterSeq !== undefined && rejected.has(record.afterSeq)) rejected.add(record.seq);
    }
    const rejectedInputs = new Set<WaitingInput>();
    for (const input of waiting) {
      const at = 'at' in input.intent.payload ? input.intent.payload.at : undefined;
      if (sharedInput ? input.after && (rejectedInputs.has(input.after) ||
        input.after.ticket?.seq !== undefined && rejected.has(input.after.ticket.seq))
        : at?.source === 'afterSeq' && rejected.has(at.seq)) {
        rejectedInputs.add(input);
        if (!sharedInput) rejected.add(input.seq);
      }
    }
    if (view.composition && [...view.pending.filter((record) => rejected.has(record.seq)), ...rejectedInputs].some((record) =>
      record.intent.kind === 'compositionStart' && record.intent.payload.compositionId === view.composition?.id)) {
      compositionVisible = false;
      view = { ...view, composition: null };
    }
    for (let index = waiting.length - 1; index >= 0; index -= 1) {
      const input = waiting[index];
      if (!rejectedInputs.has(input)) continue;
      waiting.splice(index, 1);
      const outcome: PresentationCommandOutcome = { seq: sharedInput ? lastSentSeq + index + 1 : input.seq,
        status: 'cancelled', reason: 'dependency-refused', committed: false };
      if (!input.quiet) refusals.push(new PptxCommandRefusedError(input.intent.kind, outcome));
      input.resolve({ ...ack, outcomes: [outcome] });
    }
    if (!sharedInput) {
      const sequences = new Map(waiting.map((input, index) => [input.seq, lastSentSeq + index + 1]));
      for (const input of waiting) {
        input.seq = sequences.get(input.seq)!;
        const at = 'at' in input.intent.payload ? input.intent.payload.at : undefined;
        if (at?.source === 'afterSeq' && sequences.has(at.seq)) input.intent = chainedIntent(input.intent, sequences.get(at.seq)!);
      }
      compatibilityNextSeq = lastSentSeq + waiting.length + 1;
      chainTail = Math.max(selectionTailSeq, [...waiting].reverse().find((input) => !preservesSelection(input.intent))?.seq ?? 0);
    } else if (chainTailInput && rejectedInputs.has(chainTailInput)) {
      chainTailInput = [...waiting].reverse().find((input) => !preservesSelection(input.intent));
    }
    view = { ...view, appliedThroughSeq: ack.throughSeq, editState: ack.editState,
      pending: view.pending.filter((record) => !rejected.has(record.seq)).map((record) => ({ ...record,
        status: record.seq <= ack.throughSeq ? 'applied' : record.status,
        echoed: record.echoed && !rejected.has(record.seq) })), waitingInputCount: waiting.length };
    const dirtyThumbnailIds = new Set(view.dirtyThumbnailIds);
    const thumbnails = new Map(view.thumbnails);
    for (const id of ack.dirtySlideIds) {
      dirtyThumbnailIds.add(id);
      thumbnails.delete(id);
    }
    if (ack.slides) {
      const ids = new Set(ack.slides.map((slide) => slide.id));
      for (const id of thumbnails.keys()) if (!ids.has(id)) thumbnails.delete(id);
      for (const id of dirtyThumbnailIds) if (!ids.has(id)) dirtyThumbnailIds.delete(id);
      if (ack.activeSlideId === null) {
        geometry = { hits: [], lines: new Map() };
        view = { ...view, confirmed: null, selection: null, caret: null, selectionRects: [], echo: null,
          activeLine: null, activeContextBytes: 0 };
      }
    }
    if (!view.confirmed) {
      view = { ...view, selection: ack.selection, presentedThroughSeq: ack.throughSeq,
        pending: view.pending.filter((record) => record.seq > ack.throughSeq) };
      replies.clear(); echoContexts.clear();
    }
    view = { ...view, thumbnails, dirtyThumbnailIds };
    if (refusals.length) view = { ...view, inputFailure: refusals[refusals.length - 1] };
    if (ack.slides || ack.dirtySlideIds.some((id) => id !== view.confirmed?.frame.slideId)) notify(structure);
    if (rejected.size) paintLocal();
    if (changedState) notify(editState);
    if (ack.frame) install(ack.frame);
    drain();
    for (const error of refusals) { try { onError?.(error); } catch {} }
  }

  function install(frame: PresentationCommandFrame): boolean {
    if (disposed || failure || frame.incarnation !== identity.incarnation || frame.documentEpoch !== identity.epoch ||
      frame.coversThroughSeq < view.presentedThroughSeq) return false;
    const state = client.state;
    if (matches(state) && state.throughSeq === frame.coversThroughSeq) {
      const { slides, size, ...ack } = state;
      receiveAck({ ...ack, ...(state.historyThroughSeq === state.throughSeq ? { slides, size } : {}), frame: undefined });
    }
    const previous = view.confirmed?.frame;
    const activation = navigationReply?.seq === frame.coversThroughSeq && navigationReply.slideId === frame.slideId;
    const changedSlide = frame.slideId !== previous?.slideId;
    if (changedSlide && !activation || previous && !changedSlide &&
      (frame.slideRevision < previous.slideRevision || frame.slideRevision === previous.slideRevision &&
        frame.coversThroughSeq <= view.presentedThroughSeq) ||
      frame.coversThroughSeq < view.presentedThroughSeq) return false;
    const covered = view.pending.filter((record) => record.seq <= frame.coversThroughSeq);
    if (!activation && previous && !changedSlide && frame.version === previous.version &&
      frame.slideRevision === previous.slideRevision && covered.length > 0 &&
      covered.every((record) => record.intent.kind === 'collabPipe')) {
      view.confirmed!.frame = { ...frame, displayList: previous.displayList, media: previous.media };
      echoContext = echoContexts.get(frame.coversThroughSeq) ?? echoContext;
      view = { ...view, presentedThroughSeq: frame.coversThroughSeq,
        pending: view.pending.filter((record) => record.seq > frame.coversThroughSeq) };
      retireCoverage(frame.coversThroughSeq);
      drain();
      return true;
    }
    const selection = activation ? replies.get(frame.coversThroughSeq) ?? null
      : localOnlySelection || localSelectionSeq > frame.coversThroughSeq ? view.selection
      : replies.has(frame.coversThroughSeq) ? replies.get(frame.coversThroughSeq)! : view.selection;
    let confirmed: PeerlessConfirmedFrame;
    try { confirmed = decodeFrame(frame); }
    catch (error) {
      if (!sharedInput) throw error;
      throw new PptxFrameDecodeError(frame.coversThroughSeq, error);
    }
    const saved = sharedInput ? { view, geometry, echoContext, echoCharacters, compositionVisible } : undefined;
    let rollback: PeerlessPaintCheckpoint | undefined;
    try {
      if (saved) rollback = painters.checkpoint?.(echoEnabled);
      echoContext = echoContexts.get(frame.coversThroughSeq) ??
        (sharedInput && frame.coversThroughSeq === view.presentedThroughSeq ? echoContext : null);
      geometry = indexGeometry(confirmed.displayList);
      echoCharacters = null;
      if (view.composition && view.pending.some((record) =>
        record.seq <= frame.coversThroughSeq && record.intent.kind === 'compositionStart')) {
        compositionVisible = selection !== null && selection.kind !== 'object' && selection.anchor === selection.focus;
        if (selection && selection.kind !== 'object') view = { ...view, composition: { ...view.composition, selection } };
      }
      let thumbnails = view.thumbnails;
      if (changedSlide && previous && !view.dirtyThumbnailIds.has(previous.slideId)) {
        thumbnails = new Map(thumbnails).set(previous.slideId, view.confirmed!);
      }
      view = { ...view, confirmed, selection, thumbnails, composition: activation ? null : view.composition,
        presentedThroughSeq: frame.coversThroughSeq,
        pending: view.pending.filter((record) => record.seq > frame.coversThroughSeq) };
      rebuild();
      painters.paintFrame(confirmed, view);
      painters.paintEcho(view);
    } catch (error) {
      if (!saved) throw error;
      ({ view, geometry, echoContext, echoCharacters, compositionVisible } = saved);
      rollback?.();
      return false;
    }
    rollback?.commit?.();
    if (activation) { localOnlySelection = false; compositionVisible = false; }
    retireCoverage(frame.coversThroughSeq);
    notify(presentation);
    if (changedSlide) notify(structure);
    drain();
    return true;
  }

  function retireCoverage(throughSeq: number): void {
    for (const seq of navigationTargets.keys()) if (seq <= throughSeq) navigationTargets.delete(seq);
    for (const seq of replies.keys()) if (seq < throughSeq) replies.delete(seq);
    for (const seq of echoContexts.keys()) if (seq <= throughSeq) echoContexts.delete(seq);
  }

  function blocked(): boolean {
    return view.pending.length >= PEERLESS_PENDING_INTENTS || view.pending.some((record) =>
      record.status === 'pending' && !record.echoed && isTextMutation(record.intent));
  }

  function fail(error: unknown): void {
    if (disposed || failure) return;
    failure = error;
    waiting.length = 0;
    submitted.clear();
    for (const reject of captured.values()) reject(error);
    view = { ...view, waitingInputCount: 0, editState: { ...view.editState, readOnly: true,
      toolbar: { ...view.editState.toolbar, canEditText: false } } };
    notify(editState);
  }

  function recordSent(input: WaitingInput, intent: PresentationCommandIntent): void {
    if (sharedInput) {
      const predecessor = input.after?.ticket?.seq;
      if (predecessor !== undefined) intent = chainedIntent(intent, predecessor);
      input.after = undefined;
      if (isLocalSelection(intent)) localSelectionSeq = input.seq;
    }
    const at = 'at' in intent.payload ? intent.payload.at : undefined;
    if (intent.kind === 'activateSlide') navigationTargets.set(input.seq, intent.payload.slideId);
    if (intent.kind === 'setSelection' && intent.payload.selection && intent.payload.selection.slideId !== view.confirmed?.frame.slideId) {
      navigationTargets.set(input.seq, intent.payload.selection.slideId);
    }
    if (input.quiet) quietSeqs.add(input.seq);
    if (needsWorkerResolution(intent)) {
      unresolvedThroughSeq = Math.max(unresolvedThroughSeq, input.seq);
      if (intent.kind !== 'collabPipe') localOnlySelection = false;
    }
    view = { ...view, pending: [...view.pending, { seq: input.seq, basisThroughSeq: input.basisThroughSeq,
      afterSeq: at?.source === 'afterSeq' ? at.seq : undefined, intent, status: 'pending', echoed: input.echoed }] };
    lastSentSeq = input.seq;
    if (!preservesSelection(intent)) selectionTailSeq = input.seq;
    if (intent.kind !== 'collabPipe') paintLocal();
  }

  function send(input: WaitingInput): void {
    let intent = input.intent;
    if (intent.kind === 'undo' || intent.kind === 'redo' || intent.kind === 'activateSlide' ||
      intent.kind === 'insertTextBox' || intent.kind === 'insertShape' || intent.kind === 'insertPicture' || intent.kind === 'applyEdits' || intent.kind === 'registerFont' || intent.kind === 'invoke') {
      input.basis = { epoch: identity.epoch, version: client.state.version };
      input.basisThroughSeq = client.state.throughSeq;
    }
    if (intent.kind === 'insertText' && input.chunks.length) {
      intent = { ...intent, payload: { ...intent.payload, text: intent.payload.text + input.chunks.join('') } };
    }
    let operation: Promise<PresentationCommandAck> | undefined;
    if (sharedInput) {
      const predecessor = input.after?.ticket;
      if (predecessor?.seq !== undefined) intent = chainedIntent(intent, predecessor.seq);
      let submission: ReturnType<typeof sharedInput.submit>;
      try { submission = sharedInput.submit(intent, input.basis, predecessor); }
      catch (error) { fail(error); input.reject(error); return; }
      input.ticket = submission.ticket;
      operation = submission.operation;
      if (input.ticket.seq === undefined) {
        submitted.set(input, intent);
        void operation.then((ack) => { submitted.delete(input); settleAck(input, ack); },
          (error: unknown) => { fail(error); input.reject(error); });
        return;
      }
      input.seq = input.ticket.seq;
    }
    recordSent(input, intent);
    try {
      (operation ?? client.dispatch(intent, input.basis)).then((ack) => {
        if (sharedInput) settleAck(input, ack);
        else { receiveAck(ack); input.resolve(ack); }
      }, (error: unknown) => {
        fail(error); input.reject(error);
      });
    } catch (error) { fail(error); input.reject(error); }
  }

  function settleAck(input: WaitingInput, ack: PresentationCommandAck): void {
    try { receiveAck(ack); input.resolve(ack); }
    catch (error) {
      input.reject(error);
      if (error instanceof PptxFrameDecodeError) {
        try { onError?.(error); } catch {}
        drain();
      } else fail(error);
    }
  }

  function drain(): void {
    if (draining || disposed || failure) return;
    draining = true;
    try {
      while (waiting.length && !blocked()) {
        const input = waiting.shift()!;
        view = { ...view, waitingInputCount: waiting.length };
        send(input);
      }
    } finally { draining = false; }
  }

  const store: PeerlessViewStore = {
    getSnapshot: () => view,
    subscribePresentation: (listener) => subscribe(presentation, listener),
    subscribeStructure: (listener) => subscribe(structure, listener),
    subscribeEditState: (listener) => subscribe(editState, listener),
    dispatch(intent, { reportRefusal = true } = {}) {
      if (disposed) return Promise.reject(new PptxWorkerEditorDisposedError());
      if (failure) return Promise.reject(failure);
      if (view.inputFailure) view = { ...view, inputFailure: undefined };
      const wasKnown = known();
      const predecessor = sharedInput ? chainTailInput?.ticket?.seq ?? selectionTailSeq : chainTail;
      const originalReference = 'at' in intent.payload ? intent.payload.at : undefined;
      const after = sharedInput && !wasKnown &&
        ('at' in intent.payload || intent.kind === 'setSelection' || intent.kind === 'compositionEnd')
        ? chainTailInput : originalReference && resultReferences.get(originalReference);
      if (!wasKnown && ('at' in intent.payload || intent.kind === 'setSelection' || intent.kind === 'compositionEnd')) {
        intent = chainedIntent(intent, predecessor);
      }
      if (wasKnown && intent.kind === 'setSelection') view = { ...view, selection: intent.payload.selection };
      const tail = waiting[waiting.length - 1];
      if (tail?.intent.kind === 'insertText' && !tail.echoed && intent.kind === 'insertText' &&
        intent.payload.at.source === 'afterSeq' && (sharedInput ? after === tail : intent.payload.at.seq === tail.seq) &&
        intent.payload.undoGroup === tail.intent.payload.undoGroup && intent.payload.style === tail.intent.payload.style &&
        intent.payload.inputTime === tail.intent.payload.inputTime) {
        tail.chunks.push(intent.payload.text);
        tail.echoed = false;
        return tail.operation;
      }
      const seq = sharedInput ? 0 : compatibilityNextSeq++;
      if (!sharedInput && !preservesSelection(intent)) chainTail = seq;
      if (wasKnown && intent.kind === 'setSelection') {
        localSelectionSeq = seq;
        localOnlySelection = false;
      }
      const echoed = waiting.length === 0 && !blocked() && intent.kind === 'insertText' &&
        intent.payload.style === undefined && canEchoText(intent.payload.text) &&
        view.selection !== null &&
        view.selection.anchor === view.selection.focus && !!view.activeLine &&
        (intent.payload.at.source === 'afterSeq'
          ? intent.payload.at.seq === predecessor && [...view.pending, ...waiting].some((record) =>
            record.seq === predecessor && record.echoed && record.intent.kind === 'insertText')
          : wasKnown && sameSelection(intent.payload.at.selection, view.selection)) &&
        view.pending.every((record) => record.echoed && record.intent.kind === 'insertText') &&
        (view.echo?.text.length ?? view.activeLine.text.length) * 2 + intent.payload.text.length * 2 <= PEERLESS_ACTIVE_CONTEXT_BYTES;
      let resolve!: (ack: PresentationCommandAck) => void;
      let reject!: (error: unknown) => void;
      const operation = new Promise<PresentationCommandAck>((done, fail) => { resolve = done; reject = fail; });
      captured.set(operation, reject);
      void operation.then(() => { captured.delete(operation); }, () => { captured.delete(operation); });
      const input: WaitingInput = { seq, intent, after, basisThroughSeq: view.presentedThroughSeq,
        basis: { epoch: identity.epoch, version: view.confirmed?.frame.version ?? client.state.version },
        echoed, quiet: !reportRefusal, chunks: [], operation, resolve, reject };
      let target: PresentationCommandSelected | null | undefined;
      if (intent.kind === 'selectAtPoint' && !intent.payload.drag) {
        const hit = hitGeometry(geometry, intent.payload);
        target = hit?.kind === 'text' ? { slideId: intent.payload.frame.slideId,
          shapeId: hit.shapeId, storyId: hit.storyId, anchor: hit.position, focus: hit.position,
          affinity: hit.affinity ?? 'downstream' } : null;
      } else if (intent.kind === 'setSelection') target = intent.payload.selection;
      else if (intent.kind === 'selectObject' || intent.kind === 'activateSlide') target = null;
      if (target !== undefined) {
        const pending = { inTableCell: Boolean(navigateGeometry(geometry, target, 'nextCell', false)) };
        pendingSelection = pending;
        const settled = () => { if (pendingSelection === pending) pendingSelection = undefined; };
        void operation.then(settled, settled);
      }
      if (!preservesSelection(intent)) chainTailInput = input;
      if (waiting.length || blocked()) {
        waiting.push(input);
        view = { ...view, waitingInputCount: waiting.length };
        if (intent.kind !== 'collabPipe') notify(presentation);
      } else send(input);
      return operation;
    },
    async fence() {
      if (disposed) throw new PptxWorkerEditorDisposedError();
      if (failure) throw failure;
      await Promise.all([...captured.keys()]);
      if (disposed) throw new PptxWorkerEditorDisposedError();
      if (failure) throw failure;
    },
    setSelection(selection) {
      if (disposed || failure) return;
      if (!known()) {
        void store.dispatch({ kind: 'setSelection', payload: { selection, inputTime: performance.now(), undoGroup: 'selection' } }).catch(() => {});
        return;
      }
      localOnlySelection = true;
      view = { ...view, selection };
      paintLocal();
    },
    setFocused(focused) {
      if (disposed || focused === view.focused) return;
      view = { ...view, focused };
      paintLocal();
    },
    beginComposition(id, inputTime) {
      const at = reference(true);
      const selected = view.selection;
      if (disposed || failure || view.composition || !at || selected?.kind === 'object' && at.source !== 'afterSeq') return;
      const selection = selected?.kind === 'object' ? null : selected;
      compositionVisible = known() && selection !== null && selection.anchor === selection.focus;
      view = { ...view, composition: { id, selection, text: '' } };
      void store.dispatch({ kind: 'compositionStart', payload: { compositionId: id,
        at, inputTime, undoGroup: id } }).catch(() => {});
    },
    updateComposition(text) {
      if (disposed || failure || !view.composition) return;
      view = { ...view, composition: { ...view.composition, text } };
      paintLocal();
    },
    installMedia(assets) {
      if (disposed || failure || !assets.length) return;
      for (const asset of assets) media.set(asset.assetId, asset.bytes);
      const confirmed = view.confirmed && withMedia(view.confirmed);
      const thumbnails = new Map(view.thumbnails);
      let changed = false;
      for (const [id, frame] of thumbnails) {
        const hydrated = withMedia(frame);
        if (hydrated !== frame) { thumbnails.set(id, hydrated); changed = true; }
      }
      if (confirmed === view.confirmed && !changed) return;
      const activeChanged = confirmed !== view.confirmed;
      view = { ...view, confirmed, thumbnails: changed ? thumbnails : view.thumbnails };
      if (activeChanged && confirmed) {
        painters.paintFrame(confirmed, view);
        painters.paintEcho(view);
        notify(presentation);
      }
      if (changed) notify(structure);
    },
    installThumbnail(frame) {
      if (disposed || frame.incarnation !== identity.incarnation || frame.documentEpoch !== identity.epoch) return false;
      const cached = view.thumbnails.get(frame.slideId);
      if (cached && cached.frame.slideRevision >= frame.slideRevision) return false;
      const thumbnails = new Map(view.thumbnails);
      thumbnails.set(frame.slideId, decodeFrame(frame));
      const dirtyThumbnailIds = new Set(view.dirtyThumbnailIds);
      dirtyThumbnailIds.delete(frame.slideId);
      view = { ...view, thumbnails, dirtyThumbnailIds };
      notify(structure);
      return true;
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      unsubscribeAck(); unsubscribeFrame(); unsubscribeMedia(); unsubscribeFailure(); unsubscribeSent?.();
      waiting.length = 0;
      submitted.clear();
      for (const reject of captured.values()) reject(new PptxWorkerEditorDisposedError());
      presentation.clear(); structure.clear(); editState.clear(); replies.clear(); echoContexts.clear(); captured.clear();
      navigationTargets.clear();
      media.clear();
      controllers.delete(store);
    },
  };

  function subscribe(listeners: Set<() => void>, listener: () => void): () => void {
    if (!disposed) listeners.add(listener);
    return () => { listeners.delete(listener); };
  }

  controllers.set(store, {
    install, known, reference, hit: (point, borderScale) => hitGeometry(geometry, point, undefined, borderScale),
    navigate: (motion, extend) => navigateGeometry(geometry, view.selection, motion, extend),
    inTableCell: () => pendingSelection?.inTableCell ?? Boolean(navigateGeometry(geometry, view.selection, 'nextCell', false)),
    finishComposition() {
      const composition = view.composition;
      const at = reference();
      if (!composition) return null;
      view = { ...view, composition: null };
      compositionVisible = false;
      paintLocal();
      return at ? { ...composition, at } : null;
    },
  });
  const unsubscribeAck = client.subscribe(receiveAck);
  const unsubscribeFrame = client.subscribeFrame(install);
  const unsubscribeMedia = client.subscribeMedia((assets) => store.installMedia(assets));
  const unsubscribeFailure = client.subscribeFailure((error: PresentationCommandFailure) => fail(error));
  const unsubscribeSent = sharedInput?.subscribeSent(() => {
    for (const [input, intent] of submitted) {
      const seq = input.ticket?.seq;
      if (seq === undefined) continue;
      submitted.delete(input);
      input.seq = seq;
      recordSent(input, intent);
    }
  });
  rebuild();
  if (view.confirmed) painters.paintFrame(view.confirmed, view);
  return store;
}

function chainedIntent(intent: PresentationCommandIntent, seq: number): PresentationCommandIntent {
  const at = { source: 'afterSeq' as const, seq };
  switch (intent.kind) {
    case 'moveShape': return 'delta' in intent.payload ? { ...intent, payload: { ...intent.payload, at } } : intent;
    case 'insertText': return { ...intent, payload: { ...intent.payload, at } };
    case 'replaceText': return { ...intent, payload: { ...intent.payload, at } };
    case 'formatText': return { ...intent, payload: { ...intent.payload, at } };
    case 'setParagraphAlignment': return { ...intent, payload: { ...intent.payload, at } };
    case 'deleteBackward': return { ...intent, payload: { ...intent.payload, at } };
    case 'deleteForward': return { ...intent, payload: { ...intent.payload, at } };
    case 'paragraphBreak': return { ...intent, payload: { ...intent.payload, at } };
    case 'compositionStart': return { ...intent, payload: { ...intent.payload, at } };
    case 'compositionCommit': return { ...intent, payload: { ...intent.payload, at } };
    case 'compositionEnd': return { ...intent, payload: { ...intent.payload, at } };
    case 'setSelection': return { ...intent, payload: { ...intent.payload, at } };
    case 'moveSelection': return { ...intent, payload: { ...intent.payload, at } };
    case 'selectAtPoint': return { ...intent, payload: { ...intent.payload, at } };
    default: return intent;
  }
}

function isTextMutation(intent: PresentationCommandIntent): boolean {
  if (intent.kind === 'collabPipe') return intent.payload.ops.some((op) => op.op === 'event' && op.event.type === 'message');
  return ['insertText', 'replaceText', 'deleteBackward', 'deleteForward', 'paragraphBreak', 'compositionCommit',
    'formatText', 'setParagraphAlignment', 'undo', 'redo', 'moveShape', 'setShapeRect', 'insertTextBox', 'insertShape', 'insertPicture', 'applyEdits', 'registerFont', 'applyRemoteUpdate', 'invoke'].includes(intent.kind);
}

function isLocalSelection(intent: PresentationCommandIntent): boolean {
  return intent.kind === 'setSelection' && intent.payload.at?.source !== 'afterSeq';
}

function preservesSelection(intent: PresentationCommandIntent): boolean {
  if (intent.kind === 'collabPipe') return !isTextMutation(intent);
  return intent.kind === 'invoke' && ['setSlideNotes', 'addComment', 'replyToComment', 'setCommentStatus',
    'setCommentPosition', 'removeComment', 'setCommentFlavor'].includes(intent.payload.method);
}

function needsWorkerResolution(intent: PresentationCommandIntent): boolean {
  return intent.kind === 'activateSlide' || intent.kind === 'selectObject' || isTextMutation(intent) && !preservesSelection(intent) || intent.kind === 'moveSelection' || intent.kind === 'selectAtPoint' ||
    intent.kind === 'compositionStart' || intent.kind === 'setSelection' && intent.payload.at?.source === 'afterSeq';
}

function sameSelection(a: PresentationCommandSelected, b: PresentationCommandSelected): boolean {
  return a.slideId === b.slideId && a.shapeId === b.shapeId && a.storyId === b.storyId &&
    a.anchor === b.anchor && a.focus === b.focus && a.affinity === b.affinity && a.kind === b.kind;
}

function echoScript(text: string): boolean { return !/[^\u0020-\u007e\u00a0-\u00ff]/.test(text); }

export function installCoveringFrame(store: PeerlessViewStore, frame: PresentationCommandFrame): boolean {
  return controllers.get(store)?.install(frame) ?? false;
}

export function hitConfirmedGeometry(store: PeerlessViewStore, point: { x: number; y: number }, borderScale?: number): ReturnType<typeof hitGeometry> {
  return controllers.get(store)?.hit(point, borderScale) ?? null;
}

export function peerlessKnownState(store: PeerlessViewStore): boolean {
  return controllers.get(store)?.known() ?? false;
}

export function peerlessInputReference(store: PeerlessViewStore, predecessorSelection = false): PresentationCommandSelectionReference | null {
  return controllers.get(store)?.reference(predecessorSelection) ?? null;
}

export function navigateConfirmedSelection(store: PeerlessViewStore, key: string, extend = false,
  modifiers?: Pick<KeyboardEvent, 'altKey' | 'ctrlKey' | 'metaKey' | 'shiftKey'>): PresentationCommandSelection | null {
  const motion = commandMotion(key, modifiers);
  if (!peerlessKnownState(store) || !motion) return null;
  return controllers.get(store)?.navigate(motion, extend) ?? null;
}

/** Whether the current or pending selection is in a table cell. */
export function peerlessSelectionInTableCell(store: PeerlessViewStore): boolean {
  return controllers.get(store)?.inTableCell() ?? false;
}

export function commandMotion(key: string,
  modifiers?: Pick<KeyboardEvent, 'altKey' | 'ctrlKey' | 'metaKey' | 'shiftKey'>): PresentationCommandMotion | null {
  const mac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform);
  const word = mac ? modifiers?.altKey : modifiers?.ctrlKey;
  switch (key) {
    case 'ArrowLeft': return modifiers?.metaKey ? 'lineStart' : word ? 'wordLeft' : 'left';
    case 'ArrowRight': return modifiers?.metaKey ? 'lineEnd' : word ? 'wordRight' : 'right';
    case 'Home': return modifiers?.metaKey || modifiers?.ctrlKey ? 'storyStart' : 'lineStart';
    case 'End': return modifiers?.metaKey || modifiers?.ctrlKey ? 'storyEnd' : 'lineEnd';
    case 'ArrowUp': return 'up'; case 'ArrowDown': return 'down';
    case 'Tab': return modifiers?.altKey || modifiers?.ctrlKey || modifiers?.metaKey ? null
      : modifiers?.shiftKey ? 'previousCell' : 'nextCell';
    default: return null;
  }
}

export function commitComposition(store: PeerlessViewStore, text: string, inputTime: number): Promise<PresentationCommandAck | null> {
  const composition = controllers.get(store)?.finishComposition();
  if (!composition) return Promise.resolve(null);
  const commit = text ? store.dispatch({ kind: 'compositionCommit', payload: { compositionId: composition.id,
    at: composition.at, text, inputTime, undoGroup: composition.id } }) : null;
  const end = store.dispatch({ kind: 'compositionEnd', payload: { compositionId: composition.id,
    inputTime, undoGroup: composition.id, ...(text ? {} : { cancelled: true }) } });
  if (!commit) return end;
  void end.catch(() => {});
  return commit;
}
