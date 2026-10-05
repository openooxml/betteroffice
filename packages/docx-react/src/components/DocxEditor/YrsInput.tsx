/** Text and IME input surface backed by sticky session positions. */

import React, {
  forwardRef,
  memo,
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import type { CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import {
  sameYrsSelection,
  type YrsAuthor,
  type YrsInputPositionMap,
  type YrsInlineFormatDelta,
  type YrsLoc,
  type YrsResidentCaretSnapshot,
  type YrsSelection,
  type YrsSession,
  type YrsStoryRange,
} from '@betteroffice/docx/yrs';
import {
  effectiveZoom,
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { ResidentFrameApplyResult } from './hooks/useDisplayList';
import type { ResolveDisplayListQueries } from './hooks/displayListQueryEpochGate';
import { findWordBoundaries } from '@betteroffice/docx/utils';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import {
  yrsCellLocFromStory,
  yrsCellStory,
  yrsSelectionNearTable,
  yrsSelectionPlainText,
  yrsTableSelectionRange,
} from './yrsCommands';
import { InputOperationQueue } from './inputOperationQueue';
import { awaitWorkerOpenReplica, requestWorkerOpenReplicaReadiness } from './internals/workerOpenReplica';
import { requestQueuedOpeningInput } from './internals/queuedOpeningInput';
import { scrollIntoViewDelta, scrollViewport } from './internals/viewportBand';
import { DocxCommandAdmissionError } from '../../commands/createDocxCommandStore';
import { paragraphVerticalMove, VerticalCaretGoal } from './verticalCaretGoal';
import {
  shouldScrollCaretIntoView,
  type LayoutUpdateOrigin,
} from './internals/viewportAnchoring';

const READER_SCROLL = ['wheel', 'touchmove'] as const;

export interface YrsDisplaySelection {
  anchor: number;
  head: number;
}

export interface YrsInputRef {
  focus(): void;
  blur(): void;
  isFocused(): boolean;
  flushPendingInput(): Promise<void>;
  /**
   * Runs `operation` in input order, after input accepted before the call and
   * after an active IME composition commits. Rejects with
   * {@link DocxCommandAdmissionError} when that input failed, the document was
   * replaced, or the input unmounted first.
   */
  runAfterPendingInput<T>(operation: () => T | Promise<T>): Promise<T>;
  hasPendingInput(): boolean;
  hasHeldInput?(): boolean;
  queueSelection?(prepare: () => Promise<() => void>, force?: boolean, inTable?: () => boolean): boolean;
  captureSelectionFromDisplay?(
    anchor: number, head: number, story: string, kind: 'caret' | 'range' | 'word' | 'paragraph'
  ): () => void;
  setSelectionFromDisplay(anchor: number, head?: number, story?: string, gesture?: number): void;
  selectWordAtDisplay(position: number, story?: string): void;
  selectParagraphAtDisplay(position: number, story?: string): void;
  displaySelection(): YrsDisplaySelection | null;
  /** The current selection stays where it is on screen rather than scrolling into view. */
  keepSelectionInPlace(): void;
  applyStoredFormatting(action: YrsStoredFormattingAction): void;
  clearStoredFormatting(): void;
  storedFormatting(): YrsStoredFormatting | null;
  insertText(text: string): void;
  deleteSelection(): void;
  selectAll(): void;
  /** The selection's plain text, for an input whose document lives in the worker. */
  readSelectedText?(): Promise<string> | null;
  beginGesture?(): number;
  currentGesture?(): number;
  isGestureCurrent?(gesture: number): boolean;
}

export type YrsStoredFormattingAction =
  | {
      type: 'toggle';
      mark: 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';
      active: boolean;
    }
  | { type: 'set'; delta: YrsInlineFormatDelta }
  | { type: 'clear' };

export interface YrsStoredFormatting {
  clear: boolean;
  delta: YrsInlineFormatDelta;
}

export interface YrsInputProps {
  enabled: boolean;
  readOnly: boolean;
  holdInput?: boolean;
  inputScope?: number;
  replicaReadyRef?: React.RefObject<boolean>;
  /** Advances on input that supersedes input still waiting for the replica. */
  inputEpoch?: () => number;
  /** Applies a selection gesture recorded while the replica loaded; waiting input follows it. */
  applyPendingSelection?: () => void;
  /** Whether a missing selection starts as a caret at the story start; input always starts one. */
  seedSelection?: boolean;
  session: YrsSession | null;
  story?: string;
  isSuggesting?: boolean;
  author?: string;
  inputPositionMap(story?: string): YrsInputPositionMap | null;
  displayPositionToLoc(position: number, story?: string): YrsLoc | null;
  resolveDisplayTarget?(position: number): { story: string; displayPosition: number } | null;
  locToDisplayPosition(loc: YrsLoc): number | null;
  nextParagraphStyleId?(styleId: string | null): string | null;
  displayListQueries?: DisplayListQueries | null;
  resolveDisplayListQueries?: ResolveDisplayListQueries;
  displayListFrameEpoch?: number | null;
  residentCaret?: YrsResidentCaretSnapshot | null;
  residentCaretAuthoritative?: boolean;
  layoutUpdateOrigin?: LayoutUpdateOrigin;
  canvasHostRef?: React.RefObject<HTMLDivElement | null>;
  /** Called for selection-only changes and direct document mutations. */
  onStateChange(
    selection: YrsDisplaySelection,
    docChanged: boolean,
    residentLayoutReady?: boolean,
    residentCaretReady?: boolean,
    updateOrigin?: LayoutUpdateOrigin,
    inWorker?: boolean
  ): void;
  onDirectInput(stories?: string | readonly string[]): void;
  /** One-owner body text path; false until the resident frame is initialized. */
  applyResidentInput?(text: string): Promise<ResidentFrameApplyResult | null>;
  /** One-owner collapsed delete/merge path; false until the resident frame is initialized. */
  applyResidentDelete?(
    direction: 'backward' | 'forward',
    count?: number
  ): Promise<ResidentFrameApplyResult | null>;
  onFocusChange?(focused: boolean): void;
  /** Accepted input started or finished waiting to be applied. */
  onPendingInputChange?(pending: boolean): void;
  /** Document-mutating input landed (keeps the worker-painted caret mode alive). */
  onCaretInput?(): void;
  /** Text input dispatched — called synchronously from the input event, before
   * the async apply, so the DOM caret hides before the new frame presents. */
  onCaretInputDispatched?(): void;
  /** Selection-only move or IME start (immediate swap to the DOM blink caret). */
  onCaretInterrupt?(): void;
}

const BASE_STYLE: CSSProperties = {
  position: 'fixed',
  width: '1px',
  minWidth: '1px',
  padding: 0,
  margin: 0,
  border: 0,
  outline: 0,
  opacity: 0,
  overflow: 'hidden',
  resize: 'none',
  zIndex: -1,
  background: 'transparent',
  color: 'transparent',
  caretColor: 'transparent',
};

interface HeldSelection {
  kind: 'selection';
  prepare: () => Promise<() => void>;
  apply?: () => void;
  inTable?: () => boolean;
  inputTime?: number;
}

type NavigationDirection = 'left' | 'right' | 'up' | 'down' | 'home' | 'end';

const UNDO_CAPTURE_TIMEOUT_MS = 500;

type HeldInput = (
  | HeldSelection
  | { kind: 'navigation'; direction: NavigationDirection; extend: boolean; wholeDocument: boolean; byWord: boolean }
  | { kind: 'text' | 'composition'; text: string }
  | { kind: 'tab'; shift: boolean }
  | { kind: 'split' }
  | { kind: 'delete'; direction: 'backward' | 'forward' }
  | { kind: 'delete-selection'; cut?: CutCopy }
  | { kind: 'select-all' }
  | { kind: 'copy'; apply: (session: YrsSession) => void | Promise<void>; onDropped: () => void }
  | { kind: 'undo-boundary' }
) & { inputTime?: number };

interface CutCopy {
  written: boolean;
}

function isOpeningHeldInput(entry: HeldInput): boolean {
  return entry.kind === 'text' || entry.kind === 'composition' || entry.kind === 'split' ||
    entry.kind === 'delete' || entry.kind === 'delete-selection';
}

interface HeldReplayBatch {
  operations: Promise<void>[];
  mutated: boolean;
  selectionChanged: boolean;
}

function previousCodePointOffset(text: string, offset: number): number {
  if (offset <= 0) return 0;
  const last = text.charCodeAt(offset - 1);
  if (last >= 0xdc00 && last <= 0xdfff && offset > 1) {
    const first = text.charCodeAt(offset - 2);
    if (first >= 0xd800 && first <= 0xdbff) return offset - 2;
  }
  return offset - 1;
}

function nextCodePointOffset(text: string, offset: number): number {
  if (offset >= text.length) return text.length;
  const first = text.charCodeAt(offset);
  if (first >= 0xd800 && first <= 0xdbff && offset + 1 < text.length) {
    const last = text.charCodeAt(offset + 1);
    if (last >= 0xdc00 && last <= 0xdfff) return offset + 2;
  }
  return offset + 1;
}

function previousWordOffset(text: string, offset: number): number {
  let next = Math.max(0, Math.min(offset, text.length));
  while (next > 0 && /\s/u.test(text[next - 1])) next -= 1;
  while (next > 0 && !/\s/u.test(text[next - 1])) next -= 1;
  return next;
}

function nextWordOffset(text: string, offset: number): number {
  let next = Math.max(0, Math.min(offset, text.length));
  while (next < text.length && !/\s/u.test(text[next])) next += 1;
  while (next < text.length && /\s/u.test(text[next])) next += 1;
  return next;
}

function toRange(selection: YrsSelection, map: YrsInputPositionMap): YrsStoryRange {
  const index = (loc: YrsLoc): number => {
    const para = map.paragraphs.find((entry) => entry.paraId === loc.paraId);
    return para ? para.displayStart + 1 + loc.offset : 0;
  };
  const [start, end] =
    index(selection.anchor) <= index(selection.head)
      ? [selection.anchor, selection.head]
      : [selection.head, selection.anchor];
  return {
    story: start.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
}

/** The body or a table cell or control story laid out with it: resident input can apply there. */
function isBodyFlowStory(story: string): boolean {
  return story === 'body' || story.startsWith('body:');
}

const YrsInputComponent = forwardRef<YrsInputRef, YrsInputProps>(function YrsInput(
  {
    enabled,
    readOnly,
    holdInput = false,
    inputScope,
    replicaReadyRef,
    inputEpoch,
    applyPendingSelection,
    seedSelection = true,
    session,
    story = 'body',
    isSuggesting = false,
    author = 'User',
    inputPositionMap,
    displayPositionToLoc,
    resolveDisplayTarget,
    locToDisplayPosition,
    nextParagraphStyleId,
    displayListQueries,
    resolveDisplayListQueries,
    displayListFrameEpoch = null,
    residentCaret = null,
    residentCaretAuthoritative = false,
    layoutUpdateOrigin = 'local',
    canvasHostRef,
    onStateChange,
    onDirectInput,
    applyResidentInput,
    applyResidentDelete,
    onFocusChange,
    onPendingInputChange,
    onCaretInput,
    onCaretInputDispatched,
    onCaretInterrupt,
  },
  ref
) {
  const replicaReady = replicaReadyRef?.current !== false;
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const composingRef = useRef(false);
  const compositionPendingRef = useRef(false);
  const compositionCommitRef = useRef('');
  const compositionHeldRef = useRef(false);
  const discardedCompositionRef = useRef(false);
  const compositionWaitersRef = useRef(new Set<() => void>());
  const heldInputWaitersRef = useRef(new Set<() => void>());
  const holdInputRef = useRef(holdInput);
  holdInputRef.current = holdInput;
  const heldInputRef = useRef({ scope: inputScope, entries: [] as HeldInput[] });
  const pendingSelectionsRef = useRef<HeldSelection[]>([]);
  if (heldInputRef.current.scope !== inputScope) {
    if (compositionHeldRef.current && (composingRef.current || compositionPendingRef.current)) {
      discardedCompositionRef.current = true;
    }
    for (const entry of heldInputRef.current.entries) if (entry.kind === 'copy') entry.onDropped();
    heldInputRef.current = { scope: inputScope, entries: [] };
    pendingSelectionsRef.current = [];
  }
  const inputLifetimeRef = useRef({ session, enabled, mounted: true });
  inputLifetimeRef.current.session = session;
  inputLifetimeRef.current.enabled = enabled;
  const isCurrentInput = useCallback(
    (started: YrsSession | null, queue = inputOperationQueueRef.current): boolean =>
      inputLifetimeRef.current.mounted && inputLifetimeRef.current.enabled &&
      inputLifetimeRef.current.session === started && inputOperationQueueRef.current === queue,
    []
  );
  const storedFormattingByParagraphRef = useRef(new Map<string, YrsStoredFormatting>());
  const onPendingInputChangeRef = useRef(onPendingInputChange);
  onPendingInputChangeRef.current = onPendingInputChange;
  const requestOpeningPeer = useCallback((): void => {
    if (enabled && session && (holdInput || replicaReadyRef?.current === false)) {
      requestQueuedOpeningInput(session);
    }
  }, [enabled, holdInput, replicaReadyRef, session]);
  const holdOperation = useCallback(
    (entry: HeldInput): boolean => {
      if (!holdInput) return false;
      if (!readOnly) {
        entry.inputTime ??= performance.now();
        heldInputRef.current.entries.push(entry);
        if (isOpeningHeldInput(entry)) requestOpeningPeer();
        onPendingInputChangeRef.current?.(true);
      }
      return true;
    },
    [holdInput, readOnly, requestOpeningPeer]
  );
  const inputOperationQueueRef = useRef<InputOperationQueue | null>(null);
  const queuedSessionRef = useRef(session);
  if (!inputOperationQueueRef.current || queuedSessionRef.current !== session) {
    queuedSessionRef.current = session;
    const queue = new InputOperationQueue(
      (error) => {
        console.error('[YrsInput] queued input operation failed', error);
      },
      (pending) => {
        if (!replicaReadyRef || isCurrentInput(session, queue)) {
          onPendingInputChangeRef.current?.(pending);
        }
      }
    );
    inputOperationQueueRef.current = queue;
  }
  const pendingResidentTextRef = useRef<{ text: string } | null>(null);
  const pendingResidentDeleteRef = useRef<{
    direction: 'backward' | 'forward';
    count: number;
  } | null>(null);
  const pendingResidentFrameEpochRef = useRef<number | null>(null);
  const heldReplayBatchRef = useRef<HeldReplayBatch | null>(null);
  const pendingCaretTableRef = useRef<(() => boolean) | undefined>(undefined);
  const pendingLocalCaretRevealRef = useRef<{
    scroller: HTMLElement;
    scrollTop: number;
  } | null>(null);
  const verticalCaretGoalRef = useRef(new VerticalCaretGoal());
  const displayListQueriesRef = useRef(displayListQueries);
  const displayListFrameEpochRef = useRef(displayListFrameEpoch);
  const resolveDisplayListQueriesRef = useRef(resolveDisplayListQueries);
  // Sticky (not display-position) selection the caret-into-view step last acted
  // on. `undefined` is the pre-mount sentinel; display positions shift under a
  // remote insert above the caret while sticky locs do not, so comparing locs is
  // what keeps a viewer's viewport still during someone else's edit.
  const lastCaretScrollSelectionRef = useRef<YrsSelection | null | undefined>(undefined);
  // A selection the input shows without scrolling to it: a default caret, a replayed gesture.
  const quietSelectionRef = useRef<YrsSelection | null>(null);
  displayListQueriesRef.current = displayListQueries;
  displayListFrameEpochRef.current = displayListFrameEpoch;
  resolveDisplayListQueriesRef.current = resolveDisplayListQueries;
  const [positionStyle, setPositionStyle] = useState<CSSProperties>({ left: 0, top: 0, height: 1 });
  const [selectionEpoch, setSelectionEpoch] = useState(0);

  // Every queued operation seals the pending text and delete batches: input
  // after it must not join a batch that runs before it.
  const sealInputBatches = useCallback((): void => {
    pendingResidentTextRef.current = null;
    pendingResidentDeleteRef.current = null;
  }, []);

  const replicaInputRef = useRef({ inputEpoch, applyPendingSelection });
  replicaInputRef.current = { inputEpoch, applyPendingSelection };
  const preparePendingSelections = useCallback(
    async (admitted: YrsSession | null, queue: InputOperationQueue | null): Promise<void> => {
      // Bind source positions before replayed edits move them.
      while (pendingSelectionsRef.current.length > 0 && isCurrentInput(admitted, queue)) {
        const entries = pendingSelectionsRef.current;
        pendingSelectionsRef.current = [];
        for (const entry of entries) {
          entry.apply = await entry.prepare();
          if (!isCurrentInput(admitted, queue)) return;
        }
      }
    },
    [isCurrentInput]
  );
  const readerScrollRef = useRef<{ waiting: number; scrolls: number; stop(): void } | null>(null);
  // Counts the reader's scrolling from now until the returned check runs.
  const watchReaderScroll = useCallback((): (() => boolean) => {
    let watch = readerScrollRef.current;
    if (!watch) {
      const host = canvasHostRef?.current;
      const surface = host ? findVerticalScrollParentOrRoot(host) : null;
      const created = { waiting: 0, scrolls: 0, stop: () => {} };
      const onScroll = (): void => {
        created.scrolls += 1;
      };
      for (const type of READER_SCROLL) {
        surface?.addEventListener(type, onScroll, { capture: true, passive: true });
      }
      created.stop = () => {
        for (const type of READER_SCROLL) surface?.removeEventListener(type, onScroll, true);
      };
      readerScrollRef.current = watch = created;
    }
    const current = watch;
    const from = current.scrolls;
    current.waiting += 1;
    return () => {
      current.waiting -= 1;
      if (current.waiting === 0) {
        current.stop();
        if (readerScrollRef.current === current) readerScrollRef.current = null;
      }
      return current.scrolls !== from;
    };
  }, [canvasHostRef]);

  const enqueueInputOperation = useCallback(
    (
      operation: (waited: boolean) => void | Promise<void>,
      kind: 'mutation' | 'selection' = 'selection',
      onDropped?: () => void,
      inTable?: () => boolean,
      openingInput = kind === 'mutation'
    ): void => {
      sealInputBatches();
      const replay = heldReplayBatchRef.current;
      if (replay) {
        try {
          replay.operations.push(Promise.resolve(operation(false)));
        } catch (error) {
          replay.operations.push(Promise.reject(error));
        }
        return;
      }
      if (inTable) pendingCaretTableRef.current = inTable;
      else if (!inputOperationQueueRef.current?.hasPending()) {
        const current = session?.selection();
        const currentInTable = current ? !!yrsCellLocFromStory(current.head.story) : false;
        pendingCaretTableRef.current = () => currentInTable;
      }
      const admitted = session;
      const queue = inputOperationQueueRef.current;
      const replica =
        admitted && replicaReadyRef?.current === false
          ? awaitWorkerOpenReplica(admitted)
          : undefined;
      if (replica && openingInput && !readOnly) requestOpeningPeer();
      if (!replica) {
        const apply = () => {
          if (!isCurrentInput(admitted, queue)) return onDropped?.();
          if (replicaReadyRef?.current !== false) replicaInputRef.current.applyPendingSelection?.();
          return operation(false);
        };
        queue?.enqueue(() => pendingSelectionsRef.current.length > 0
          ? preparePendingSelections(admitted, queue).then(apply)
          : apply());
        return;
      }
      const epoch = replicaInputRef.current.inputEpoch?.();
      const readerScrolled = watchReaderScroll();
      queue?.enqueue(async () => {
        if (!isCurrentInput(admitted, queue)) {
          readerScrolled();
          return onDropped?.();
        }
        try {
          await replica;
        } catch (error) {
          readerScrolled();
          onDropped?.();
          if (kind === 'mutation' && isCurrentInput(admitted, queue)) throw error;
          return;
        }
        const scrolled = readerScrolled();
        const superseded = replicaInputRef.current.inputEpoch?.() !== epoch;
        if (!isCurrentInput(admitted, queue) || (readOnly && kind === 'selection' && superseded)) {
          onDropped?.();
          return;
        }
        if (pendingSelectionsRef.current.length > 0) await preparePendingSelections(admitted, queue);
        if (!isCurrentInput(admitted, queue)) return onDropped?.();
        replicaInputRef.current.applyPendingSelection?.();
        await operation(true);
        if (scrolled && isCurrentInput(admitted, queue)) {
          quietSelectionRef.current = admitted?.selection() ?? null;
        }
      });
    },
    [isCurrentInput, preparePendingSelections, readOnly, replicaReadyRef, requestOpeningPeer, sealInputBatches, session, watchReaderScroll]
  );

  const replaySelection = useCallback((entry: HeldSelection): void => {
    enqueueInputOperation(() => entry.apply?.(), 'mutation', undefined, entry.inTable, false);
  }, [enqueueInputOperation]);
  const queueSelection = useCallback(
    (prepare: HeldSelection['prepare'], force = false, inTable?: () => boolean): boolean => {
      if (!force && !holdInput && replicaReadyRef?.current !== false && !inputOperationQueueRef.current?.hasPending()) return false;
      if (readOnly) return true;
      const entry: HeldSelection = { kind: 'selection', prepare, inTable };
      pendingSelectionsRef.current.push(entry);
      if (!holdOperation(entry)) replaySelection(entry);
      return true;
    },
    [holdInput, holdOperation, readOnly, replaySelection, replicaReadyRef]
  );

  const advanceInteractionEpoch = useCallback((): void => {
    inputOperationQueueRef.current?.advanceInteractionEpoch();
  }, []);

  const suggestingAuthor = useCallback((): YrsAuthor | undefined => {
    return isSuggesting ? { name: author, date: new Date().toISOString() } : undefined;
  }, [author, isSuggesting]);

  const belongsToRootStory = useCallback(
    (candidate: string): boolean => candidate === story || candidate.startsWith(`${story}:`),
    [story]
  );

  const readSelection = useCallback((): YrsSelection | null => {
    if (!session) return null;
    const current = session.selection();
    const currentMap =
      current &&
      current.anchor.story === current.head.story &&
      belongsToRootStory(current.anchor.story) &&
      belongsToRootStory(current.head.story)
        ? inputPositionMap(current.anchor.story)
        : null;
    if (
      current &&
      currentMap?.paragraphs.some((paragraph) => paragraph.paraId === current.anchor.paraId) &&
      currentMap.paragraphs.some((paragraph) => paragraph.paraId === current.head.paraId)
    ) {
      return current;
    }
    return null;
  }, [belongsToRootStory, inputPositionMap, session]);

  const ensureSelection = useCallback((): YrsSelection | null => {
    if (!session) return null;
    const current = readSelection();
    if (current) return current;
    const first = inputPositionMap(story)?.paragraphs[0];
    if (!first) return null;
    const loc = { story, paraId: first.paraId, offset: 0 };
    session.setSelection(loc);
    return { anchor: loc, head: loc };
  }, [inputPositionMap, readSelection, session, story]);

  const displaySelection = useCallback((): YrsDisplaySelection | null => {
    if (holdInput) return null;
    const current = seedSelection ? ensureSelection() : readSelection();
    if (!current) return null;
    const anchor = locToDisplayPosition(current.anchor);
    const head = locToDisplayPosition(current.head);
    return anchor == null || head == null ? null : { anchor, head };
  }, [ensureSelection, holdInput, locToDisplayPosition, readSelection, seedSelection]);

  const emitSelection = useCallback(
    (docChanged: boolean, residentLayoutReady = false, residentCaretReady = false, inWorker = false): void => {
      if (heldReplayBatchRef.current) {
        heldReplayBatchRef.current.selectionChanged = true;
        return;
      }
      const selection = displaySelection();
      if (!selection) return;
      setSelectionEpoch((epoch) => epoch + 1);
      if (inWorker) onStateChange(selection, docChanged, residentLayoutReady, residentCaretReady, 'local', true);
      else onStateChange(selection, docChanged, residentLayoutReady, residentCaretReady);
    },
    [displaySelection, onStateChange]
  );

  const requestLocalCaretReveal = useCallback((): void => {
    const host = canvasHostRef?.current;
    const scroller = host ? findVerticalScrollParentOrRoot(host) : null;
    pendingLocalCaretRevealRef.current = scroller ? { scroller, scrollTop: scroller.scrollTop } : null;
  }, [canvasHostRef]);

  const setSelection = useCallback(
    (anchor: YrsLoc, head: YrsLoc = anchor, emit = true): void => {
      if (!session) return;
      session.setSelection(anchor, head);
      if (emit) {
        if (!readOnly) requestLocalCaretReveal();
        onCaretInterrupt?.();
        emitSelection(false);
      }
    },
    [emitSelection, onCaretInterrupt, readOnly, requestLocalCaretReveal, session]
  );

  const finishMutation = useCallback(
    (
      residentLayoutReady = false,
      residentCaretReady = false,
      dirtyStories?: string | readonly string[]
    ): void => {
      verticalCaretGoalRef.current.reset();
      if (residentLayoutReady) requestLocalCaretReveal();
      if (!composingRef.current && textareaRef.current) textareaRef.current.value = '';
      if (heldReplayBatchRef.current) {
        onDirectInput(dirtyStories);
        heldReplayBatchRef.current.mutated = true;
        return;
      }
      onCaretInput?.();
      onDirectInput(dirtyStories);
      emitSelection(true, residentLayoutReady, residentCaretReady);
    },
    [emitSelection, onCaretInput, onDirectInput, requestLocalCaretReveal]
  );

  const finishResidentMutation = useCallback(
    (result: ResidentFrameApplyResult): void => {
      if (result.frameEpoch !== null) {
        pendingResidentFrameEpochRef.current = result.frameEpoch;
      }
      finishMutation(result.frameEpoch !== null, result.caretSynchronized);
    },
    [finishMutation]
  );

  // Body-story only: painted-caret coverage for other stories is unproven, and
  // an unhonored dispatch hold would blank the caret per keystroke there.
  const dispatchCaretInput = useCallback((): void => {
    if (heldReplayBatchRef.current) return;
    if (!session || readOnly || replicaReadyRef?.current === false) return;
    if (session.selection()?.head.story !== 'body') return;
    onCaretInputDispatched?.();
  }, [onCaretInputDispatched, readOnly, session, replicaReadyRef]);

  const storedFormatting = useCallback((): YrsStoredFormatting | null => {
    const current = seedSelection ? ensureSelection() : readSelection();
    if (!current) return null;
    return (
      storedFormattingByParagraphRef.current.get(
        `${current.head.story}\u0000${current.head.paraId}`
      ) ?? null
    );
  }, [ensureSelection, readSelection, seedSelection]);

  const applyStoredFormatting = useCallback(
    (action: YrsStoredFormattingAction): void => {
      const selection = ensureSelection();
      if (!selection) return;
      const key = `${selection.head.story}\u0000${selection.head.paraId}`;
      if (action.type === 'clear') {
        storedFormattingByParagraphRef.current.set(key, { clear: true, delta: {} });
        emitSelection(false);
        return;
      }
      const current = storedFormattingByParagraphRef.current.get(key) ?? {
        clear: false,
        delta: {},
      };
      if (action.type === 'set') {
        storedFormattingByParagraphRef.current.set(key, {
          clear: current.clear,
          delta: { ...current.delta, ...action.delta },
        });
        emitSelection(false);
        return;
      }
      if (action.mark === 'superscript' || action.mark === 'subscript') {
        const other = current.delta.other ?? {};
        const stored = other[action.mark];
        const isActive =
          stored === undefined ? (current.clear ? false : action.active) : stored === true;
        const counterpart = action.mark === 'superscript' ? 'subscript' : 'superscript';
        storedFormattingByParagraphRef.current.set(key, {
          clear: current.clear,
          delta: {
            ...current.delta,
            other: isActive
              ? { ...other, [action.mark]: null }
              : { ...other, [action.mark]: true, [counterpart]: null },
          },
        });
        emitSelection(false);
        return;
      }
      const storedValue = current.delta[action.mark];
      const isActive =
        storedValue === undefined
          ? current.clear
            ? false
            : action.active
          : storedValue !== false && storedValue !== null;
      storedFormattingByParagraphRef.current.set(key, {
        clear: current.clear,
        delta: { ...current.delta, [action.mark]: !isActive },
      });
      emitSelection(false);
    },
    [emitSelection, ensureSelection]
  );

  const deleteSelected = useCallback((): YrsLoc | null => {
    if (!session) return null;
    const current = ensureSelection();
    const map = current ? inputPositionMap(current.anchor.story) : null;
    if (!current || !map) return null;
    const range = toRange(current, map);
    if (range.start.paraId === range.end.paraId && range.start.offset === range.end.offset) {
      return null;
    }
    const landed = session.deleteRange(range, suggestingAuthor()).range;
    const collapsed = landed
      ? { story: landed.story, ...landed.start }
      : { story: range.story, ...range.start };
    session.setSelection(collapsed);
    return collapsed;
  }, [ensureSelection, inputPositionMap, session, suggestingAuthor]);

  const insertText = useCallback(
    (text: string, inputTime?: number, composition = false): void => {
      verticalCaretGoalRef.current.reset();
      if (!session || readOnly || text.length === 0) return;
      if (holdOperation({ kind: composition ? 'composition' : 'text', text, inputTime })) return;
      dispatchCaretInput();
      const applyText = async (inputText: string) => {
        const current = ensureSelection();
        const map = current ? inputPositionMap(current.anchor.story) : null;
        if (!current || !map) return;
        const selectedRange = toRange(current, map);
        const hasSelection =
          selectedRange.start.paraId !== selectedRange.end.paraId ||
          selectedRange.start.offset !== selectedRange.end.offset;
        const stored = storedFormattingByParagraphRef.current.get(
          `${current.head.story}\u0000${current.head.paraId}`
        );
        const commitCompatibilityInput = (): void => {
          const at = hasSelection
            ? { story: selectedRange.story, ...selectedRange.start }
            : current.head;
          // beforeinput may surface pasted line endings as text. Preserve the
          // structural contract by splitting those instead of inserting pilcrows.
          const pieces = inputText.replace(/\r\n?/g, '\n').split('\n');
          let caret = at;
          const storedKey = (loc: YrsLoc): string => `${loc.story}\u0000${loc.paraId}`;
          // A suggested replacement's text lands after the struck-out text, possibly
          // in another paragraph: the head's stored formatting goes with it.
          const carried = hasSelection && isSuggesting ? stored : undefined;
          for (let i = 0; i < pieces.length; i += 1) {
            const piece = pieces[i];
            if (piece || (i === 0 && hasSelection)) {
              let insertedAt = caret;
              if (i === 0 && hasSelection) {
                const receipt = session.replaceRange(selectedRange, piece, suggestingAuthor());
                if (receipt.range) {
                  insertedAt = { story: receipt.range.story, ...receipt.range.start };
                  caret = { story: receipt.range.story, ...receipt.range.end };
                } else {
                  caret = { ...caret, offset: caret.offset + piece.length };
                }
              } else {
                const landed = session.insertText(caret, piece, suggestingAuthor()).range;
                insertedAt = landed ? { story: landed.story, ...landed.start } : caret;
                caret = landed
                  ? { story: landed.story, ...landed.end }
                  : { ...caret, offset: caret.offset + piece.length };
              }
              const insertedStored =
                carried ?? storedFormattingByParagraphRef.current.get(storedKey(insertedAt));
              if (insertedStored && piece) {
                const insertedRange: YrsStoryRange = {
                  story: insertedAt.story,
                  start: { paraId: insertedAt.paraId, offset: insertedAt.offset },
                  end: { paraId: caret.paraId, offset: caret.offset },
                };
                if (insertedStored.clear) session.clearFormatting(insertedRange);
                if (Object.keys(insertedStored.delta).length > 0) {
                  session.formatRange(insertedRange, insertedStored.delta);
                }
              }
            }
            if (i < pieces.length - 1) {
              const receipt = session.splitParagraph(caret, suggestingAuthor());
              caret = { story: caret.story, paraId: receipt.secondParaId, offset: 0 };
            }
          }
          if (carried && !storedFormattingByParagraphRef.current.has(storedKey(caret))) {
            storedFormattingByParagraphRef.current.set(storedKey(caret), carried);
          }
          session.setSelection(caret);
          finishMutation();
        };
        if (
          !heldReplayBatchRef.current &&
          !hasSelection &&
          isBodyFlowStory(current.head.story) &&
          !isSuggesting &&
          !stored &&
          /^[\x20-\x7e]+$/u.test(inputText) &&
          !inputText.includes('\r') &&
          !inputText.includes('\n') &&
          applyResidentInput
        ) {
          const applied = await applyResidentInput(inputText);
          if (!isCurrentInput(session)) return;
          if (applied) finishResidentMutation(applied);
          else commitCompatibilityInput();
          return;
        }
        commitCompatibilityInput();
      };

      if (composition && replicaReadyRef) {
        enqueueInputOperation(async () => {
          session.addUndoBoundary();
          try {
            await applyText(text);
          } finally {
            if (isCurrentInput(session)) session.addUndoBoundary();
          }
        }, 'mutation');
        return;
      }

      const canBatchResidentText =
        !heldReplayBatchRef.current &&
        !isSuggesting &&
        Boolean(applyResidentInput) &&
        /^[\x20-\x7e]+$/u.test(text) &&
        !text.includes('\r') &&
        !text.includes('\n');
      if (!canBatchResidentText) {
        enqueueInputOperation(() => applyText(text), 'mutation');
        return;
      }

      const pending = pendingResidentTextRef.current;
      if (pending) {
        pending.text += text;
        return;
      }

      const batch = { text };
      enqueueInputOperation(async () => {
        if (pendingResidentTextRef.current === batch) pendingResidentTextRef.current = null;
        await applyText(batch.text);
      }, 'mutation');
      pendingResidentTextRef.current = batch;
    },
    [
      applyResidentInput,
      dispatchCaretInput,
      enqueueInputOperation,
      ensureSelection,
      finishMutation,
      finishResidentMutation,
      holdOperation,
      inputPositionMap,
      isCurrentInput,
      isSuggesting,
      readOnly,
      replicaReadyRef,
      session,
      suggestingAuthor,
    ]
  );

  const deleteUnits = useCallback(
    async (direction: 'backward' | 'forward', count: number): Promise<void> => {
      if (!session || readOnly || replicaReadyRef?.current === false) return;
      let remaining = count;
      while (remaining > 0) {
        if (deleteSelected()) {
          finishMutation();
          remaining -= 1;
          continue;
        }
        const current = ensureSelection();
        const activeStory = current?.head.story;
        const map = activeStory ? inputPositionMap(activeStory) : null;
        if (!current || !activeStory || !map) return;
        const caret = current.head;
        const mapIndex = map.paragraphs.findIndex((entry) => entry.paraId === caret.paraId);
        const hasTarget =
          mapIndex >= 0 &&
          (direction === 'backward'
            ? caret.offset > 0 || mapIndex > 0
            : caret.offset < map.paragraphs[mapIndex].length ||
              mapIndex + 1 < map.paragraphs.length);
        if (!heldReplayBatchRef.current && hasTarget && isBodyFlowStory(activeStory) && !isSuggesting && applyResidentDelete) {
          const applied = await applyResidentDelete(direction, remaining);
          if (!isCurrentInput(session)) return;
          if (applied) {
            finishResidentMutation(applied);
            remaining -= Math.max(1, applied.deletedUnits ?? remaining);
            continue;
          }
        }
        const paragraphs = session.paragraphs(activeStory);
        const index = paragraphs.findIndex((paragraph) => paragraph.paraId === caret.paraId);
        if (index < 0) return;
        const paragraph = paragraphs[index];
        if (direction === 'backward') {
          if (caret.offset > 0) {
            const start = previousCodePointOffset(paragraph.text, caret.offset);
            const landed = session.deleteRange(
              {
                story: activeStory,
                start: { paraId: caret.paraId, offset: start },
                end: { paraId: caret.paraId, offset: caret.offset },
              },
              suggestingAuthor()
            ).range;
            session.setSelection(
              landed ? { story: landed.story, ...landed.start } : { ...caret, offset: start }
            );
          } else if (index > 0) {
            const previous = paragraphs[index - 1];
            const offset = inputPositionMap(activeStory)?.paragraphs.find(
              (entry) => entry.paraId === previous.paraId
            )?.length;
            session.mergeParagraphs(activeStory, previous.paraId, suggestingAuthor());
            session.setSelection({
              story: activeStory,
              paraId: previous.paraId,
              offset: offset ?? previous.text.length,
            });
          } else {
            return;
          }
        } else {
          const length = map.paragraphs.find((entry) => entry.paraId === caret.paraId)?.length ?? 0;
          if (caret.offset < length) {
            const end = nextCodePointOffset(paragraph.text, caret.offset);
            const landed = session.deleteRange(
              {
                story: activeStory,
                start: { paraId: caret.paraId, offset: caret.offset },
                end: { paraId: caret.paraId, offset: end },
              },
              suggestingAuthor()
            ).range;
            session.setSelection(landed ? { story: landed.story, ...landed.start } : caret);
          } else if (index + 1 < paragraphs.length) {
            session.mergeParagraphs(activeStory, caret.paraId, suggestingAuthor());
            session.setSelection(caret);
          } else {
            return;
          }
        }
        finishMutation();
        remaining -= 1;
      }
    },
    [
      applyResidentDelete,
      deleteSelected,
      ensureSelection,
      finishMutation,
      finishResidentMutation,
      inputPositionMap,
      isCurrentInput,
      isSuggesting,
      readOnly,
      replicaReadyRef,
      session,
      suggestingAuthor,
    ]
  );

  // Deletes queued behind busy input join one batch, applied with one layout.
  const deleteDirection = useCallback(
    (direction: 'backward' | 'forward'): void => {
      if (holdOperation({ kind: 'delete', direction })) return;
      verticalCaretGoalRef.current.reset();
      dispatchCaretInput();
      if (heldReplayBatchRef.current) {
        enqueueInputOperation(() => deleteUnits(direction, 1), 'mutation');
        return;
      }
      const pending = pendingResidentDeleteRef.current;
      if (pending?.direction === direction) {
        pending.count += 1;
        return;
      }
      const batch = { direction, count: 1 };
      enqueueInputOperation(async () => {
        if (pendingResidentDeleteRef.current === batch) pendingResidentDeleteRef.current = null;
        await deleteUnits(batch.direction, batch.count);
      }, 'mutation');
      pendingResidentDeleteRef.current = batch;
    },
    [deleteUnits, dispatchCaretInput, enqueueInputOperation, holdOperation]
  );

  const splitParagraph = useCallback((): void => {
    if (holdOperation({ kind: 'split' })) return;
    verticalCaretGoalRef.current.reset();
    dispatchCaretInput();
    enqueueInputOperation(() => {
      if (!session || readOnly || replicaReadyRef?.current === false) return;
      const selectedStart = deleteSelected();
      const current = selectedStart ?? ensureSelection()?.head;
      if (!current) return;
      const currentParagraph = session
        .paragraphs(current.story)
        .find((paragraph) => paragraph.paraId === current.paraId);
      const inheritedStored = storedFormattingByParagraphRef.current.get(
        `${current.story}\u0000${current.paraId}`
      );
      const receipt = session.splitParagraph(current, suggestingAuthor());
      const currentStyleId =
        typeof currentParagraph?.properties.pStyle === 'string'
          ? currentParagraph.properties.pStyle
          : null;
      const nextStyleId =
        currentParagraph && current.offset === currentParagraph.text.length
          ? (nextParagraphStyleId?.(currentStyleId) ?? null)
          : null;
      if (nextStyleId) {
        session.applyParagraphStyle(
          {
            story: current.story,
            start: { paraId: receipt.secondParaId, offset: 0 },
            end: { paraId: receipt.secondParaId, offset: 0 },
          },
          nextStyleId
        );
      } else if (currentParagraph?.text && inheritedStored) {
        storedFormattingByParagraphRef.current.set(
          `${current.story}\u0000${receipt.secondParaId}`,
          inheritedStored
        );
      }
      session.setSelection({
        story: current.story,
        paraId: receipt.secondParaId,
        offset: 0,
      });
      finishMutation();
    }, 'mutation');
  }, [
    deleteSelected,
    dispatchCaretInput,
    enqueueInputOperation,
    ensureSelection,
    finishMutation,
    holdOperation,
    nextParagraphStyleId,
    readOnly,
    replicaReadyRef,
    session,
    suggestingAuthor,
  ]);

  const moveSelection = useCallback(
    (
      direction: NavigationDirection,
      extend: boolean,
      wholeDocument: boolean,
      byWord = false,
      replayed = false
    ): void => {
      if (holdOperation({ kind: 'navigation', direction, extend, wholeDocument, byWord })) return;
      const verticalDirection = direction === 'up' || direction === 'down' ? direction : null;
      let interactionEpoch = verticalDirection
        ? inputOperationQueueRef.current?.captureInteractionEpoch()
        : undefined;
      enqueueInputOperation(async (waited) => {
        // The gesture applied before input that waited is older than it.
        if ((waited || replayed) && verticalDirection) {
          interactionEpoch = inputOperationQueueRef.current?.captureInteractionEpoch();
        }
        if (!verticalDirection) verticalCaretGoalRef.current.reset();
        if (!session) return;
        if (replayed && !ensureSelection()) return;
        const queryResolver = resolveDisplayListQueriesRef.current;
        const currentQueries = displayListQueriesRef.current;
        const querySnapshot = verticalDirection
          ? queryResolver
            ? await queryResolver(pendingResidentFrameEpochRef.current)
            : currentQueries
              ? { queries: currentQueries, frameEpoch: displayListFrameEpochRef.current }
              : null
          : null;
        if (!isCurrentInput(session)) return;
        if (
          verticalDirection &&
          interactionEpoch !== undefined &&
          !inputOperationQueueRef.current?.isInteractionEpochCurrent(interactionEpoch)
        ) {
          return;
        }
        const current = ensureSelection();
        const activeStory = current?.head.story;
        const map = activeStory ? inputPositionMap(activeStory) : null;
        if (!current || !activeStory || !map || map.paragraphs.length === 0) return;
        const collapsed =
          current.anchor.paraId === current.head.paraId &&
          current.anchor.offset === current.head.offset;
        const ordered = toRange(current, map);
        if (!extend && !collapsed && (direction === 'left' || direction === 'right')) {
          const edge = direction === 'left' ? ordered.start : ordered.end;
          setSelection({ story: activeStory, ...edge });
          return;
        }

        const head = current.head;
        if (verticalDirection) {
          const displayPosition = locToDisplayPosition(head);
          const movement =
            displayPosition == null
              ? null
              : querySnapshot?.queries.verticalMove(
                  displayPosition,
                  verticalDirection,
                  verticalCaretGoalRef.current.current()
                );
          if (movement) {
            verticalCaretGoalRef.current.retain(movement.goalX);
            const target = resolveDisplayTarget?.(movement.position) ?? {
              story: activeStory,
              displayPosition: movement.position,
            };
            const next = displayPositionToLoc(target.displayPosition, target.story);
            if (!next || (extend && current.anchor.story !== next.story)) return;
            if (
              next.story === head.story &&
              next.paraId === head.paraId &&
              next.offset === head.offset
            ) {
              if (!readOnly) setSelection(current.anchor, current.head);
              return;
            }
            setSelection(extend ? current.anchor : next, next);
            return;
          }
          const next = paragraphVerticalMove(map.paragraphs, head, verticalDirection);
          setSelection(extend ? current.anchor : next, next);
          return;
        }
        const index = map.paragraphs.findIndex((entry) => entry.paraId === head.paraId);
        if (index < 0) return;
        const entry = map.paragraphs[index];
        const text = session.paragraphs(activeStory)[index]?.text ?? '';
        let next = head;
        if (direction === 'home') {
          const target = wholeDocument ? map.paragraphs[0] : entry;
          next = { story: activeStory, paraId: target.paraId, offset: 0 };
        } else if (direction === 'end') {
          const target = wholeDocument ? map.paragraphs[map.paragraphs.length - 1] : entry;
          next = { story: activeStory, paraId: target.paraId, offset: target.length };
        } else if (direction === 'left') {
          if (head.offset > 0)
            next = {
              ...head,
              offset: byWord
                ? previousWordOffset(text, head.offset)
                : previousCodePointOffset(text, head.offset),
            };
          else if (index > 0) {
            const target = map.paragraphs[index - 1];
            next = { story: activeStory, paraId: target.paraId, offset: target.length };
          }
        } else if (direction === 'right') {
          if (head.offset < entry.length)
            next = {
              ...head,
              offset: byWord
                ? nextWordOffset(text, head.offset)
                : nextCodePointOffset(text, head.offset),
            };
          else if (index + 1 < map.paragraphs.length) {
            next = { story: activeStory, paraId: map.paragraphs[index + 1].paraId, offset: 0 };
          }
        }
        setSelection(extend ? current.anchor : next, next);
      });
    },
    [
      displayPositionToLoc,
      enqueueInputOperation,
      ensureSelection,
      holdOperation,
      inputPositionMap,
      isCurrentInput,
      locToDisplayPosition,
      readOnly,
      resolveDisplayTarget,
      session,
      setSelection,
    ]
  );

  const selectAll = useCallback((): void => {
    if (holdOperation({ kind: 'select-all' })) return;
    enqueueInputOperation(() => {
      verticalCaretGoalRef.current.reset();
      const current = ensureSelection();
      // read-only select all takes the whole document, not the table cell holding the caret
      const activeStory = readOnly ? story : current?.head.story;
      const map = activeStory ? inputPositionMap(activeStory) : null;
      if (!session || !activeStory || !map || map.paragraphs.length === 0) return;
      const first = map.paragraphs[0];
      const last = map.paragraphs[map.paragraphs.length - 1];
      setSelection(
        { story: activeStory, paraId: first.paraId, offset: 0 },
        { story: activeStory, paraId: last.paraId, offset: last.length }
      );
    });
  }, [enqueueInputOperation, ensureSelection, holdOperation, inputPositionMap, readOnly, session, setSelection, story]);

  const deleteSelection = useCallback(
    (queued = false, cut?: CutCopy): void => {
      if (holdOperation({ kind: 'delete-selection', cut })) return;
      const apply = (): void => {
        if (cut && !cut.written) return;
        if (!readOnly && replicaReadyRef?.current !== false && deleteSelected()) {
          advanceInteractionEpoch();
          finishMutation();
        }
      };
      if (queued) enqueueInputOperation(apply, 'mutation');
      else apply();
    },
    [
      advanceInteractionEpoch,
      deleteSelected,
      enqueueInputOperation,
      finishMutation,
      holdOperation,
      readOnly,
      replicaReadyRef,
    ]
  );

  const moveTableCell = useCallback(
    (backward: boolean, apply = true): boolean => {
      if (!session) return false;
      const current = ensureSelection();
      const focused = current ? yrsCellLocFromStory(current.head.story) : null;
      if (!current || !focused) return false;
      const tableRange = yrsTableSelectionRange(session, focused, 'table');
      if (!tableRange) return false;
      if (!apply) return true;
      verticalCaretGoalRef.current.reset();

      let row = focused.row;
      let column = focused.column + (backward ? -1 : 1);
      const lastRow = tableRange.head.row;
      const lastColumn = tableRange.head.column;
      if (column > lastColumn) {
        row += 1;
        column = 0;
      } else if (column < 0) {
        row -= 1;
        column = lastColumn;
      }

      if (row < 0 || row > lastRow) {
        if (backward) {
          setSelection(current.anchor, current.head);
          return true;
        }
        const nearby = yrsSelectionNearTable(session, {
          story: focused.story,
          tableIndex: focused.tableIndex,
        });
        if (nearby) {
          setSelection(nearby);
          return true;
        }
        // Terminal-Tab behavior when the document has no trailing paragraph:
        // append a row and enter its first cell.
        session.insertRow(focused, 'below');
        row = lastRow + 1;
        column = 0;
      }

      const next = { ...focused, row, column };
      const nextStory = yrsCellStory(session, next);
      const paragraph = nextStory ? session.paragraphs(nextStory)[0] : null;
      if (!nextStory || !paragraph) return true;
      session.setCellSelection({ anchor: next, head: next });
      setSelection({ story: nextStory, paraId: paragraph.paraId, offset: 0 });
      return true;
    },
    [ensureSelection, session, setSelection]
  );

  const handleTab = useCallback((shift: boolean, replayed = false): boolean => {
    if (readOnly) return false;
    if (!replayed) {
      const pending = inputOperationQueueRef.current?.hasPending();
      if (holdInput || replicaReadyRef?.current === false) {
        const last = heldInputRef.current.entries.slice().reverse().find(
          (entry): entry is HeldSelection => entry.kind === 'selection' && !!entry.inTable
        );
        const current = !last && !pending ? readSelection() : null;
        const inTable = last
          ? last.inTable?.()
          : pending ? pendingCaretTableRef.current?.()
          : current ? !!yrsCellLocFromStory(current.head.story) : undefined;
        if (!inTable) return false;
        if (holdOperation({ kind: 'tab', shift })) return true;
      } else if (!moveTableCell(shift, false)) return false;
    }
    enqueueInputOperation(() => {
      moveTableCell(shift);
    }, 'mutation', undefined, undefined, false);
    return true;
  }, [enqueueInputOperation, holdInput, holdOperation, moveTableCell, readOnly, readSelection, replicaReadyRef]);

  const inputHandlersRef = useRef({ insertText, splitParagraph, deleteDirection, deleteSelection, selectAll, replaySelection, moveSelection, handleTab, enqueueInputOperation, session });
  inputHandlersRef.current = { insertText, splitParagraph, deleteDirection, deleteSelection, selectAll, replaySelection, moveSelection, handleTab, enqueueInputOperation, session };
  const replayHeldEntry = useCallback((entry: HeldInput, handlers = inputHandlersRef.current): void => {
    if (entry.kind === 'selection') handlers.replaySelection(entry);
    else if (entry.kind === 'navigation') handlers.moveSelection(entry.direction, entry.extend, entry.wholeDocument, entry.byWord, true);
    else if (entry.kind === 'text' || entry.kind === 'composition') handlers.insertText(entry.text);
    else if (entry.kind === 'tab') handlers.handleTab(entry.shift, true);
    else if (entry.kind === 'split') handlers.splitParagraph();
    else if (entry.kind === 'delete') handlers.deleteDirection(entry.direction);
    else if (entry.kind === 'delete-selection') handlers.deleteSelection(true, entry.cut);
    else if (entry.kind === 'select-all') handlers.selectAll();
    else if (entry.kind === 'copy') handlers.enqueueInputOperation(
      () => handlers.session ? entry.apply(handlers.session) : undefined,
      'selection', entry.onDropped, undefined, false
    );
  }, []);
  const replayHeldBatch = useCallback((entries: HeldInput[]): void => {
    const handlers = inputHandlersRef.current;
    const last = entries.slice().reverse().find(
      (entry): entry is HeldSelection => entry.kind === 'selection' && !!entry.inTable
    );
    let prepareFailure: { error: unknown } | undefined;
    for (const entry of entries) {
      if (entry.kind !== 'selection') continue;
      const prepare = entry.prepare;
      entry.prepare = async () => {
        try {
          return await prepare();
        } catch (error) {
          prepareFailure ??= { error };
          return () => {};
        }
      };
    }
    enqueueInputOperation(async () => {
      if (!session || readOnly) return;
      const batch: HeldReplayBatch = { operations: [], mutated: false, selectionChanged: false };
      const autoCapture = session.undoCaptureMode() === 'auto';
      session.beginUndoCapture();
      session.addUndoBoundary();
      heldReplayBatchRef.current = batch;
      let previousTime: number | undefined;
      let previousStory = session.selection()?.head.story;
      try {
        for (const entry of entries) {
          const activeStory = ensureSelection()?.head.story;
          if (
            (autoCapture && previousTime !== undefined && entry.inputTime! - previousTime >= UNDO_CAPTURE_TIMEOUT_MS) ||
            activeStory !== previousStory || entry.kind === 'selection' ||
            entry.kind === 'navigation' || entry.kind === 'select-all' ||
            entry.kind === 'composition' || entry.kind === 'undo-boundary'
          ) session.addUndoBoundary();
          const version = entry.kind === 'tab' ? session.version() : null;
          replayHeldEntry(entry, handlers);
          if (entry.kind === 'composition') session.addUndoBoundary();
          if (entry.kind !== 'tab' || session.version() !== version) previousTime = entry.inputTime;
          previousStory = activeStory;
        }
      } finally {
        heldReplayBatchRef.current = null;
        if (batch.mutated) {
          onCaretInput?.();
          emitSelection(true, false, false, true);
        } else if (batch.selectionChanged) emitSelection(false);
      }
      await Promise.all(batch.operations);
      if (prepareFailure) throw prepareFailure.error;
    }, 'mutation', undefined, last?.inTable, entries.some(isOpeningHeldInput));
  }, [emitSelection, enqueueInputOperation, ensureSelection, onCaretInput, readOnly, replayHeldEntry, replicaReadyRef, session]);
  const heldReplayHandlersRef = useRef({ replayHeldBatch, replayHeldEntry, enqueueInputOperation });
  heldReplayHandlersRef.current = { replayHeldBatch, replayHeldEntry, enqueueInputOperation };
  useLayoutEffect(() => {
    const { replayHeldBatch, replayHeldEntry, enqueueInputOperation } = heldReplayHandlersRef.current;
    if ((readOnly || !enabled || (!holdInput && !session)) &&
      (heldInputRef.current.entries.length > 0 || pendingSelectionsRef.current.length > 0 || compositionHeldRef.current)) {
      for (const entry of heldInputRef.current.entries) if (entry.kind === 'copy') entry.onDropped();
      heldInputRef.current = { scope: inputScope, entries: [] };
      pendingSelectionsRef.current = [];
      if (compositionHeldRef.current) {
        discardedCompositionRef.current = true;
        composingRef.current = false;
        compositionPendingRef.current = false;
        compositionCommitRef.current = '';
        compositionHeldRef.current = false;
        if (textareaRef.current) textareaRef.current.value = '';
        for (const resolve of compositionWaitersRef.current) resolve();
        compositionWaitersRef.current.clear();
      }
      onPendingInputChangeRef.current?.(inputOperationQueueRef.current?.hasPending() ?? false);
    } else if (!holdInput && !readOnly && enabled && session) {
      const entries = heldInputRef.current.entries;
      heldInputRef.current.entries = [];
      let batch: HeldInput[] = [];
      for (const entry of entries) {
        if (!replicaReadyRef) {
          replayHeldEntry(entry);
        } else if (entry.kind === 'copy') {
          if (batch.length) replayHeldBatch(batch);
          batch = [];
          replayHeldEntry(entry);
        } else if (entry.kind === 'navigation' && (entry.direction === 'up' || entry.direction === 'down')) {
          if (batch.length) replayHeldBatch(batch);
          batch = [];
          enqueueInputOperation(() => session.addUndoBoundary(), 'mutation', undefined, undefined, false);
          replayHeldEntry(entry);
        } else batch.push(entry);
      }
      if (batch.length) replayHeldBatch(batch);
    }
    for (const notify of heldInputWaitersRef.current) notify();
  }, [enabled, holdInput, inputScope, readOnly, session]);

  const handleBeforeInput = useCallback(
    (event: React.FormEvent<HTMLTextAreaElement>): void => {
      const native = event.nativeEvent as InputEvent;
      if (discardedCompositionRef.current) {
        event.preventDefault();
        event.currentTarget.value = '';
        return;
      }
      if (composingRef.current || native.isComposing) return;
      if (compositionPendingRef.current) {
        event.preventDefault();
        if (native.data) compositionCommitRef.current = native.data;
        return;
      }
      if (native.inputType === 'insertText' || native.inputType === 'insertReplacementText') {
        event.preventDefault();
        insertText(native.data ?? '');
      } else if (native.inputType === 'insertParagraph' || native.inputType === 'insertLineBreak') {
        event.preventDefault();
        splitParagraph();
      } else if (native.inputType === 'deleteContentBackward') {
        event.preventDefault();
        deleteDirection('backward');
      } else if (native.inputType === 'deleteContentForward') {
        event.preventDefault();
        deleteDirection('forward');
      }
    },
    [deleteDirection, insertText, splitParagraph]
  );

  // Browsers only run a copy shortcut over a non-empty native selection, so
  // the textarea holds the selected text until the shortcut has run.
  const primedCopyRef = useRef<string | null>(null);
  const primeCopy = useCallback((): void => {
    const textarea = textareaRef.current;
    const text = session ? yrsSelectionPlainText(session) : '';
    if (!textarea || !text || textarea.value) return;
    primedCopyRef.current = text;
    textarea.value = text;
    textarea.select();
    setTimeout(() => {
      if (primedCopyRef.current === text) primedCopyRef.current = null;
      if (textarea.value === text) textarea.value = '';
    });
  }, [session]);

  // A copy asked before the edit peer loads writes the selection it then has.
  const copyAfterReplica = useCallback((cut = false): CutCopy | null => {
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
    const held = holdInput && !readOnly && (cut || heldInputRef.current.entries.length > 0);
    const queued = held || replicaReadyRef?.current === false || !!inputOperationQueueRef.current?.hasPending();
    if (!session || !queued || !clipboard) {
      return null;
    }
    const result: CutCopy = { written: false };
    const copied = session;
    let written: Promise<void>;
    const text = new Promise<Blob>((resolve, reject) => {
      const apply = (source: YrsSession): void | Promise<void> => {
        const selected = yrsSelectionPlainText(source);
        if (selected) resolve(new Blob([selected], { type: 'text/plain' }));
        else reject(new Error('Nothing is selected to copy'));
        if (cut) return written;
      };
      const onDropped = () => reject(new Error('Newer input replaced the copy'));
      if (!held || !holdOperation({ kind: 'copy', apply, onDropped })) {
        enqueueInputOperation(() => apply(copied), 'selection', onDropped, undefined, false);
      }
    });
    void text.catch(() => {});
    const write = typeof ClipboardItem === 'function' && clipboard.write
      ? clipboard.write([new ClipboardItem({ 'text/plain': text })])
      : text.then((blob) => blob.text()).then((value) => clipboard.writeText(value));
    written = write.then(() => { result.written = true; }, () => {});
    return result;
  }, [enqueueInputOperation, holdInput, holdOperation, readOnly, replicaReadyRef, session]);

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
      if (!event.nativeEvent.isComposing && event.key !== 'Process' && event.keyCode !== 229) {
        discardedCompositionRef.current = false;
      }
      if (event.nativeEvent.isComposing || composingRef.current) return;
      const mod = event.metaKey || event.ctrlKey;
      const key = event.key.toLowerCase();
      if (mod && key === 'a') {
        event.preventDefault();
        selectAll();
      } else if (mod && key === 'c' && !event.shiftKey && !event.altKey) {
        if (copyAfterReplica()) event.preventDefault();
        else primeCopy();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        splitParagraph();
      } else if (event.key === 'Tab' && handleTab(event.shiftKey)) {
        event.preventDefault();
      } else if (event.key === 'Backspace') {
        event.preventDefault();
        deleteDirection('backward');
      } else if (event.key === 'Delete') {
        event.preventDefault();
        deleteDirection('forward');
      } else if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
        event.preventDefault();
        moveSelection(
          event.key === 'ArrowLeft' ? 'left' : 'right',
          event.shiftKey,
          false,
          event.altKey || event.ctrlKey
        );
      } else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
        event.preventDefault();
        moveSelection(event.key === 'ArrowUp' ? 'up' : 'down', event.shiftKey, false);
      } else if (event.key === 'Home' || event.key === 'End') {
        event.preventDefault();
        moveSelection(event.key === 'Home' ? 'home' : 'end', event.shiftKey, mod);
      }
    },
    [
      copyAfterReplica,
      deleteDirection,
      handleTab,
      moveSelection,
      primeCopy,
      selectAll,
      splitParagraph,
    ]
  );

  const handleCompositionStart = useCallback(
    (event: React.CompositionEvent<HTMLTextAreaElement>) => {
      verticalCaretGoalRef.current.reset();
      composingRef.current = true;
      discardedCompositionRef.current = false;
      compositionHeldRef.current = holdInput;
      compositionPendingRef.current = false;
      compositionCommitRef.current = '';
      event.currentTarget.value = '';
      if (replicaReadyRef) holdOperation({ kind: 'undo-boundary' });
      if (!readOnly) requestOpeningPeer();
      onCaretInterrupt?.();
      onPendingInputChangeRef.current?.(true);
    },
    [holdInput, holdOperation, onCaretInterrupt, readOnly, replicaReadyRef, requestOpeningPeer]
  );

  const handleCompositionUpdate = useCallback(
    (event: React.CompositionEvent<HTMLTextAreaElement>) => {
      if (discardedCompositionRef.current) return;
      compositionCommitRef.current = event.data;
    },
    []
  );

  const handleCompositionEnd = useCallback(
    (event: React.CompositionEvent<HTMLTextAreaElement>) => {
      if (discardedCompositionRef.current) {
        event.currentTarget.value = '';
        return;
      }
      composingRef.current = false;
      compositionPendingRef.current = true;
      compositionCommitRef.current =
        event.currentTarget.value || event.data || compositionCommitRef.current;
      const scope = heldInputRef.current;
      const heldComposition = compositionHeldRef.current;
      const inputTime = heldComposition ? performance.now() : undefined;
      queueMicrotask(() => {
        const lifetime = inputLifetimeRef.current;
        if (
          !compositionPendingRef.current || !lifetime.mounted || !lifetime.enabled ||
          heldInputRef.current !== scope ||
          ((!heldComposition || inputScope === undefined) && !isCurrentInput(session))
        ) return;
        const text = textareaRef.current?.value || compositionCommitRef.current;
        // Reset the browser model before applying the document op. A trailing
        // post-composition beforeinput therefore observes an empty model and
        // cannot double-apply the commit.
        if (textareaRef.current) textareaRef.current.value = '';
        compositionPendingRef.current = false;
        compositionCommitRef.current = '';
        compositionHeldRef.current = false;
        if (heldComposition) inputHandlersRef.current.insertText(text, inputTime, true);
        else insertText(text);
        for (const resolve of compositionWaitersRef.current) resolve();
        compositionWaitersRef.current.clear();
        onPendingInputChangeRef.current?.(
          heldInputRef.current.entries.length > 0 || (inputOperationQueueRef.current?.hasPending() ?? false)
        );
        for (const notify of heldInputWaitersRef.current) notify();
      });
    },
    [inputScope, insertText, isCurrentInput, session]
  );

  const handleInput = useCallback(
    (event: React.FormEvent<HTMLTextAreaElement>) => {
      if (discardedCompositionRef.current) {
        event.currentTarget.value = '';
        return;
      }
      if (composingRef.current || compositionPendingRef.current) return;
      // Mobile/browser fallback for input types whose beforeinput carried no
      // data. The textarea is otherwise always empty outside composition.
      const value = event.currentTarget.value;
      if (value) {
        event.currentTarget.value = '';
        insertText(value);
      }
    },
    [insertText]
  );

  const handleCopy = useCallback(
    (event: React.ClipboardEvent<HTMLTextAreaElement>) => {
      const text = primedCopyRef.current ?? (session ? yrsSelectionPlainText(session) : '');
      primedCopyRef.current = null;
      if (!text) return;
      event.preventDefault();
      event.clipboardData.setData('text/plain', text);
    },
    [session]
  );

  const handlePaste = useCallback(
    (event: React.ClipboardEvent<HTMLTextAreaElement>) => {
      event.preventDefault();
      insertText(event.clipboardData.getData('text/plain'));
    },
    [insertText]
  );

  const handleCut = useCallback(
    (event: React.ClipboardEvent<HTMLTextAreaElement>) => {
      if (readOnly || (!holdInput && replicaReadyRef?.current !== false)) return;
      event.preventDefault();
      const copied = copyAfterReplica(true);
      if (copied) deleteSelection(true, copied);
    },
    [copyAfterReplica, deleteSelection, holdInput, readOnly, replicaReadyRef]
  );

  const hasHeldInput = useCallback(
    () => heldInputRef.current.entries.length > 0 ||
      (compositionHeldRef.current && (composingRef.current || compositionPendingRef.current)),
    []
  );
  const flushPendingInputRef = useRef<(() => Promise<void>) | null>(null);
  const flushPendingInput = useCallback(async (): Promise<void> => {
    if (session && !session.isDisplayOnly()) requestWorkerOpenReplicaReadiness(session);
    if (hasHeldInput()) {
      const scope = heldInputRef.current;
      let notify!: () => void;
      let flushing = false;
      try {
        await new Promise<void>((resolve, reject) => {
          notify = () => {
            const lifetime = inputLifetimeRef.current;
            if (
              !lifetime.mounted || !lifetime.enabled || !lifetime.session ||
              heldInputRef.current !== scope ||
              (inputScope === undefined && lifetime.session !== session)
            ) {
              reject(new Error('The document changed while flushing input'));
              return;
            }
            if (!lifetime.session.isDisplayOnly()) requestWorkerOpenReplicaReadiness(lifetime.session);
            if (flushing || holdInputRef.current || hasHeldInput()) return;
            flushing = true;
            flushPendingInputRef.current!().then(resolve, reject);
          };
          heldInputWaitersRef.current.add(notify);
          notify();
        });
      } finally {
        heldInputWaitersRef.current.delete(notify);
      }
      return;
    }
    const queue = inputOperationQueueRef.current;
    const since = queue?.failureCheckpoint();
    const assertCurrent = () => {
      if (!session || !isCurrentInput(session, queue)) {
        throw new Error('The editor input changed or is unavailable while flushing');
      }
    };
    assertCurrent();
    if (composingRef.current || compositionPendingRef.current) {
      await new Promise<void>((resolve) => compositionWaitersRef.current.add(resolve));
      assertCurrent();
    }
    sealInputBatches();
    await queue?.flush(since);
    assertCurrent();
  }, [hasHeldInput, inputScope, isCurrentInput, sealInputBatches, session]);
  flushPendingInputRef.current = flushPendingInput;

  const runAfterPendingInput = useCallback(
    <T,>(operation: () => T | Promise<T>): Promise<T> => {
      const queue = inputOperationQueueRef.current;
      const since = queue?.failureCheckpoint();
      const lifetime = inputLifetimeRef.current;
      const admitted = session;
      const admit = (): Promise<T> => {
        if (!queue) return Promise.reject(new DocxCommandAdmissionError('editor-unavailable'));
        sealInputBatches();
        const assertCurrent = () => {
          if (!lifetime.mounted || !lifetime.enabled || !admitted) {
            throw new DocxCommandAdmissionError('editor-unavailable');
          }
          if (lifetime.session !== admitted || inputOperationQueueRef.current !== queue) {
            throw new DocxCommandAdmissionError('document-replaced');
          }
        };
        return queue.run(async (inputLost) => {
          assertCurrent();
          if (inputLost) throw new DocxCommandAdmissionError('input-failed');
          const result = await operation();
          assertCurrent();
          return result;
        }, since);
      };
      if (!composingRef.current && !compositionPendingRef.current) return admit();
      return new Promise<T>((resolve, reject) => {
        compositionWaitersRef.current.add(() => {
          admit().then(resolve, reject);
        });
      });
    },
    [sealInputBatches, session]
  );

  const hasPendingInput = useCallback(
    () =>
      heldInputRef.current.entries.length > 0 ||
      pendingResidentTextRef.current !== null ||
      composingRef.current ||
      compositionPendingRef.current ||
      (inputOperationQueueRef.current?.hasPending() ?? false),
    []
  );

  useEffect(() => {
    inputLifetimeRef.current.mounted = true;
    return () => {
      inputLifetimeRef.current.mounted = false;
      for (const entry of heldInputRef.current.entries) if (entry.kind === 'copy') entry.onDropped();
      heldInputRef.current.entries = [];
      pendingSelectionsRef.current = [];
      for (const notify of heldInputWaitersRef.current) notify();
      heldInputWaitersRef.current.clear();
      readerScrollRef.current?.stop();
      readerScrollRef.current = null;
      for (const resolve of compositionWaitersRef.current) resolve();
      compositionWaitersRef.current.clear();
    };
  }, []);

  const compositionContextRef = useRef({ inputScope, holdInput, enabled });
  useEffect(() => {
    const previous = compositionContextRef.current;
    compositionContextRef.current = { inputScope, holdInput, enabled };
    sealInputBatches();
    if (
      inputScope !== undefined && inputScope === previous.inputScope && enabled && previous.enabled &&
      (holdInput || previous.holdInput || compositionHeldRef.current)
    ) return;
    if (inputScope !== previous.inputScope && (previous.holdInput || compositionHeldRef.current) && textareaRef.current) {
      textareaRef.current.value = '';
    }
    composingRef.current = false;
    compositionPendingRef.current = false;
    compositionCommitRef.current = '';
    compositionHeldRef.current = false;
    for (const resolve of compositionWaitersRef.current) resolve();
    compositionWaitersRef.current.clear();
  }, [enabled, holdInput, inputScope, sealInputBatches, session]);

  useImperativeHandle(
    ref,
    () => ({
      focus: () => textareaRef.current?.focus({ preventScroll: true }),
      blur: () => textareaRef.current?.blur(),
      isFocused: () => document.activeElement === textareaRef.current,
      flushPendingInput,
      runAfterPendingInput,
      hasPendingInput,
      hasHeldInput,
      queueSelection,
      captureSelectionFromDisplay(anchor, head, targetStory, kind) {
        let anchorLoc = displayPositionToLoc(anchor, targetStory);
        let headLoc = displayPositionToLoc(head, targetStory);
        if (!session || !anchorLoc || !headLoc) return () => {};
        if (kind === 'word' || kind === 'paragraph') {
          const paraId = anchorLoc.paraId;
          const paragraph = session.paragraphs(anchorLoc.story).find((candidate) => candidate.paraId === paraId);
          if (!paragraph) return () => {};
          const [start, end] = kind === 'word'
            ? findWordBoundaries(paragraph.text, anchorLoc.offset)
            : [0, paragraph.text.length];
          headLoc = { ...anchorLoc, offset: end };
          anchorLoc = { ...anchorLoc, offset: start };
        }
        const stickyAnchor = session.encodeStickyPosition(anchorLoc);
        const stickyHead = session.encodeStickyPosition(headLoc);
        return () => {
          if (!isCurrentInput(session)) return;
          const currentAnchor = session.resolveStickyPosition(stickyAnchor);
          const currentHead = session.resolveStickyPosition(stickyHead);
          if (!currentAnchor || !currentHead) return;
          advanceInteractionEpoch();
          verticalCaretGoalRef.current.reset();
          setSelection(currentAnchor, currentHead);
        };
      },
      setSelectionFromDisplay(anchor, head = anchor, targetStory = story) {
        const anchorLoc = displayPositionToLoc(anchor, targetStory);
        const headLoc = displayPositionToLoc(head, targetStory);
        if (session && anchorLoc && headLoc) {
          advanceInteractionEpoch();
          verticalCaretGoalRef.current.reset();
          setSelection(anchorLoc, headLoc);
        }
      },
      selectWordAtDisplay(position, targetStory = story) {
        if (!session) return;
        const loc = displayPositionToLoc(position, targetStory);
        if (!loc) return;
        const paragraph = session
          .paragraphs(loc.story)
          .find((candidate) => candidate.paraId === loc.paraId);
        if (!paragraph) return;
        const [start, end] = findWordBoundaries(paragraph.text, loc.offset);
        if (start < end) {
          advanceInteractionEpoch();
          verticalCaretGoalRef.current.reset();
          setSelection({ ...loc, offset: start }, { ...loc, offset: end });
        }
      },
      selectParagraphAtDisplay(position, targetStory = story) {
        if (!session) return;
        const loc = displayPositionToLoc(position, targetStory);
        if (!loc) return;
        const paragraph = session
          .paragraphs(loc.story)
          .find((candidate) => candidate.paraId === loc.paraId);
        if (!paragraph) return;
        advanceInteractionEpoch();
        verticalCaretGoalRef.current.reset();
        setSelection({ ...loc, offset: 0 }, { ...loc, offset: paragraph.text.length });
      },
      displaySelection,
      keepSelectionInPlace() {
        quietSelectionRef.current = session?.selection() ?? null;
      },
      applyStoredFormatting,
      clearStoredFormatting() {
        const current = ensureSelection();
        if (current) {
          storedFormattingByParagraphRef.current.delete(
            `${current.head.story}\u0000${current.head.paraId}`
          );
        }
      },
      storedFormatting,
      insertText,
      deleteSelection,
      selectAll() {
        advanceInteractionEpoch();
        selectAll();
      },
    }),
    [
      advanceInteractionEpoch,
      applyStoredFormatting,
      displayPositionToLoc,
      displaySelection,
      ensureSelection,
      flushPendingInput,
      hasPendingInput,
      hasHeldInput,
      insertText,
      isCurrentInput,
      queueSelection,
      deleteSelection,
      runAfterPendingInput,
      selectAll,
      session,
      setSelection,
      storedFormatting,
      story,
    ]
  );

  useEffect(() => {
    if (!enabled) return;
    storedFormattingByParagraphRef.current.clear();
  }, [enabled, session]);

  useEffect(() => {
    verticalCaretGoalRef.current.reset();
    pendingResidentFrameEpochRef.current = null;
    pendingLocalCaretRevealRef.current = null;
  }, [session, story]);

  useEffect(() => {
    if (!enabled || !session || !replicaReady || holdInput) return;
    if (seedSelection) ensureSelection();
    emitSelection(false);
  }, [emitSelection, enabled, ensureSelection, holdInput, seedSelection, session, replicaReady]);

  const focusedSessionRef = useRef<YrsSession | null>(null);
  useEffect(() => {
    if (!enabled || !session || readOnly) {
      focusedSessionRef.current = null;
      return;
    }
    const storyOnly = focusedSessionRef.current === session;
    focusedSessionRef.current = session;
    const frame = requestAnimationFrame(() => {
      const textarea = textareaRef.current;
      if (!textarea) return;
      if (storyOnly) {
        const active = textarea.ownerDocument.activeElement;
        const root = textarea.closest('.paged-editor') ?? textarea;
        if (active && active !== textarea.ownerDocument.body && !root.contains(active)) return;
      }
      textarea.focus({ preventScroll: true });
    });
    return () => cancelAnimationFrame(frame);
  }, [enabled, readOnly, session, story]);

  useEffect(() => {
    if (!enabled || !displayListQueries) return;
    if (!displayListQueries.isReady()) return;
    const pendingFrameEpoch = pendingResidentFrameEpochRef.current;
    if (pendingFrameEpoch !== null) {
      if (displayListFrameEpoch === null || displayListFrameEpoch < pendingFrameEpoch) return;
      pendingResidentFrameEpochRef.current = null;
    }
    const selection = displaySelection();
    if (!selection) return;
    onStateChange(selection, false);

    // Same-frame worker caret geometry needs no facade query (and so cannot
    // force handle adoption on the typing path).
    const authoritativeCaret =
      residentCaretAuthoritative &&
      residentCaret?.caretRect &&
      residentCaret.frameEpoch === displayListFrameEpoch &&
      residentCaret.selection &&
      sameYrsSelection(residentCaret.selection, session?.selection() ?? null)
        ? residentCaret.caretRect
        : null;
    const caret = authoritativeCaret ?? displayListQueries.caretRect(selection.head);
    const host = canvasHostRef?.current;
    if (!caret || !host) return;
    const pageRect = resolveDisplayPageClientRect(host, displayListQueries, caret.pageIndex);
    const pageSize = displayListQueries.pageSize(caret.pageIndex);
    if (!pageRect || !pageSize) return;
    const scaleX = pageSize.width > 0 ? pageRect.width / pageSize.width : 1;
    const scaleY = pageSize.height > 0 ? pageRect.height / pageSize.height : 1;
    const nextLeft = pageRect.left + caret.x * scaleX;
    const nextTop = pageRect.top + caret.y * scaleY;
    const nextHeight = Math.max(1, caret.height * scaleY);
    const inputZoom = textareaRef.current ? effectiveZoom(textareaRef.current) : 1;
    const style = {
      left: nextLeft / inputZoom,
      top: nextTop / inputZoom,
      height: nextHeight / inputZoom,
    };
    setPositionStyle((current) =>
      current.left === style.left && current.top === style.top && current.height === style.height
        ? current
        : style
    );
    const stickySelection = session?.selection() ?? null;
    const previousStickySelection = lastCaretScrollSelectionRef.current;
    lastCaretScrollSelectionRef.current = stickySelection;
    const quiet = quietSelectionRef.current;
    quietSelectionRef.current = null;
    const selectionIsQuiet = quiet && stickySelection && sameYrsSelection(quiet, stickySelection);
    const selectionChanged =
      (previousStickySelection === undefined ||
        !sameYrsSelection(previousStickySelection, stickySelection)) &&
      !selectionIsQuiet;
    const scroller = findVerticalScrollParentOrRoot(host);
    const pendingReveal = pendingLocalCaretRevealRef.current;
    const caretRevealOrigin =
      pendingReveal &&
      pendingReveal.scroller === scroller &&
      pendingReveal.scrollTop === scroller.scrollTop &&
      !selectionIsQuiet
        ? 'local'
        : layoutUpdateOrigin;
    pendingLocalCaretRevealRef.current = null;
    if (
      selection.anchor === selection.head &&
      shouldScrollCaretIntoView(caretRevealOrigin, selectionChanged, readOnly)
    ) {
      const delta = scrollIntoViewDelta(scrollViewport(scroller), nextTop, nextTop + nextHeight, 24);
      if (delta !== 0) scroller.scrollTop += delta;
    }
  }, [
    canvasHostRef,
    displayListFrameEpoch,
    displayListQueries,
    displaySelection,
    enabled,
    layoutUpdateOrigin,
    onStateChange,
    readOnly,
    residentCaret,
    residentCaretAuthoritative,
    selectionEpoch,
    session,
  ]);

  if (!enabled) return null;
  const textarea = (
    <textarea
      ref={textareaRef}
      // Preserve the editor focus-target contract while the implementation is
      // yrs-owned. This is only a compatibility class on the hidden textarea;
      // no ProseMirror view or dependency is mounted here.
      className="paged-editor__yrs-input paged-editor__hidden-pm ProseMirror"
      data-testid="yrs-input"
      data-yrs-story={story}
      aria-label="Document input"
      autoCapitalize="sentences"
      autoCorrect="on"
      spellCheck
      readOnly={readOnly || !session}
      rows={1}
      style={{ ...BASE_STYLE, ...positionStyle }}
      onBeforeInput={handleBeforeInput}
      onInput={handleInput}
      onKeyDown={handleKeyDown}
      onCompositionStart={handleCompositionStart}
      onCompositionUpdate={handleCompositionUpdate}
      onCompositionEnd={handleCompositionEnd}
      onCopy={handleCopy}
      onPaste={handlePaste}
      onCut={handleCut}
      onFocus={(event) => {
        event.currentTarget.classList.add('ProseMirror-focused');
        onFocusChange?.(true);
      }}
      onBlur={(event) => {
        event.currentTarget.classList.remove('ProseMirror-focused');
        onFocusChange?.(false);
      }}
    />
  );
  return typeof document !== 'undefined' && document.body
    ? createPortal(textarea, document.body)
    : textarea;
});

export const YrsInput = memo(YrsInputComponent);

export default YrsInput;
