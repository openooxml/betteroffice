/** Text and IME input surface backed by sticky session positions. */

import React, {
  forwardRef,
  memo,
  useCallback,
  useEffect,
  useImperativeHandle,
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
import { awaitWorkerOpenReplica } from './internals/workerOpenReplica';
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
  replicaReadyRef?: React.RefObject<boolean>;
  /** Asks for the replica when input reaches the textarea before it has loaded. */
  requestReplica?: () => void;
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
    residentCaretReady?: boolean
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
    replicaReadyRef,
    requestReplica,
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
  useEffect(() => {
    const textarea = textareaRef.current;
    if (!textarea || !requestReplica) return;
    const onInput = () => {
      if (replicaReadyRef?.current === false) requestReplica();
    };
    textarea.addEventListener('focus', onInput);
    textarea.addEventListener('keydown', onInput, true);
    return () => {
      textarea.removeEventListener('focus', onInput);
      textarea.removeEventListener('keydown', onInput, true);
    };
  }, [enabled, replicaReadyRef, requestReplica]);
  const composingRef = useRef(false);
  const compositionPendingRef = useRef(false);
  const compositionCommitRef = useRef('');
  const compositionWaitersRef = useRef(new Set<() => void>());
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
      onDropped?: () => void
    ): void => {
      sealInputBatches();
      const admitted = session;
      const queue = inputOperationQueueRef.current;
      const replica =
        admitted && replicaReadyRef?.current === false
          ? awaitWorkerOpenReplica(admitted)
          : undefined;
      if (!replica) {
        queue?.enqueue(() => {
          if (!isCurrentInput(admitted, queue)) return onDropped?.();
          if (replicaReadyRef?.current !== false) replicaInputRef.current.applyPendingSelection?.();
          return operation(false);
        });
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
        if (!isCurrentInput(admitted, queue) || (kind === 'selection' && superseded)) {
          onDropped?.();
          return;
        }
        replicaInputRef.current.applyPendingSelection?.();
        await operation(true);
        if (scrolled && isCurrentInput(admitted, queue)) {
          quietSelectionRef.current = admitted?.selection() ?? null;
        }
      });
    },
    [isCurrentInput, replicaReadyRef, sealInputBatches, session, watchReaderScroll]
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
    const current = seedSelection ? ensureSelection() : readSelection();
    if (!current) return null;
    const anchor = locToDisplayPosition(current.anchor);
    const head = locToDisplayPosition(current.head);
    return anchor == null || head == null ? null : { anchor, head };
  }, [ensureSelection, locToDisplayPosition, readSelection, seedSelection]);

  const emitSelection = useCallback(
    (docChanged: boolean, residentLayoutReady = false, residentCaretReady = false): void => {
      const selection = displaySelection();
      if (!selection) return;
      setSelectionEpoch((epoch) => epoch + 1);
      onStateChange(selection, docChanged, residentLayoutReady, residentCaretReady);
    },
    [displaySelection, onStateChange]
  );

  const setSelection = useCallback(
    (anchor: YrsLoc, head: YrsLoc = anchor, emit = true): void => {
      if (!session) return;
      session.setSelection(anchor, head);
      if (emit) {
        onCaretInterrupt?.();
        emitSelection(false);
      }
    },
    [emitSelection, onCaretInterrupt, session]
  );

  const finishMutation = useCallback(
    (
      residentLayoutReady = false,
      residentCaretReady = false,
      dirtyStories?: string | readonly string[]
    ): void => {
      verticalCaretGoalRef.current.reset();
      if (!composingRef.current && textareaRef.current) textareaRef.current.value = '';
      onCaretInput?.();
      onDirectInput(dirtyStories);
      emitSelection(true, residentLayoutReady, residentCaretReady);
    },
    [emitSelection, onCaretInput, onDirectInput]
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
    (text: string): void => {
      verticalCaretGoalRef.current.reset();
      if (!session || readOnly || text.length === 0) return;
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

      const canBatchResidentText =
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
      inputPositionMap,
      isCurrentInput,
      isSuggesting,
      readOnly,
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
        const paragraphs = session.paragraphs(activeStory);
        const index = paragraphs.findIndex((paragraph) => paragraph.paraId === caret.paraId);
        if (index < 0) return;
        const paragraph = paragraphs[index];
        const hasTarget =
          direction === 'backward'
            ? caret.offset > 0 || index > 0
            : caret.offset < map.paragraphs[index].length || index + 1 < paragraphs.length;
        if (hasTarget && isBodyFlowStory(activeStory) && !isSuggesting && applyResidentDelete) {
          const applied = await applyResidentDelete(direction, remaining);
          if (!isCurrentInput(session)) return;
          if (applied) {
            finishResidentMutation(applied);
            remaining -= Math.max(1, applied.deletedUnits ?? remaining);
            continue;
          }
        }
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
      verticalCaretGoalRef.current.reset();
      dispatchCaretInput();
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
    [deleteUnits, dispatchCaretInput, enqueueInputOperation]
  );

  const splitParagraph = useCallback((): void => {
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
    nextParagraphStyleId,
    readOnly,
    replicaReadyRef,
    session,
    suggestingAuthor,
  ]);

  const moveSelection = useCallback(
    (
      direction: 'left' | 'right' | 'up' | 'down' | 'home' | 'end',
      extend: boolean,
      wholeDocument: boolean,
      byWord = false
    ): void => {
      const verticalDirection = direction === 'up' || direction === 'down' ? direction : null;
      let interactionEpoch = verticalDirection
        ? inputOperationQueueRef.current?.captureInteractionEpoch()
        : undefined;
      enqueueInputOperation(async (waited) => {
        // The gesture applied before input that waited is older than it.
        if (waited && verticalDirection) {
          interactionEpoch = inputOperationQueueRef.current?.captureInteractionEpoch();
        }
        if (!verticalDirection) verticalCaretGoalRef.current.reset();
        if (!session) return;
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
      inputPositionMap,
      isCurrentInput,
      locToDisplayPosition,
      resolveDisplayTarget,
      session,
      setSelection,
    ]
  );

  const selectAll = useCallback((): void => {
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
  }, [enqueueInputOperation, ensureSelection, inputPositionMap, readOnly, session, setSelection, story]);

  const moveTableCell = useCallback(
    (backward: boolean): boolean => {
      verticalCaretGoalRef.current.reset();
      if (!session) return false;
      const current = ensureSelection();
      const focused = current ? yrsCellLocFromStory(current.head.story) : null;
      if (!focused) return false;
      const tableRange = yrsTableSelectionRange(session, focused, 'table');
      if (!tableRange) return false;

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
        if (backward) return true;
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

  const handleBeforeInput = useCallback(
    (event: React.FormEvent<HTMLTextAreaElement>): void => {
      const native = event.nativeEvent as InputEvent;
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

  // A copy asked before an on-demand replica loads writes the selection it then has.
  const copyAfterReplica = useCallback((): boolean => {
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
    if (!session || !requestReplica || replicaReadyRef?.current !== false || !clipboard) {
      return false;
    }
    const copied = session;
    const text = new Promise<Blob>((resolve, reject) => {
      enqueueInputOperation(
        () => {
          const selected = yrsSelectionPlainText(copied);
          if (selected) resolve(new Blob([selected], { type: 'text/plain' }));
          else reject(new Error('Nothing is selected to copy'));
        },
        'selection',
        () => reject(new Error('Newer input replaced the copy'))
      );
    });
    void text.catch(() => {});
    if (typeof ClipboardItem === 'function' && clipboard.write) {
      void clipboard.write([new ClipboardItem({ 'text/plain': text })]).catch(() => {});
    } else {
      void text
        .then((blob) => blob.text())
        .then((value) => clipboard.writeText(value))
        .catch(() => {});
    }
    return true;
  }, [enqueueInputOperation, replicaReadyRef, requestReplica, session]);

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
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
      } else if (event.key === 'Tab' && !readOnly && moveTableCell(event.shiftKey)) {
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
      moveSelection,
      moveTableCell,
      primeCopy,
      readOnly,
      selectAll,
      splitParagraph,
    ]
  );

  const handleCompositionStart = useCallback(
    (event: React.CompositionEvent<HTMLTextAreaElement>) => {
      verticalCaretGoalRef.current.reset();
      composingRef.current = true;
      compositionPendingRef.current = false;
      compositionCommitRef.current = '';
      event.currentTarget.value = '';
      onCaretInterrupt?.();
      onPendingInputChangeRef.current?.(true);
    },
    [onCaretInterrupt]
  );

  const handleCompositionUpdate = useCallback(
    (event: React.CompositionEvent<HTMLTextAreaElement>) => {
      compositionCommitRef.current = event.data;
    },
    []
  );

  const handleCompositionEnd = useCallback(
    (event: React.CompositionEvent<HTMLTextAreaElement>) => {
      composingRef.current = false;
      compositionPendingRef.current = true;
      compositionCommitRef.current =
        event.currentTarget.value || event.data || compositionCommitRef.current;
      queueMicrotask(() => {
        if (!compositionPendingRef.current || !isCurrentInput(session)) return;
        const text = textareaRef.current?.value || compositionCommitRef.current;
        // Reset the browser model before applying the document op. A trailing
        // post-composition beforeinput therefore observes an empty model and
        // cannot double-apply the commit.
        if (textareaRef.current) textareaRef.current.value = '';
        compositionPendingRef.current = false;
        compositionCommitRef.current = '';
        insertText(text);
        for (const resolve of compositionWaitersRef.current) resolve();
        compositionWaitersRef.current.clear();
        onPendingInputChangeRef.current?.(inputOperationQueueRef.current?.hasPending() ?? false);
      });
    },
    [insertText, isCurrentInput, session]
  );

  const handleInput = useCallback(
    (event: React.FormEvent<HTMLTextAreaElement>) => {
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

  const flushPendingInput = useCallback(async (): Promise<void> => {
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
  }, [isCurrentInput, sealInputBatches, session]);

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
      readerScrollRef.current?.stop();
      readerScrollRef.current = null;
      for (const resolve of compositionWaitersRef.current) resolve();
      compositionWaitersRef.current.clear();
    };
  }, []);

  useEffect(() => {
    composingRef.current = false;
    compositionPendingRef.current = false;
    compositionCommitRef.current = '';
    sealInputBatches();
    for (const resolve of compositionWaitersRef.current) resolve();
    compositionWaitersRef.current.clear();
  }, [sealInputBatches, session, enabled]);

  useImperativeHandle(
    ref,
    () => ({
      focus: () => textareaRef.current?.focus({ preventScroll: true }),
      blur: () => textareaRef.current?.blur(),
      isFocused: () => document.activeElement === textareaRef.current,
      flushPendingInput,
      runAfterPendingInput,
      hasPendingInput,
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
      deleteSelection() {
        if (!readOnly && replicaReadyRef?.current !== false && deleteSelected()) {
          advanceInteractionEpoch();
          finishMutation();
        }
      },
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
      finishMutation,
      flushPendingInput,
      hasPendingInput,
      insertText,
      deleteSelected,
      readOnly,
      replicaReadyRef,
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
  }, [session, story]);

  useEffect(() => {
    if (!enabled || !session || !replicaReady) return;
    if (seedSelection) ensureSelection();
    emitSelection(false);
  }, [emitSelection, enabled, ensureSelection, seedSelection, session, replicaReady]);

  useEffect(() => {
    if (!enabled || !session || readOnly) return;
    const frame = requestAnimationFrame(() =>
      textareaRef.current?.focus({ preventScroll: true })
    );
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
    const selectionChanged =
      (previousStickySelection === undefined ||
        !sameYrsSelection(previousStickySelection, stickySelection)) &&
      !(quiet && stickySelection && sameYrsSelection(quiet, stickySelection));
    if (
      selection.anchor === selection.head &&
      shouldScrollCaretIntoView(layoutUpdateOrigin, selectionChanged, readOnly)
    ) {
      const scroller = findVerticalScrollParentOrRoot(host);
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
