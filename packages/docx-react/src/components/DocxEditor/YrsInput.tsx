/**
 * Monaco-style textarea input surface for the experimental yrs core.
 *
 * The textarea owns browser text/IME events but never owns document text.
 * Outside composition it is reset after every event; during composition it
 * temporarily holds the browser's composing string and commits exactly once
 * on compositionend. The document selection itself is the session's pair of
 * Rust-backed sticky positions.
 */

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
  type YrsRunMark,
  type YrsSelection,
  type YrsSession,
  type YrsStoryRange,
} from '@betteroffice/docx/yrs';
import {
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { ResidentFrameApplyResult } from './hooks/useDisplayList';
import type { ResolveDisplayListQueries } from './hooks/displayListQueryEpochGate';
import { findWordBoundaries } from '@betteroffice/docx/utils';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import {
  performYrsHistoryAction,
  yrsCellLocFromStory,
  yrsCellStory,
  yrsSelectionNearTable,
  yrsTableSelectionRange,
} from './yrsCommands';
import { InputOperationQueue } from './inputOperationQueue';
import { paragraphVerticalMove, VerticalCaretGoal } from './verticalCaretGoal';
import {
  shouldScrollCaretIntoView,
  type LayoutUpdateOrigin,
} from './internals/viewportAnchoring';

export interface YrsDisplaySelection {
  anchor: number;
  head: number;
}

export interface YrsInputRef {
  focus(): void;
  blur(): void;
  isFocused(): boolean;
  setSelectionFromDisplay(anchor: number, head?: number, story?: string): void;
  selectWordAtDisplay(position: number, story?: string): void;
  selectParagraphAtDisplay(position: number, story?: string): void;
  displaySelection(): YrsDisplaySelection | null;
  applyStoredFormatting(action: YrsStoredFormattingAction): void;
  clearStoredFormatting(): void;
  storedFormatting(): YrsStoredFormatting | null;
  insertText(text: string): void;
  deleteSelection(): void;
  selectAll(): void;
}

export type YrsStoredFormattingAction =
  | {
      type: 'toggle';
      mark: 'bold' | 'italic' | 'underline' | 'strike';
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
  onDirectInput(): void;
  /** One-owner body text path; false until the resident frame is initialized. */
  applyResidentInput?(text: string): Promise<ResidentFrameApplyResult | null>;
  /** One-owner collapsed delete/merge path; false until the resident frame is initialized. */
  applyResidentDelete?(
    direction: 'backward' | 'forward'
  ): Promise<ResidentFrameApplyResult | null>;
  onFocusChange?(focused: boolean): void;
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

const YrsInputComponent = forwardRef<YrsInputRef, YrsInputProps>(function YrsInput(
  {
    enabled,
    readOnly,
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
    onCaretInput,
    onCaretInputDispatched,
    onCaretInterrupt,
  },
  ref
) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const composingRef = useRef(false);
  const compositionPendingRef = useRef(false);
  const compositionCommitRef = useRef('');
  const storedFormattingByParagraphRef = useRef(new Map<string, YrsStoredFormatting>());
  const inputOperationQueueRef = useRef<InputOperationQueue | null>(null);
  if (!inputOperationQueueRef.current) {
    inputOperationQueueRef.current = new InputOperationQueue((error) => {
      console.error('[YrsInput] queued input operation failed', error);
    });
  }
  const pendingResidentTextRef = useRef<{ text: string } | null>(null);
  const pendingResidentFrameEpochRef = useRef<number | null>(null);
  const verticalCaretGoalRef = useRef(new VerticalCaretGoal());
  const displayListQueriesRef = useRef(displayListQueries);
  const displayListFrameEpochRef = useRef(displayListFrameEpoch);
  const resolveDisplayListQueriesRef = useRef(resolveDisplayListQueries);
  const lastCaretScrollSelectionEpochRef = useRef<number | null>(null);
  displayListQueriesRef.current = displayListQueries;
  displayListFrameEpochRef.current = displayListFrameEpoch;
  resolveDisplayListQueriesRef.current = resolveDisplayListQueries;
  const [positionStyle, setPositionStyle] = useState<CSSProperties>({ left: 0, top: 0, height: 1 });
  const [selectionEpoch, setSelectionEpoch] = useState(0);

  const enqueueInputOperation = useCallback((operation: () => void | Promise<void>): void => {
    inputOperationQueueRef.current?.enqueue(operation);
  }, []);

  const suggestingAuthor = useCallback((): YrsAuthor | undefined => {
    return isSuggesting ? { name: author, date: new Date().toISOString() } : undefined;
  }, [author, isSuggesting]);

  const belongsToRootStory = useCallback(
    (candidate: string): boolean => candidate === story || candidate.startsWith(`${story}:`),
    [story]
  );

  const ensureSelection = useCallback((): YrsSelection | null => {
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
    const first = inputPositionMap(story)?.paragraphs[0];
    if (!first) return null;
    const loc = { story, paraId: first.paraId, offset: 0 };
    session.setSelection(loc);
    return { anchor: loc, head: loc };
  }, [belongsToRootStory, inputPositionMap, session, story]);

  const displaySelection = useCallback((): YrsDisplaySelection | null => {
    const current = ensureSelection();
    if (!current) return null;
    const anchor = locToDisplayPosition(current.anchor);
    const head = locToDisplayPosition(current.head);
    return anchor == null || head == null ? null : { anchor, head };
  }, [ensureSelection, locToDisplayPosition]);

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
    (residentLayoutReady = false, residentCaretReady = false): void => {
      verticalCaretGoalRef.current.reset();
      if (!composingRef.current && textareaRef.current) textareaRef.current.value = '';
      onCaretInput?.();
      onDirectInput();
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
    if (!session || readOnly) return;
    if (session.selection()?.head.story !== 'body') return;
    onCaretInputDispatched?.();
  }, [onCaretInputDispatched, readOnly, session]);

  const storedFormatting = useCallback((): YrsStoredFormatting | null => {
    const current = ensureSelection();
    if (!current) return null;
    return (
      storedFormattingByParagraphRef.current.get(
        `${current.head.story}\u0000${current.head.paraId}`
      ) ?? null
    );
  }, [ensureSelection]);

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
    session.deleteRange(range, suggestingAuthor());
    const collapsed = { story: range.story, ...range.start };
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
          for (let i = 0; i < pieces.length; i += 1) {
            const piece = pieces[i];
            if (piece || (i === 0 && hasSelection)) {
              const insertedAt = caret;
              if (i === 0 && hasSelection) {
                session.replaceRange(selectedRange, piece, suggestingAuthor());
              } else {
                session.insertText(caret, piece, suggestingAuthor());
              }
              caret = { ...caret, offset: caret.offset + piece.length };
              const insertedStored = storedFormattingByParagraphRef.current.get(
                `${insertedAt.story}\u0000${insertedAt.paraId}`
              );
              if (insertedStored) {
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
          session.setSelection(caret);
          finishMutation();
        };
        if (
          !hasSelection &&
          current.head.story === 'body' &&
          !isSuggesting &&
          !stored &&
          /^[\x20-\x7e]+$/u.test(inputText) &&
          !inputText.includes('\r') &&
          !inputText.includes('\n') &&
          applyResidentInput
        ) {
          const applied = await applyResidentInput(inputText);
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
        // Seal any earlier text batch so a synchronous structural operation
        // remains an ordering barrier for later input.
        pendingResidentTextRef.current = null;
        enqueueInputOperation(() => applyText(text));
        return;
      }

      const pending = pendingResidentTextRef.current;
      if (pending) {
        pending.text += text;
        return;
      }

      const batch = { text };
      pendingResidentTextRef.current = batch;
      enqueueInputOperation(async () => {
        if (pendingResidentTextRef.current === batch) pendingResidentTextRef.current = null;
        await applyText(batch.text);
      });
    },
    [
      applyResidentInput,
      dispatchCaretInput,
      enqueueInputOperation,
      ensureSelection,
      finishMutation,
      finishResidentMutation,
      inputPositionMap,
      isSuggesting,
      readOnly,
      session,
      suggestingAuthor,
    ]
  );

  const deleteDirection = useCallback(
    (direction: 'backward' | 'forward'): void => {
      verticalCaretGoalRef.current.reset();
      dispatchCaretInput();
      enqueueInputOperation(async () => {
        if (!session || readOnly) return;
        if (deleteSelected()) {
          finishMutation();
          return;
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
        if (hasTarget && activeStory === 'body' && !isSuggesting && applyResidentDelete) {
          const applied = await applyResidentDelete(direction);
          if (applied) {
            finishResidentMutation(applied);
            return;
          }
        }
        if (direction === 'backward') {
          if (caret.offset > 0) {
            const start = previousCodePointOffset(paragraph.text, caret.offset);
            session.deleteRange(
              {
                story: activeStory,
                start: { paraId: caret.paraId, offset: start },
                end: { paraId: caret.paraId, offset: caret.offset },
              },
              suggestingAuthor()
            );
            session.setSelection({ ...caret, offset: start });
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
            session.deleteRange(
              {
                story: activeStory,
                start: { paraId: caret.paraId, offset: caret.offset },
                end: { paraId: caret.paraId, offset: end },
              },
              suggestingAuthor()
            );
            session.setSelection(caret);
          } else if (index + 1 < paragraphs.length) {
            session.mergeParagraphs(activeStory, caret.paraId, suggestingAuthor());
            session.setSelection(caret);
          } else {
            return;
          }
        }
        finishMutation();
      });
    },
    [
      applyResidentDelete,
      deleteSelected,
      dispatchCaretInput,
      enqueueInputOperation,
      ensureSelection,
      finishMutation,
      finishResidentMutation,
      inputPositionMap,
      isSuggesting,
      readOnly,
      session,
      suggestingAuthor,
    ]
  );

  const splitParagraph = useCallback((): void => {
    verticalCaretGoalRef.current.reset();
    dispatchCaretInput();
    enqueueInputOperation(() => {
      if (!session || readOnly) return;
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
    });
  }, [
    deleteSelected,
    dispatchCaretInput,
    enqueueInputOperation,
    ensureSelection,
    finishMutation,
    nextParagraphStyleId,
    readOnly,
    session,
    suggestingAuthor,
  ]);

  const toggleMark = useCallback(
    (mark: YrsRunMark): void => {
      enqueueInputOperation(() => {
        if (!session || readOnly) return;
        const current = ensureSelection();
        const map = current ? inputPositionMap(current.anchor.story) : null;
        if (!current || !map) return;
        const range = toRange(current, map);
        if (range.start.paraId === range.end.paraId && range.start.offset === range.end.offset) {
          const context = session.selectionContext(range);
          const active =
            mark.type === 'bold'
              ? context.bold === true
              : mark.type === 'italic'
                ? context.italic === true
                : mark.type === 'underline'
                  ? context.underline === true
                  : false;
          if (mark.type === 'bold' || mark.type === 'italic' || mark.type === 'underline') {
            applyStoredFormatting({ type: 'toggle', mark: mark.type, active });
          }
          return;
        }
        storedFormattingByParagraphRef.current.delete(
          `${current.head.story}\u0000${current.head.paraId}`
        );
        session.toggleMark(range, mark);
        finishMutation();
      });
    },
    [
      applyStoredFormatting,
      enqueueInputOperation,
      ensureSelection,
      finishMutation,
      inputPositionMap,
      readOnly,
      session,
    ]
  );

  const undoRedo = useCallback(
    (redo: boolean): void => {
      enqueueInputOperation(() => {
        if (!session || readOnly) return;
        const changed = performYrsHistoryAction(session, redo);
        if (changed) finishMutation();
      });
    },
    [enqueueInputOperation, finishMutation, readOnly, session]
  );

  const setAlignment = useCallback(
    (alignment: 'left' | 'center' | 'right' | 'both'): void => {
      enqueueInputOperation(() => {
        if (!session || readOnly) return;
        const current = ensureSelection();
        const map = current ? inputPositionMap(current.anchor.story) : null;
        if (!current || !map) return;
        session.setParagraphAttrs(toRange(current, map), { alignment }, suggestingAuthor());
        finishMutation();
      });
    },
    [
      enqueueInputOperation,
      ensureSelection,
      finishMutation,
      inputPositionMap,
      readOnly,
      session,
      suggestingAuthor,
    ]
  );

  const moveSelection = useCallback(
    (
      direction: 'left' | 'right' | 'up' | 'down' | 'home' | 'end',
      extend: boolean,
      wholeDocument: boolean,
      byWord = false
    ): void => {
      const verticalDirection = direction === 'up' || direction === 'down' ? direction : null;
      enqueueInputOperation(async () => {
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
      const activeStory = current?.head.story;
      const map = activeStory ? inputPositionMap(activeStory) : null;
      if (!session || !activeStory || !map || map.paragraphs.length === 0) return;
      const first = map.paragraphs[0];
      const last = map.paragraphs[map.paragraphs.length - 1];
      setSelection(
        { story: activeStory, paraId: first.paraId, offset: 0 },
        { story: activeStory, paraId: last.paraId, offset: last.length }
      );
    });
  }, [enqueueInputOperation, ensureSelection, inputPositionMap, session, setSelection]);

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
        // Match ProseMirror tables' terminal-Tab behavior when the document
        // has no trailing paragraph: append a row and enter its first cell.
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

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
      if (event.nativeEvent.isComposing || composingRef.current) return;
      const mod = event.metaKey || event.ctrlKey;
      const key = event.key.toLowerCase();
      if (mod && key === 'b') {
        event.preventDefault();
        toggleMark({ type: 'bold' });
      } else if (mod && key === 'i') {
        event.preventDefault();
        toggleMark({ type: 'italic' });
      } else if (mod && key === 'u') {
        event.preventDefault();
        toggleMark({ type: 'underline' });
      } else if (mod && key === 'z') {
        event.preventDefault();
        undoRedo(event.shiftKey);
      } else if (mod && key === 'y') {
        event.preventDefault();
        undoRedo(true);
      } else if (mod && key === 'a') {
        event.preventDefault();
        selectAll();
      } else if (mod && (key === 'e' || key === 'l' || key === 'r' || key === 'j')) {
        event.preventDefault();
        setAlignment(
          key === 'e' ? 'center' : key === 'r' ? 'right' : key === 'j' ? 'both' : 'left'
        );
      } else if (event.key === 'Enter') {
        event.preventDefault();
        splitParagraph();
      } else if (event.key === 'Tab' && moveTableCell(event.shiftKey)) {
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
      deleteDirection,
      moveSelection,
      moveTableCell,
      selectAll,
      setAlignment,
      splitParagraph,
      toggleMark,
      undoRedo,
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
        const text = textareaRef.current?.value || compositionCommitRef.current;
        // Reset the browser model before applying the document op. A trailing
        // post-composition beforeinput therefore observes an empty model and
        // cannot double-apply the commit.
        if (textareaRef.current) textareaRef.current.value = '';
        compositionPendingRef.current = false;
        compositionCommitRef.current = '';
        insertText(text);
      });
    },
    [insertText]
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

  const handlePaste = useCallback(
    (event: React.ClipboardEvent<HTMLTextAreaElement>) => {
      event.preventDefault();
      insertText(event.clipboardData.getData('text/plain'));
    },
    [insertText]
  );

  useImperativeHandle(
    ref,
    () => ({
      focus: () => textareaRef.current?.focus({ preventScroll: true }),
      blur: () => textareaRef.current?.blur(),
      isFocused: () => document.activeElement === textareaRef.current,
      setSelectionFromDisplay(anchor, head = anchor, targetStory = story) {
        verticalCaretGoalRef.current.reset();
        const anchorLoc = displayPositionToLoc(anchor, targetStory);
        const headLoc = displayPositionToLoc(head, targetStory);
        if (session && anchorLoc && headLoc) setSelection(anchorLoc, headLoc);
      },
      selectWordAtDisplay(position, targetStory = story) {
        verticalCaretGoalRef.current.reset();
        if (!session) return;
        const loc = displayPositionToLoc(position, targetStory);
        if (!loc) return;
        const paragraph = session
          .paragraphs(loc.story)
          .find((candidate) => candidate.paraId === loc.paraId);
        if (!paragraph) return;
        const [start, end] = findWordBoundaries(paragraph.text, loc.offset);
        if (start < end) {
          setSelection({ ...loc, offset: start }, { ...loc, offset: end });
        }
      },
      selectParagraphAtDisplay(position, targetStory = story) {
        verticalCaretGoalRef.current.reset();
        if (!session) return;
        const loc = displayPositionToLoc(position, targetStory);
        if (!loc) return;
        const paragraph = session
          .paragraphs(loc.story)
          .find((candidate) => candidate.paraId === loc.paraId);
        if (!paragraph) return;
        setSelection({ ...loc, offset: 0 }, { ...loc, offset: paragraph.text.length });
      },
      displaySelection,
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
        if (!readOnly && deleteSelected()) finishMutation();
      },
      selectAll,
    }),
    [
      applyStoredFormatting,
      displayPositionToLoc,
      displaySelection,
      ensureSelection,
      finishMutation,
      insertText,
      deleteSelected,
      readOnly,
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
    if (!enabled || !session) return;
    ensureSelection();
    emitSelection(false);
    if (!readOnly) requestAnimationFrame(() => textareaRef.current?.focus({ preventScroll: true }));
  }, [emitSelection, enabled, ensureSelection, readOnly, session]);

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
    setPositionStyle((current) =>
      current.left === nextLeft && current.top === nextTop && current.height === nextHeight
        ? current
        : { left: nextLeft, top: nextTop, height: nextHeight }
    );
    const selectionChanged = lastCaretScrollSelectionEpochRef.current !== selectionEpoch;
    lastCaretScrollSelectionEpochRef.current = selectionEpoch;
    if (
      selection.anchor === selection.head &&
      shouldScrollCaretIntoView(layoutUpdateOrigin, selectionChanged)
    ) {
      const scroller = findVerticalScrollParentOrRoot(host);
      const viewport = scroller.getBoundingClientRect();
      const margin = 24;
      const caretBottom = nextTop + nextHeight;
      if (nextTop < viewport.top + margin) {
        scroller.scrollTop += nextTop - viewport.top - margin;
      } else if (caretBottom > viewport.bottom - margin) {
        scroller.scrollTop += caretBottom - viewport.bottom + margin;
      }
    }
  }, [
    canvasHostRef,
    displayListFrameEpoch,
    displayListQueries,
    displaySelection,
    enabled,
    layoutUpdateOrigin,
    onStateChange,
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
