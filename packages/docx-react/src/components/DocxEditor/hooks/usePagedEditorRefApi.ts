import { useEffect, useImperativeHandle, useRef } from 'react';

import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { Document } from '@betteroffice/docx/types/document';
import type { ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import type {
  YrsInputPositionMap,
  YrsLoc,
  YrsSession,
  YrsStickyPosition,
} from '@betteroffice/docx/yrs';

import type { YrsInputRef } from '../YrsInput';
import type { PagedEditorRef } from '../PagedEditor';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { performYrsHistoryAction, type YrsEditorCommand } from '../yrsCommands';
import {
  currentYrsToolbarSelection,
  withStoredYrsFormatting,
  type FormattingAction,
  type YrsToolbarSelection,
} from '../yrsToolbar';
import { DocxCommandAdmissionError } from '../../../commands/createDocxCommandStore';
import type { RevealPositionOutcome } from './usePagedScrollApi';

/** The image under a one-unit selection. */
export interface PagedEditorSelectedImage {
  pos: number;
  attrs: Readonly<Record<string, unknown>>;
}

/** Sticky positions before and after one image, which follow it through later edits. */
export type PagedEditorImageHandle = readonly [YrsStickyPosition, YrsStickyPosition];

/**
 * Ordered command entry points for editor chrome. Immediate methods act on
 * the current selection without moving focus; callers admit them through
 * {@link PagedEditorCommandBridge.runAfterPendingInput} first.
 */
export interface PagedEditorCommandBridge {
  runAfterPendingInput<T>(operation: () => T | Promise<T>): Promise<T>;
  hasPendingInput(): boolean;
  /** Notified after each published selection, document or pending-input change. */
  subscribe(listener: () => void): () => void;
  session(): YrsSession | null;
  rootStory(): string;
  hasSelection(): boolean;
  /** Selection read model with stored caret formatting; `live` recomputes it. */
  toolbarSelection(live: boolean): YrsToolbarSelection | null;
  selectedImage(): PagedEditorSelectedImage | null;
  /** A handle on the image at display position `pos`. */
  imageHandle(pos: number): PagedEditorImageHandle | null;
  /** Current display position of the image `handle` holds, or null once it is gone. */
  imagePosition(handle: PagedEditorImageHandle): number | null;
  /** Applies formatting; throws when the engine refuses it. */
  format(action: FormattingAction): boolean;
  /** Applies a structural command; throws when the engine refuses it. */
  command(command: YrsEditorCommand): boolean;
  history(redo: boolean): boolean;
  /** Selects a story range, publishes it and scrolls it into view. */
  select(start: YrsLoc, end: YrsLoc): boolean;
}

interface RefApiInputs {
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  layout: Layout | null;
  runLayoutPipeline: () => void;
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
  revealPositionImpl: (position: number) => RevealPositionOutcome;
  scrollToParaIdImpl: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  scrollToPageImpl: (pageNumber: number) => void;
  setIsFocused: React.Dispatch<React.SetStateAction<boolean>>;
  documentFromYrsRef: React.MutableRefObject<() => Document | null>;
  yrsSessionRef: React.MutableRefObject<YrsSession | null>;
  yrsLocToDisplayPositionRef: React.MutableRefObject<(loc: YrsLoc) => number | null>;
  syncYrsInputStateRef: React.MutableRefObject<
    (docChanged: boolean, dirtyStory?: string | readonly string[]) => boolean
  >;
  applyYrsFormattingRef: React.MutableRefObject<(action: FormattingAction) => boolean>;
  applyYrsCommandRef: React.MutableRefObject<(command: YrsEditorCommand) => boolean>;
  getYrsPositionProjectionRef: React.MutableRefObject<() => YrsPositionProjection | null>;
  displayPositionToYrsLocRef: React.MutableRefObject<(position: number) => YrsLoc | null>;
}

function storyOffsetToLoc(session: YrsSession, story: string, offset: number): YrsLoc | null {
  const paragraphs = session.paragraphs(story);
  if (paragraphs.length === 0) return null;
  for (const paragraph of paragraphs) {
    const span = session.locateParagraph(story, paragraph.paraId);
    if (offset <= span.end) {
      return {
        story,
        paraId: paragraph.paraId,
        offset: Math.min(Math.max(0, offset - span.start), span.end - span.start),
      };
    }
  }
  const last = paragraphs[paragraphs.length - 1];
  const span = session.locateParagraph(story, last.paraId);
  return { story, paraId: last.paraId, offset: span.end - span.start };
}

function buildRefApi(inputs: RefApiInputs): PagedEditorRef {
  const {
    yrsInputRef,
    layout,
    runLayoutPipeline,
    scrollToPositionImpl,
    revealPositionImpl,
    scrollToParaIdImpl,
    scrollToPageImpl,
    setIsFocused,
    documentFromYrsRef,
    yrsSessionRef,
    yrsLocToDisplayPositionRef,
    syncYrsInputStateRef,
    applyYrsFormattingRef,
    applyYrsCommandRef,
    getYrsPositionProjectionRef,
    displayPositionToYrsLocRef,
  } = inputs;

  const setDisplaySelection = (anchor: number, head = anchor): void => {
    const session = yrsSessionRef.current;
    const projection = getYrsPositionProjectionRef.current();
    if (!session || !projection) return;
    const anchorTarget = projection.targetAt(anchor);
    const headTarget = projection.targetAt(head);
    if (anchorTarget.story !== headTarget.story) return;
    yrsInputRef.current?.setSelectionFromDisplay(
      anchorTarget.displayPosition,
      headTarget.displayPosition,
      anchorTarget.story
    );
  };

  const selectLocRange = (start: YrsLoc, end: YrsLoc): boolean => {
    const session = yrsSessionRef.current;
    if (!session || start.story !== end.story) return false;
    session.setSelection(start, end);
    const startPos = yrsLocToDisplayPositionRef.current(start);
    if (startPos != null) scrollToPositionImpl(startPos, true);
    yrsInputRef.current?.focus();
    return true;
  };

  return {
    getDocument: () => documentFromYrsRef.current(),
    focus: () => {
      yrsInputRef.current?.focus();
      setIsFocused(true);
    },
    blur: () => {
      yrsInputRef.current?.blur();
      setIsFocused(false);
    },
    isFocused: () => yrsInputRef.current?.isFocused() ?? false,
    undo: () => {
      const session = yrsSessionRef.current;
      const result = session
        ? performYrsHistoryAction(session, false)
        : { changed: false, stories: [] };
      if (result.changed) syncYrsInputStateRef.current(true, result.stories);
      return result.changed;
    },
    redo: () => {
      const session = yrsSessionRef.current;
      const result = session
        ? performYrsHistoryAction(session, true)
        : { changed: false, stories: [] };
      if (result.changed) syncYrsInputStateRef.current(true, result.stories);
      return result.changed;
    },
    canUndo: () => yrsSessionRef.current?.canUndo() ?? false,
    canRedo: () => yrsSessionRef.current?.canRedo() ?? false,
    setSelection: setDisplaySelection,
    insertText: (text) => yrsInputRef.current?.insertText(text),
    deleteSelection: () => yrsInputRef.current?.deleteSelection(),
    selectAll: () => yrsInputRef.current?.selectAll(),
    getSelectionRange: () => {
      const selection = yrsInputRef.current?.displaySelection();
      return selection
        ? {
            from: Math.min(selection.anchor, selection.head),
            to: Math.max(selection.anchor, selection.head),
          }
        : null;
    },
    displayPositionToYrsLoc: (position) => displayPositionToYrsLocRef.current(position),
    getYrsSession: () => yrsSessionRef.current,
    flushPendingInput: async () => {
      const input = yrsInputRef.current;
      const session = yrsSessionRef.current;
      if (!input || !session) throw new Error('The editor input is unavailable');
      // The input rejects its own flush when it unmounts or changes session; its handle object
      // is rebuilt whenever a new frame changes its callbacks, so only the session is compared.
      await input.flushPendingInput();
      if (session !== yrsSessionRef.current) {
        throw new Error('The document changed while flushing input');
      }
    },
    getYrsStoredFormatting: () => yrsInputRef.current?.storedFormatting() ?? null,
    yrsLocToDisplayPosition: (loc) => yrsLocToDisplayPositionRef.current(loc),
    syncYrsInputState: (docChanged, dirtyStories) =>
      syncYrsInputStateRef.current(docChanged, dirtyStories),
    applyYrsFormatting: (action) => applyYrsFormattingRef.current(action),
    applyYrsCommand: (command) => applyYrsCommandRef.current(command),
    getLayout: () => layout,
    relayout: runLayoutPipeline,
    scrollToPosition: scrollToPositionImpl,
    revealDisplayPosition: revealPositionImpl,
    scrollToParaId: scrollToParaIdImpl,
    scrollToPage: scrollToPageImpl,
    highlightRange: (from, to) => {
      const projection = getYrsPositionProjectionRef.current();
      if (!projection || !Number.isFinite(from) || !Number.isFinite(to) || from < 0 || from > to)
        return;
      const end = Math.min(to, projection.size);
      if (from > projection.size) return;
      setDisplaySelection(from, end);
      scrollToPositionImpl(from, true);
    },
    scrollToCommentId: (commentId) => {
      const session = yrsSessionRef.current;
      if (!session) return false;
      try {
        const anchor = session.resolveComment(String(commentId))[0];
        if (!anchor) return false;
        const start = storyOffsetToLoc(session, anchor.story, anchor.start);
        const end = storyOffsetToLoc(session, anchor.story, anchor.end);
        return !!start && !!end && selectLocRange(start, end);
      } catch {
        return false;
      }
    },
    scrollToChangeId: (revisionId) => {
      const revision = yrsSessionRef.current
        ?.listRevisions()
        .find((candidate) => candidate.revisionId === String(revisionId));
      return revision
        ? selectLocRange(
            { story: revision.story, ...revision.range.start },
            { story: revision.story, ...revision.range.end }
          )
        : false;
    },
  };
}

export interface UsePagedEditorRefApiOptions {
  ref: React.Ref<PagedEditorRef>;
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  layout: Layout | null;
  runLayoutPipeline: () => void;
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
  revealPositionImpl: (position: number) => RevealPositionOutcome;
  scrollToParaIdImpl: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  scrollToPageImpl: (pageNumber: number) => void;
  setIsFocused: React.Dispatch<React.SetStateAction<boolean>>;
  onReadyRef: React.MutableRefObject<((ref: PagedEditorRef) => void) | undefined>;
  documentFromYrs: () => Document | null;
  yrsSession: YrsSession | null;
  yrsLocToDisplayPosition: (loc: YrsLoc) => number | null;
  syncYrsInputState: (docChanged: boolean, dirtyStory?: string | readonly string[]) => boolean;
  applyYrsFormatting: (action: FormattingAction) => boolean;
  applyYrsCommand: (command: YrsEditorCommand) => boolean;
  getYrsPositionProjection: () => YrsPositionProjection | null;
  displayPositionToYrsLoc: (position: number) => YrsLoc | null;
}

export function usePagedEditorRefApi(opts: UsePagedEditorRefApiOptions): void {
  const {
    ref,
    yrsInputRef,
    layout,
    runLayoutPipeline,
    scrollToPositionImpl,
    revealPositionImpl,
    scrollToParaIdImpl,
    scrollToPageImpl,
    setIsFocused,
    onReadyRef,
    documentFromYrs,
    yrsSession,
    yrsLocToDisplayPosition,
    syncYrsInputState,
    applyYrsFormatting,
    applyYrsCommand,
    getYrsPositionProjection,
    displayPositionToYrsLoc,
  } = opts;
  const documentFromYrsRef = useRef(documentFromYrs);
  const yrsSessionRef = useRef(yrsSession);
  const yrsLocToDisplayPositionRef = useRef(yrsLocToDisplayPosition);
  const syncYrsInputStateRef = useRef(syncYrsInputState);
  const applyYrsFormattingRef = useRef(applyYrsFormatting);
  const applyYrsCommandRef = useRef(applyYrsCommand);
  const getYrsPositionProjectionRef = useRef(getYrsPositionProjection);
  const displayPositionToYrsLocRef = useRef(displayPositionToYrsLoc);
  documentFromYrsRef.current = documentFromYrs;
  yrsSessionRef.current = yrsSession;
  yrsLocToDisplayPositionRef.current = yrsLocToDisplayPosition;
  syncYrsInputStateRef.current = syncYrsInputState;
  applyYrsFormattingRef.current = applyYrsFormatting;
  applyYrsCommandRef.current = applyYrsCommand;
  getYrsPositionProjectionRef.current = getYrsPositionProjection;
  displayPositionToYrsLocRef.current = displayPositionToYrsLoc;

  const inputs = {
    yrsInputRef,
    layout,
    runLayoutPipeline,
    scrollToPositionImpl,
    revealPositionImpl,
    scrollToParaIdImpl,
    scrollToPageImpl,
    setIsFocused,
    documentFromYrsRef,
    yrsSessionRef,
    yrsLocToDisplayPositionRef,
    syncYrsInputStateRef,
    applyYrsFormattingRef,
    applyYrsCommandRef,
    getYrsPositionProjectionRef,
    displayPositionToYrsLocRef,
  };

  useImperativeHandle(ref, () => buildRefApi(inputs), [
    layout,
    runLayoutPipeline,
    scrollToPositionImpl,
    revealPositionImpl,
    scrollToParaIdImpl,
    scrollToPageImpl,
  ]);

  useEffect(() => {
    if (onReadyRef.current && yrsSession) onReadyRef.current(buildRefApi(inputs));
  }, [layout, runLayoutPipeline, scrollToParaIdImpl, scrollToPageImpl, yrsSession]);
}

export interface UsePagedEditorCommandBridgeOptions {
  bridgeRef: React.MutableRefObject<PagedEditorCommandBridge | null> | undefined;
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  session: YrsSession | null;
  rootStory: string;
  inputPositionMap: (story: string) => YrsInputPositionMap | null;
  latestSelectionRef: React.RefObject<YrsToolbarSelection | null>;
  listenersRef: React.RefObject<Set<() => void>>;
  getPositionProjection: () => YrsPositionProjection | null;
  displayPositionToLoc: (position: number) => YrsLoc | null;
  format: (action: FormattingAction) => boolean;
  command: (command: YrsEditorCommand) => boolean;
  syncYrsInputState: (docChanged: boolean, dirtyStory?: string | readonly string[]) => boolean;
  yrsLocToDisplayPosition: (loc: YrsLoc) => number | null;
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
}

/** Publishes {@link PagedEditorCommandBridge} for the owning editor's chrome. */
export function usePagedEditorCommandBridge(options: UsePagedEditorCommandBridgeOptions): void {
  const latest = useRef(options);
  latest.current = options;
  const bridge = useRef<PagedEditorCommandBridge | null>(null);
  if (!bridge.current) {
    bridge.current = {
      runAfterPendingInput(operation) {
        const input = latest.current.yrsInputRef.current;
        if (!input) return Promise.reject(new DocxCommandAdmissionError('editor-unavailable'));
        return input.runAfterPendingInput(operation);
      },
      hasPendingInput: () => latest.current.yrsInputRef.current?.hasPendingInput() ?? false,
      subscribe(listener) {
        const listeners = latest.current.listenersRef.current;
        listeners.add(listener);
        return () => {
          listeners.delete(listener);
        };
      },
      session: () => latest.current.session,
      rootStory: () => latest.current.rootStory,
      hasSelection: () => latest.current.session?.selection() != null,
      toolbarSelection(live) {
        const current = latest.current;
        if (!live) return current.latestSelectionRef.current;
        const session = current.session;
        if (!session) return null;
        const story = session.selection()?.head.story ?? current.rootStory;
        const map = current.inputPositionMap(story);
        const selection = map ? currentYrsToolbarSelection(session, map) : null;
        return selection
          ? withStoredYrsFormatting(
              selection,
              current.yrsInputRef.current?.storedFormatting() ?? null
            )
          : null;
      },
      selectedImage() {
        const current = latest.current;
        const selection = current.yrsInputRef.current?.displaySelection();
        if (!selection || Math.abs(selection.anchor - selection.head) !== 1) return null;
        const pos = Math.min(selection.anchor, selection.head);
        const node = current.getPositionProjection()?.nodeAt(pos);
        return node?.kind === 'image' ? { pos, attrs: node.attrs } : null;
      },
      imageHandle(pos) {
        const current = latest.current;
        const session = current.session;
        if (!session || current.getPositionProjection()?.nodeAt(pos)?.kind !== 'image') return null;
        const at = current.displayPositionToLoc(pos);
        return at
          ? [
              session.encodeStickyPosition(at),
              session.encodeStickyPosition({ ...at, offset: at.offset + 1 }),
            ]
          : null;
      },
      imagePosition([before, after]) {
        const current = latest.current;
        const start = current.session?.resolveStickyPosition(before);
        const end = current.session?.resolveStickyPosition(after);
        if (!start || !end || start.paraId !== end.paraId || end.offset !== start.offset + 1) {
          return null;
        }
        const pos = current.yrsLocToDisplayPosition(start);
        const node = pos == null ? null : current.getPositionProjection()?.nodeAt(pos);
        return node?.kind === 'image' ? node.start : null;
      },
      format: (action) => latest.current.format(action),
      command: (command) => latest.current.command(command),
      history(redo) {
        const current = latest.current;
        const session = current.session;
        if (!session) return false;
        const result = performYrsHistoryAction(session, redo);
        if (result.changed) current.syncYrsInputState(true, result.stories);
        return result.changed;
      },
      select(start, end) {
        const current = latest.current;
        const session = current.session;
        if (!session || start.story !== end.story) return false;
        session.setSelection(start, end);
        current.syncYrsInputState(false);
        const position = current.yrsLocToDisplayPosition(start);
        if (position != null) current.scrollToPositionImpl(position, true);
        return true;
      },
    };
  }
  const { bridgeRef } = options;
  useEffect(() => {
    if (!bridgeRef) return;
    bridgeRef.current = bridge.current;
    return () => {
      if (bridgeRef.current === bridge.current) bridgeRef.current = null;
    };
  }, [bridgeRef]);
}
