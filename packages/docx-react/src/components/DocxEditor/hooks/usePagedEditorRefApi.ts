import { useEffect, useImperativeHandle, useRef } from 'react';

import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { Document } from '@betteroffice/docx/types/document';
import type { ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import type { YrsLoc, YrsSession } from '@betteroffice/docx/yrs';

import type { YrsInputRef } from '../YrsInput';
import type { PagedEditorRef } from '../PagedEditor';
import type { FormattingAction } from '../../Toolbar';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { performYrsHistoryAction, type YrsEditorCommand } from '../yrsCommands';

interface RefApiInputs {
  yrsInputRef: React.RefObject<YrsInputRef | null>;
  layout: Layout | null;
  runLayoutPipeline: () => void;
  scrollToPositionImpl: (pmPos: number, forParaIdScroll?: boolean) => void;
  scrollToParaIdImpl: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  scrollToPageImpl: (pageNumber: number) => void;
  setIsFocused: React.Dispatch<React.SetStateAction<boolean>>;
  documentFromYrsRef: React.MutableRefObject<() => Document | null>;
  yrsSessionRef: React.MutableRefObject<YrsSession | null>;
  yrsLocToDisplayPositionRef: React.MutableRefObject<(loc: YrsLoc) => number | null>;
  syncYrsInputStateRef: React.MutableRefObject<(docChanged: boolean) => boolean>;
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
      const changed = session ? performYrsHistoryAction(session, false) : false;
      if (changed) syncYrsInputStateRef.current(true);
      return changed;
    },
    redo: () => {
      const session = yrsSessionRef.current;
      const changed = session ? performYrsHistoryAction(session, true) : false;
      if (changed) syncYrsInputStateRef.current(true);
      return changed;
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
    getYrsStoredFormatting: () => yrsInputRef.current?.storedFormatting() ?? null,
    yrsLocToDisplayPosition: (loc) => yrsLocToDisplayPositionRef.current(loc),
    syncYrsInputState: (docChanged) => syncYrsInputStateRef.current(docChanged),
    applyYrsFormatting: (action) => applyYrsFormattingRef.current(action),
    applyYrsCommand: (command) => applyYrsCommandRef.current(command),
    getLayout: () => layout,
    relayout: runLayoutPipeline,
    scrollToPosition: scrollToPositionImpl,
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
  scrollToParaIdImpl: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  scrollToPageImpl: (pageNumber: number) => void;
  setIsFocused: React.Dispatch<React.SetStateAction<boolean>>;
  onReadyRef: React.MutableRefObject<((ref: PagedEditorRef) => void) | undefined>;
  documentFromYrs: () => Document | null;
  yrsSession: YrsSession | null;
  yrsLocToDisplayPosition: (loc: YrsLoc) => number | null;
  syncYrsInputState: (docChanged: boolean) => boolean;
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
    scrollToParaIdImpl,
    scrollToPageImpl,
  ]);

  useEffect(() => {
    if (onReadyRef.current && yrsSession) onReadyRef.current(buildRefApi(inputs));
  }, [layout, runLayoutPipeline, scrollToParaIdImpl, scrollToPageImpl, yrsSession]);
}
