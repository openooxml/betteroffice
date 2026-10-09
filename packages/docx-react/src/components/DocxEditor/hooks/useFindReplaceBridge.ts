import { useCallback, useEffect, useRef } from 'react';
import { findBodyMatches, type YrsSelection, type YrsStoryRange } from '@betteroffice/docx/yrs';
import { DocxCommandAdmissionError } from '../../../commands/createDocxCommandStore';
import type { DocxCommandResult } from '../../../commands/types';
import type { FindMatch, FindOptions, FindResult } from '../../dialogs/findReplaceUtils';
import type { useFindReplace } from '../../../hooks/useFindReplace';
import type { PagedEditorRef } from '../PagedEditor';
import { commandOutcome } from './useDocxCommands';

export type YrsFindMatch = FindMatch & {
  displayFrom: number;
  displayTo: number;
  yrsRange: YrsStoryRange;
};

type YrsFindResult = FindResult & {
  matches: YrsFindMatch[];
};

function selects(selection: YrsSelection | null, range: YrsStoryRange): boolean {
  if (!selection || selection.anchor.story !== range.story || selection.head.story !== range.story) {
    return false;
  }
  const at = (loc: YrsSelection['anchor'], point: YrsStoryRange['start']) =>
    loc.paraId === point.paraId && loc.offset === point.offset;
  return (
    (at(selection.anchor, range.start) && at(selection.head, range.end)) ||
    (at(selection.anchor, range.end) && at(selection.head, range.start))
  );
}

/**
 * Yrs-backed find, navigation, and replacement for the canvas editor.
 * Replacements run through `complete`, after accepted input, and find their
 * target again in the document as it is then.
 */
export function useFindReplaceBridge({
  pagedEditorRef,
  findReplace,
  complete,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  findReplace: ReturnType<typeof useFindReplace>;
  complete: (write: () => DocxCommandResult) => Promise<DocxCommandResult>;
}) {
  const findResultRef = useRef<FindResult | null>(null);
  const searchRef = useRef<{ text: string; options: FindOptions } | null>(null);
  const generationRef = useRef(0);
  const viewerVersionRef = useRef<string | null>(null);
  const isOpen = findReplace.state.isOpen;

  useEffect(() => () => {
    generationRef.current += 1;
  }, [isOpen]);

  const goToMatch = useCallback(
    (match: YrsFindMatch | undefined, index: number): FindMatch | null => {
      const editor = pagedEditorRef.current;
      if (!editor || !match) return null;
      try {
        if (editor.isWorkerViewer?.()) {
          editor.setSelection(match.displayFrom, match.displayTo);
        } else {
          const session = editor.getYrsSession();
          if (!session) return null;
          session.setSelection(
            { story: match.yrsRange.story, ...match.yrsRange.start },
            { story: match.yrsRange.story, ...match.yrsRange.end }
          );
          if (!editor.syncYrsInputState(false)) return null;
        }
        editor.scrollToPosition(match.displayFrom);
      } catch (error) {
        console.error('Find navigation failed:', error);
        return null;
      }
      const result = findResultRef.current as YrsFindResult | null;
      if (result) findResultRef.current = { ...result, currentIndex: index };
      findReplace.goToMatch(index);
      return match;
    },
    [findReplace, pagedEditorRef]
  );

  const searchViewer = useCallback(
    (editor: PagedEditorRef, searchText: string, options: FindOptions, pick: (count: number) => number) => {
      const generation = ++generationRef.current;
      const session = editor.getYrsSession();
      void editor.readViewerFindMatches(searchText, options).then((read) => {
        const current = pagedEditorRef.current;
        if (read === null || generationRef.current !== generation ||
          !current?.isWorkerViewer() || current.getYrsSession() !== session) return;
        const { matches } = read;
        const index = matches.length > 0 ? pick(matches.length) : 0;
        const result: YrsFindResult = { matches, totalCount: matches.length, currentIndex: index };
        viewerVersionRef.current = read.version;
        findResultRef.current = result;
        findReplace.setMatches(matches, index);
        if (matches.length > 0) goToMatch(matches[index], index);
      }).catch(() => {});
    },
    [findReplace, goToMatch, pagedEditorRef]
  );

  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      const viewer = editor?.isWorkerViewer?.() === true;
      generationRef.current += 1;
      if (!editor || (!viewer && !session) || !searchText.trim()) {
        findResultRef.current = null;
        searchRef.current = null;
        findReplace.setMatches([], 0);
        return null;
      }
      searchRef.current = { text: searchText, options };
      if (viewer) {
        findResultRef.current = null;
        viewerVersionRef.current = null;
        findReplace.setMatches([], 0);
        searchViewer(editor, searchText, options, () => 0);
        return null;
      }
      const matches = findBodyMatches(
        session!,
        (loc) => editor.yrsLocToDisplayPosition(loc),
        searchText,
        options
      );
      const result: YrsFindResult = { matches, totalCount: matches.length, currentIndex: 0 };
      findResultRef.current = result;
      findReplace.setMatches(matches, 0);
      if (matches.length > 0) goToMatch(matches[0], 0);
      return result;
    },
    [findReplace, goToMatch, pagedEditorRef, searchViewer]
  );

  const step = useCallback((next: (current: number, count: number) => number): FindMatch | null => {
    const result = findResultRef.current as YrsFindResult | null;
    const editor = pagedEditorRef.current;
    const search = searchRef.current;
    if (result && search && editor?.isWorkerViewer?.() &&
      editor.getYrsSession()?.version() !== viewerVersionRef.current) {
      searchViewer(editor, search.text, search.options, (count) =>
        result.matches.length > 0 ? next(Math.min(result.currentIndex, count - 1), count) : 0
      );
      return null;
    }
    if (!result?.matches.length) return null;
    const index = next(result.currentIndex, result.matches.length);
    return goToMatch(result.matches[index], index);
  }, [goToMatch, pagedEditorRef, searchViewer]);

  const handleFindNext = useCallback(
    (): FindMatch | null => step((current, count) => (current + 1) % count),
    [step]
  );

  const handleFindPrevious = useCallback(
    (): FindMatch | null => step((current, count) => (current === 0 ? count - 1 : current - 1)),
    [step]
  );

  const handleReplace = useCallback(
    async (replaceText: string): Promise<boolean> => {
      if (pagedEditorRef.current?.isWorkerViewer?.()) return false;
      const result = await complete(() => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const search = searchRef.current;
        if (!editor || !session || !search) return commandOutcome(false);
        const selection = session.selection();
        const match = findBodyMatches(
          session,
          (loc) => editor.yrsLocToDisplayPosition(loc),
          search.text,
          search.options
        ).find((candidate) => selects(selection, candidate.yrsRange));
        if (!match) throw new DocxCommandAdmissionError('target-changed');
        const landed = session.replaceRange(match.yrsRange, replaceText).range;
        session.setSelection(
          landed
            ? { story: landed.story, ...landed.end }
            : {
                story: match.yrsRange.story,
                paraId: match.yrsRange.start.paraId,
                offset: match.yrsRange.start.offset + replaceText.length,
              }
        );
        return commandOutcome(editor.syncYrsInputState(true));
      });
      return result.ok && result.status === 'executed';
    },
    [complete, pagedEditorRef]
  );

  const handleReplaceAll = useCallback(
    async (searchText: string, replaceText: string, options: FindOptions): Promise<number> => {
      if (pagedEditorRef.current?.isWorkerViewer?.()) return 0;
      let replaced = 0;
      const result = await complete(() => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        if (!editor || !session || !searchText.trim()) return commandOutcome(false);
        const matches = findBodyMatches(
          session,
          (loc) => editor.yrsLocToDisplayPosition(loc),
          searchText,
          options
        );
        if (matches.length === 0) return commandOutcome(false);
        let landed: YrsStoryRange | undefined;
        for (const match of [...matches].sort((a, b) => b.displayFrom - a.displayFrom)) {
          landed = session.replaceRange(match.yrsRange, replaceText).range;
        }
        const first = matches[0];
        session.setSelection(
          landed
            ? { story: landed.story, ...landed.end }
            : {
                story: first.yrsRange.story,
                paraId: first.yrsRange.start.paraId,
                offset: first.yrsRange.start.offset + replaceText.length,
              }
        );
        editor.syncYrsInputState(true);
        findResultRef.current = null;
        findReplace.setMatches([], 0);
        replaced = matches.length;
        return commandOutcome(true);
      });
      return result.ok ? replaced : 0;
    },
    [complete, findReplace, pagedEditorRef]
  );

  return {
    findResultRef,
    handleFind,
    handleFindNext,
    handleFindPrevious,
    handleReplace,
    handleReplaceAll,
  };
}
