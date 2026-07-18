import { useCallback, useRef } from 'react';
import type { YrsLoc, YrsSession, YrsStoryRange } from '@betteroffice/docx/yrs';
import {
  findAllMatches,
  type FindMatch,
  type FindOptions,
  type FindResult,
} from '../../dialogs/findReplaceUtils';
import type { useFindReplace } from '../../../hooks/useFindReplace';
import type { PagedEditorRef } from '../PagedEditor';

export type YrsFindMatch = FindMatch & {
  displayFrom: number;
  displayTo: number;
  yrsRange: YrsStoryRange;
};

type YrsFindResult = FindResult & {
  matches: YrsFindMatch[];
};

function findMatchesInYrs(
  session: YrsSession,
  locToDisplayPosition: (loc: YrsLoc) => number | null,
  searchText: string,
  options: FindOptions
): YrsFindMatch[] {
  const matches: YrsFindMatch[] = [];
  const paragraphs = session.paragraphs('body');
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    const paragraph = paragraphs[paragraphIndex];
    if (!paragraph.text) continue;
    for (const match of findAllMatches(paragraph.text, searchText, options)) {
      const startLoc = { story: 'body', paraId: paragraph.paraId, offset: match.start };
      const endLoc = { story: 'body', paraId: paragraph.paraId, offset: match.end };
      const displayFrom = locToDisplayPosition(startLoc);
      const displayTo = locToDisplayPosition(endLoc);
      if (displayFrom == null || displayTo == null || displayFrom >= displayTo) continue;
      matches.push({
        paragraphIndex,
        contentIndex: 0,
        startOffset: match.start,
        endOffset: match.end,
        text: paragraph.text.slice(match.start, match.end),
        displayFrom,
        displayTo,
        yrsRange: {
          story: 'body',
          start: { paraId: paragraph.paraId, offset: match.start },
          end: { paraId: paragraph.paraId, offset: match.end },
        },
      });
    }
  }
  return matches;
}

/** Yrs-backed find, navigation, and replacement for the canvas editor. */
export function useFindReplaceBridge({
  pagedEditorRef,
  findReplace,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  findReplace: ReturnType<typeof useFindReplace>;
}) {
  const findResultRef = useRef<FindResult | null>(null);

  const goToMatch = useCallback(
    (match: YrsFindMatch | undefined, index: number): FindMatch | null => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (!editor || !session || !match) return null;
      try {
        session.setSelection(
          { story: match.yrsRange.story, ...match.yrsRange.start },
          { story: match.yrsRange.story, ...match.yrsRange.end }
        );
        if (!editor.syncYrsInputState(false)) return null;
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

  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (!editor || !session || !searchText.trim()) {
        findResultRef.current = null;
        findReplace.setMatches([], 0);
        return null;
      }
      const matches = findMatchesInYrs(
        session,
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
    [findReplace, goToMatch, pagedEditorRef]
  );

  const handleFindNext = useCallback((): FindMatch | null => {
    const result = findResultRef.current as YrsFindResult | null;
    if (!result?.matches.length) return null;
    const index = (result.currentIndex + 1) % result.matches.length;
    return goToMatch(result.matches[index], index);
  }, [goToMatch]);

  const handleFindPrevious = useCallback((): FindMatch | null => {
    const result = findResultRef.current as YrsFindResult | null;
    if (!result?.matches.length) return null;
    const index = result.currentIndex === 0 ? result.matches.length - 1 : result.currentIndex - 1;
    return goToMatch(result.matches[index], index);
  }, [goToMatch]);

  const handleReplace = useCallback(
    (replaceText: string): boolean => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      const result = findResultRef.current as YrsFindResult | null;
      const match = result?.matches[result.currentIndex];
      if (!editor || !session || !match) return false;
      try {
        session.replaceRange(match.yrsRange, replaceText);
        session.setSelection({
          story: match.yrsRange.story,
          paraId: match.yrsRange.start.paraId,
          offset: match.yrsRange.start.offset + replaceText.length,
        });
        return editor.syncYrsInputState(true);
      } catch (error) {
        console.error('Replace failed:', error);
        return false;
      }
    },
    [pagedEditorRef]
  );

  const handleReplaceAll = useCallback(
    (searchText: string, replaceText: string, options: FindOptions): number => {
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      if (!editor || !session || !searchText.trim()) return 0;
      const matches = findMatchesInYrs(
        session,
        (loc) => editor.yrsLocToDisplayPosition(loc),
        searchText,
        options
      );
      if (matches.length === 0) return 0;
      for (const match of [...matches].sort((a, b) => b.displayFrom - a.displayFrom)) {
        session.replaceRange(match.yrsRange, replaceText);
      }
      const first = matches[0];
      session.setSelection({
        story: first.yrsRange.story,
        paraId: first.yrsRange.start.paraId,
        offset: first.yrsRange.start.offset + replaceText.length,
      });
      editor.syncYrsInputState(true);
      findResultRef.current = null;
      findReplace.setMatches([], 0);
      return matches.length;
    },
    [findReplace, pagedEditorRef]
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
