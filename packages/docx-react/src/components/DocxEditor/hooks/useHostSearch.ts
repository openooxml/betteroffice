import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';
import type { CanvasFindMatch } from '../overlays/CanvasFindHighlightOverlay';

/** Options for {@link DocxHostSearch.search}. */
export interface DocxSearchOptions {
  /** Defaults to false. */
  caseSensitive?: boolean;
}

/** Where a host search stands. */
export interface DocxSearchState {
  query: string;
  options: Required<DocxSearchOptions>;
  /** Matches in the document body, tables included. */
  total: number;
  /** Zero-based index of the current match, or -1 without matches. */
  current: number;
}

/** Find driven from the host's own UI; see `DocxEditorRef.search`. */
export interface DocxHostSearch {
  search: (query: string, options?: DocxSearchOptions) => Promise<DocxSearchState>;
  searchNext: () => DocxSearchState | null;
  searchPrevious: () => DocxSearchState | null;
  searchGoTo: (index: number) => DocxSearchState | null;
  clearSearch: () => void;
  getSearchState: () => DocxSearchState | null;
  onSearchChange: (listener: (state: DocxSearchState | null) => void) => () => void;
}

interface SearchMatch extends CanvasFindMatch {
  paraId: string;
  start: number;
}

interface SearchRun {
  query: string;
  options: Required<DocxSearchOptions>;
  session: YrsSession;
  version: string;
  matches: SearchMatch[];
  current: number;
}

function isBodyStory(story: string): boolean {
  return story === 'body' || story.startsWith('body:');
}

/** Body matches in display order; ones without a display range are left out. */
function collectMatches(
  editor: PagedEditorRef,
  session: YrsSession,
  query: string,
  options: Required<DocxSearchOptions>
): SearchMatch[] {
  const matches: SearchMatch[] = [];
  for (const hit of session.searchText(query, options)) {
    if (!isBodyStory(hit.story)) continue;
    const loc = { story: hit.story, paraId: hit.paraId };
    const displayFrom = editor.yrsLocToDisplayPosition({ ...loc, offset: hit.start });
    const displayTo = editor.yrsLocToDisplayPosition({ ...loc, offset: hit.end });
    if (displayFrom == null || displayTo == null || displayFrom >= displayTo) continue;
    matches.push({ displayFrom, displayTo, paraId: hit.paraId, start: hit.start });
  }
  // Rust walks block content controls after the body; display order puts them in place.
  return matches.sort((a, b) => a.displayFrom - b.displayFrom);
}

function stateOf(run: SearchRun | null): DocxSearchState | null {
  return run
    ? {
        query: run.query,
        options: { ...run.options },
        total: run.matches.length,
        current: run.current,
      }
    : null;
}

/** The first match on or after the page in view, wrapping to the first. */
function firstInView(
  matches: readonly SearchMatch[],
  queries: DisplayListQueries | null,
  pageIndex: number
): number {
  if (matches.length === 0) return -1;
  if (!queries || pageIndex <= 0) return 0;
  let low = 0;
  let high = matches.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    const page = queries.anchorRect(matches[middle].displayFrom)?.pageIndex;
    if (page != null && page < pageIndex) low = middle + 1;
    else high = middle;
  }
  return low < matches.length ? low : 0;
}

/**
 * Host-driven find over the live session: every body match is highlighted, the
 * current one is revealed without moving selection or focus, and a document
 * change re-runs the search on the next display list.
 */
export function useHostSearch({
  pagedEditorRef,
  displayListQueries,
  currentPage,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  displayListQueries: DisplayListQueries | null;
  /** The one-based page in view. */
  currentPage: () => number;
}): { api: DocxHostSearch; highlight: { matches: readonly CanvasFindMatch[]; current: number } | null } {
  const runRef = useRef<SearchRun | null>(null);
  const listenersRef = useRef(new Set<(state: DocxSearchState | null) => void>());
  const queriesRef = useRef(displayListQueries);
  queriesRef.current = displayListQueries;
  const currentPageRef = useRef(currentPage);
  currentPageRef.current = currentPage;
  const [highlight, setHighlight] = useState<{
    matches: readonly CanvasFindMatch[];
    current: number;
  } | null>(null);

  const publish = useCallback((run: SearchRun | null) => {
    runRef.current = run;
    setHighlight(run && run.matches.length > 0 ? { matches: run.matches, current: run.current } : null);
    const state = stateOf(run);
    for (const listener of [...listenersRef.current]) {
      try {
        listener(state);
      } catch (error) {
        console.error('[DocxEditor] search listener failed', error);
      }
    }
  }, []);

  const goTo = useCallback(
    (index: number): DocxSearchState | null => {
      const run = runRef.current;
      if (!run) return null;
      if (run.matches.length === 0 || !Number.isInteger(index)) return stateOf(run);
      const current = ((index % run.matches.length) + run.matches.length) % run.matches.length;
      pagedEditorRef.current?.revealDisplayPosition(run.matches[current].displayFrom);
      publish({ ...run, current });
      return stateOf(runRef.current);
    },
    [pagedEditorRef, publish]
  );

  const clearSearch = useCallback(() => {
    if (runRef.current) publish(null);
  }, [publish]);

  const search = useCallback(
    async (query: string, options: DocxSearchOptions = {}): Promise<DocxSearchState> => {
      const normalized = { caseSensitive: options.caseSensitive === true };
      const editor = pagedEditorRef.current;
      if (editor?.hasPendingInput()) await editor.flushPendingInput();
      const session = editor?.getYrsSession();
      if (!editor || !session || query === '') {
        publish(null);
        return { query, options: normalized, total: 0, current: -1 };
      }
      const matches = collectMatches(editor, session, query, normalized);
      const current = firstInView(matches, queriesRef.current, currentPageRef.current() - 1);
      const run: SearchRun = {
        query,
        options: normalized,
        session,
        version: session.version(),
        matches,
        current,
      };
      if (current >= 0) editor.revealDisplayPosition(matches[current].displayFrom);
      publish(run);
      return stateOf(run)!;
    },
    [pagedEditorRef, publish]
  );

  // A new display list follows every document change; the matches move with it.
  useEffect(() => {
    const run = runRef.current;
    const editor = pagedEditorRef.current;
    if (!run || !editor) return;
    const session = editor.getYrsSession();
    if (session !== run.session) {
      publish(null);
      return;
    }
    if (!session || session.version() === run.version) return;
    const previous = run.matches[run.current];
    const matches = collectMatches(editor, session, run.query, run.options);
    let current = previous
      ? matches.findIndex((match) => match.paraId === previous.paraId && match.start === previous.start)
      : -1;
    if (current < 0 && previous) {
      current = matches.findIndex((match) => match.displayFrom >= previous.displayFrom);
      if (current < 0) current = matches.length - 1;
    }
    publish({ ...run, version: session.version(), matches, current: matches.length ? Math.max(0, current) : -1 });
  }, [displayListQueries, pagedEditorRef, publish]);

  const api = useMemo<DocxHostSearch>(
    () => ({
      search,
      searchNext: () => {
        const run = runRef.current;
        return run ? goTo(run.current + 1) : null;
      },
      searchPrevious: () => {
        const run = runRef.current;
        return run ? goTo(run.current - 1) : null;
      },
      searchGoTo: (index) => {
        const run = runRef.current;
        if (!run) return null;
        return index >= 0 && index < run.matches.length ? goTo(index) : stateOf(run);
      },
      clearSearch,
      getSearchState: () => stateOf(runRef.current),
      onSearchChange: (listener) => {
        listenersRef.current.add(listener);
        return () => {
          listenersRef.current.delete(listener);
        };
      },
    }),
    [clearSearch, goTo, search]
  );

  return { api, highlight };
}
