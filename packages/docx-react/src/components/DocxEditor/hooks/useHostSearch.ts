import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { displayPageCanvases, type DisplayListQueries } from '@betteroffice/docx/layout/render';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import type { YrsSession, YrsStickyPosition } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';
import type { CanvasFindMatch } from '../overlays/CanvasFindHighlightOverlay';
import { scrollViewport } from '../internals/viewportBand';
import { sourceVersionOf } from '../internals/layoutProvenance';
import {
  displayOrder,
  matchesInRange,
  pagePositionIntervals,
} from '../overlays/CanvasFindHighlightOverlay';

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
  story: string;
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
  /** Where the current match starts, carried across document changes. */
  anchor: YrsStickyPosition | null;
  /** Whether a layout of `version` was shown when the display ranges were mapped. */
  placed: boolean;
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
    matches.push({ displayFrom, displayTo, story: hit.story, paraId: hit.paraId, start: hit.start });
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

/** The index of the topmost page showing in the scroller, from the painted page canvases. */
export function topPageInView(host: HTMLElement | null): number {
  if (!host) return 0;
  const viewport = scrollViewport(findVerticalScrollParentOrRoot(host));
  for (const canvas of displayPageCanvases(host)) {
    if (canvas.getBoundingClientRect().bottom > viewport.top + 1) {
      return Number(canvas.dataset.pageIndex) || 0;
    }
  }
  return 0;
}

/**
 * The first match `pageIndex` paints that a reveal shows there or later: a repeated table
 * header paints matches whose reveal goes back to the table's start.
 */
function firstOnPage(
  matches: readonly SearchMatch[],
  order: ReturnType<typeof displayOrder>,
  queries: DisplayListQueries,
  pageIndex: number,
  fromPage: number
): number {
  for (const { from, to } of pagePositionIntervals(queries.displayList, { start: pageIndex, end: pageIndex })) {
    for (const index of matchesInRange(matches, order, from, to)) {
      const page = queries.anchorRect(matches[index].displayFrom)?.pageIndex;
      if (page == null || page >= fromPage) return index;
    }
  }
  return -1;
}

/** How many candidate matches past the look-ahead have their page read. */
const FALLBACK_PAGE_READS = 64;

/** The first match on or after the page in view, wrapping to the first. */
function firstInView(
  matches: readonly SearchMatch[],
  queries: DisplayListQueries | null,
  pageIndex: number
): number {
  if (matches.length === 0) return -1;
  if (!queries || pageIndex <= 0) return 0;
  // A table row split across pages puts page indices out of position order, so the positions
  // each page paints decide first, the page in view and then the next ones.
  const order = displayOrder(matches);
  const pages = queries.displayList.pages.length;
  for (let index = pageIndex; index < Math.min(pages, pageIndex + 32); index += 1) {
    const first = firstOnPage(matches, order, queries, index, pageIndex);
    if (first >= 0) return first;
  }
  // Further on, the candidates are the matches from the first position painted on or after
  // the page in view (or after the last one painted before it), in position order: their
  // pages are read only until one is on or after that page.
  let start: number | undefined;
  for (let index = pageIndex; index < pages && start === undefined; index += 1) {
    start = pagePositionIntervals(queries.displayList, { start: index, end: index })[0]?.from;
  }
  for (let index = Math.min(pageIndex, pages) - 1; index >= 0 && start === undefined; index -= 1) {
    start = pagePositionIntervals(queries.displayList, { start: index, end: index }).at(-1)?.to;
  }
  const candidates = matchesInRange(matches, order, start ?? 0, Number.POSITIVE_INFINITY);
  for (const index of candidates.slice(0, FALLBACK_PAGE_READS)) {
    const page = queries.anchorRect(matches[index].displayFrom)?.pageIndex;
    if (page == null || page >= pageIndex) return index;
  }
  return candidates[0] ?? 0;
}

function anchorOf(session: YrsSession, match: SearchMatch | undefined): YrsStickyPosition | null {
  if (!match) return null;
  try {
    return session.encodeStickyPosition({ story: match.story, paraId: match.paraId, offset: match.start });
  } catch {
    return null;
  }
}

/** The match that starts at the carried anchor, else the first one after it, else the last. */
function carriedCurrent(
  editor: PagedEditorRef,
  session: YrsSession,
  matches: readonly SearchMatch[],
  anchor: YrsStickyPosition | null
): number {
  if (matches.length === 0) return -1;
  let loc = null;
  try {
    loc = anchor ? session.resolveStickyPosition(anchor) : null;
  } catch {
    loc = null;
  }
  if (!loc) return 0;
  const exact = matches.findIndex(
    (match) => match.story === loc.story && match.paraId === loc.paraId && match.start === loc.offset
  );
  if (exact >= 0) return exact;
  const position = editor.yrsLocToDisplayPosition(loc);
  const after = position == null ? -1 : matches.findIndex((match) => match.displayFrom >= position);
  return after >= 0 ? after : matches.length - 1;
}

/**
 * Host-driven find over the live session: every body match is highlighted, the
 * current one is revealed without moving selection or focus, and a document
 * change re-runs the search on the next display list.
 */
export function useHostSearch({
  pagedEditorRef,
  displayListQueries,
  canvasHostRef,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  displayListQueries: DisplayListQueries | null;
  /** `.canvas-pages` host; its painted pages tell which page is in view. */
  canvasHostRef: React.RefObject<HTMLElement | null>;
}): { api: DocxHostSearch; highlight: { matches: readonly CanvasFindMatch[]; current: number } | null } {
  const runRef = useRef<SearchRun | null>(null);
  const listenersRef = useRef(new Set<(state: DocxSearchState | null) => void>());
  const queriesRef = useRef(displayListQueries);
  queriesRef.current = displayListQueries;
  // Bumped by every search and clear; a search that resumes after another started stops.
  const generationRef = useRef(0);
  // A reveal the layout could not place yet, retried as the layout grows.
  const pendingRevealRef = useRef<number | null>(null);
  const [highlight, setHighlight] = useState<{
    matches: readonly CanvasFindMatch[];
    current: number;
  } | null>(null);

  /** Makes `run` current and tells subscribers; a listener that replaces it ends the round. */
  const publish = useCallback((run: SearchRun | null) => {
    runRef.current = run;
    setHighlight(run ? { matches: run.matches, current: run.current } : null);
    const state = stateOf(run);
    for (const listener of [...listenersRef.current]) {
      if (runRef.current !== run) return;
      try {
        listener(state);
      } catch (error) {
        console.error('[DocxEditor] search listener failed', error);
      }
    }
  }, []);

  // Stops the editor following the last reveal onto a page that is still being built.
  const revealAbortRef = useRef<AbortController | null>(null);
  const stopRevealing = useCallback(() => {
    pendingRevealRef.current = null;
    revealAbortRef.current?.abort();
    revealAbortRef.current = null;
  }, []);

  /**
   * Scrolls to `position`, and keeps it pending until a layout of `version` has placed it:
   * a layout of another version may show the match elsewhere.
   */
  const reveal = useCallback(
    (position: number, version: string) => {
      stopRevealing();
      const abort = new AbortController();
      revealAbortRef.current = abort;
      const outcome = pagedEditorRef.current?.revealDisplayPosition(position, abort.signal);
      const shown = sourceVersionOf(queriesRef.current);
      const placed = shown === null ? runRef.current?.placed !== false : shown === version;
      pendingRevealRef.current = outcome === 'scrolled' && placed ? null : position;
    },
    [pagedEditorRef, stopRevealing]
  );

  const clearSearch = useCallback(() => {
    generationRef.current += 1;
    stopRevealing();
    if (runRef.current) publish(null);
  }, [publish, stopRevealing]);

  /**
   * `run` against the session as it is now, keeping its current match. `fresh` says whether a
   * layout of an unknown version counts as showing the current one.
   */
  const refreshed = useCallback(
    (run: SearchRun, editor: PagedEditorRef, session: YrsSession, fresh: boolean): SearchRun => {
      const matches = collectMatches(editor, session, run.query, run.options);
      const current = carriedCurrent(editor, session, matches, run.anchor);
      const version = session.version();
      const shown = sourceVersionOf(queriesRef.current);
      return {
        ...run,
        version,
        matches,
        current,
        anchor: anchorOf(session, matches[current]),
        placed: shown === null ? fresh : shown === version,
      };
    },
    []
  );

  /**
   * The run as of the current document, or null after clearing it when its editor or document
   * is gone.
   */
  const liveRun = useCallback((): SearchRun | null => {
    const run = runRef.current;
    if (!run) return null;
    const editor = pagedEditorRef.current;
    const session = editor?.getYrsSession();
    if (!editor || session !== run.session) {
      clearSearch();
      return null;
    }
    if (session.version() === run.version) return run;
    const next = refreshed(run, editor, session, false);
    publish(next);
    return runRef.current === next ? next : null;
  }, [clearSearch, pagedEditorRef, publish, refreshed]);

  const goTo = useCallback(
    (index: number): DocxSearchState | null => {
      const run = liveRun();
      if (!run) return null;
      if (run.matches.length === 0 || !Number.isInteger(index)) return stateOf(run);
      const current = ((index % run.matches.length) + run.matches.length) % run.matches.length;
      const next = { ...run, current, anchor: anchorOf(run.session, run.matches[current]) };
      publish(next);
      if (runRef.current === next) reveal(run.matches[current].displayFrom, run.version);
      return stateOf(runRef.current);
    },
    [liveRun, publish, reveal]
  );

  const search = useCallback(
    async (query: string, options: DocxSearchOptions = {}): Promise<DocxSearchState> => {
      const normalized = { caseSensitive: options.caseSensitive === true };
      const generation = (generationRef.current += 1);
      const empty = { query, options: normalized, total: 0, current: -1 };
      const pending = pagedEditorRef.current;
      if (pending?.hasPendingInput()) {
        // a document replaced mid-flush leaves nothing to search in; the checks below see that
        await pending.flushPendingInput().catch(() => undefined);
      }
      if (generation !== generationRef.current) return stateOf(runRef.current) ?? empty;
      const editor = pagedEditorRef.current;
      const session = editor?.getYrsSession();
      stopRevealing();
      if (!editor || !session || query === '') {
        publish(null);
        return empty;
      }
      const matches = collectMatches(editor, session, query, normalized);
      const current = firstInView(matches, queriesRef.current, topPageInView(canvasHostRef.current));
      const shown = sourceVersionOf(queriesRef.current);
      const run: SearchRun = {
        query,
        options: normalized,
        session,
        version: session.version(),
        matches,
        current,
        anchor: anchorOf(session, matches[current]),
        placed: shown === null || shown === session.version(),
      };
      publish(run);
      if (runRef.current !== run) return stateOf(runRef.current) ?? empty;
      if (current >= 0) reveal(matches[current].displayFrom, run.version);
      return stateOf(run)!;
    },
    [canvasHostRef, pagedEditorRef, publish, reveal, stopRevealing]
  );

  // A new display list follows every document change, and every page the layout adds.
  useEffect(() => {
    const run = runRef.current;
    if (!run) return;
    const editor = pagedEditorRef.current;
    const session = editor?.getYrsSession();
    if (!editor || session !== run.session) {
      clearSearch();
      return;
    }
    if (session.version() !== run.version || !run.placed) {
      const revealing = pendingRevealRef.current !== null;
      // positions moved: never keep following the old one onto an unbuilt page
      stopRevealing();
      const next = refreshed(run, editor, session, true);
      publish(next);
      if (revealing && runRef.current === next && next.current >= 0) {
        reveal(next.matches[next.current].displayFrom, next.version);
      }
    } else if (pendingRevealRef.current !== null) {
      reveal(pendingRevealRef.current, run.version);
    }
  }, [clearSearch, displayListQueries, pagedEditorRef, publish, refreshed, reveal, stopRevealing]);

  useEffect(() => clearSearch, [clearSearch]);

  const api = useMemo<DocxHostSearch>(
    () => ({
      search,
      searchNext: () => {
        const run = liveRun();
        return run ? goTo(run.current + 1) : null;
      },
      searchPrevious: () => {
        const run = liveRun();
        return run ? goTo(run.current - 1) : null;
      },
      searchGoTo: goTo,
      clearSearch,
      getSearchState: () => stateOf(liveRun()),
      onSearchChange: (listener) => {
        listenersRef.current.add(listener);
        return () => {
          listenersRef.current.delete(listener);
        };
      },
    }),
    [clearSearch, goTo, liveRun, search]
  );

  return { api, highlight };
}
