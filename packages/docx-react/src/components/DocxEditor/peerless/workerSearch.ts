import { findAllMatches, type FindMatch, type FindOptions, type FindResult } from '@betteroffice/docx/utils/findReplace';
import type { DocxTextRange, YrsStoryRange, YrsStickyPosition } from '@betteroffice/docx/yrs';
import type { DocxHostSearch, DocxSearchState } from '../hooks/useHostSearch';
import type { createPeerlessShellController } from './shellController';
import type { DocxEchoRect } from '@betteroffice/docx/worker';

type Hit = { displayFrom: number; range: YrsStoryRange; ordinal: number; match: FindMatch };
type Run = { query: string; options: FindOptions; token: string; version: number; session: string; hits: Hit[]; all: Hit[]; current: number; starting: boolean; anchor?: YrsStickyPosition | null };

export function createWorkerSearch(controller: ReturnType<typeof createPeerlessShellController>, reveal: (rect: DocxEchoRect) => void) {
  let run: Run | null = null;
  let generation = 0;
  let pendingReplaces = 0;
  let disposed = false;
  let state = { isOpen: false, replaceMode: false, result: null as FindResult | null, highlights: [] as Array<DocxEchoRect & { current: boolean }> };
  const listeners = new Set<() => void>();
  const hostListeners = new Set<(state: DocxSearchState | null) => void>();
  const hostState = (): DocxSearchState | null => run ? { query: run.query, options: { caseSensitive: run.options.matchCase }, total: run.hits.length, current: run.current } : null;
  const publish = () => {
    state = { ...state, result: run ? { matches: run.hits.map(hit => hit.match), totalCount: run.hits.length, currentIndex: Math.max(0, run.current) } : null };
    listeners.forEach(listener => listener());
    hostListeners.forEach(listener => { try { listener(hostState()); } catch {} });
  };
  const clear = () => { generation++; run = null; state = { ...state, highlights: [] }; publish(); };
  const current = (candidate: Run) => {
    const view = controller.getSnapshot().view;
    return !disposed && run === candidate && view?.identity.session === candidate.session && view.version === candidate.version;
  };
  const select = (candidate: Run) => {
    const hit = candidate.hits[candidate.current];
    if (!state.isOpen || !hit || !current(candidate)) return;
    void controller.submit({ kind: 'select', payload: { selection: {
      anchor: { story: hit.range.story, ...hit.range.start }, focus: { story: hit.range.story, ...hit.range.end },
    } } }).catch(() => {});
  };
  const paragraphs = async (candidate: Run) => {
    const groups = new Map<string, Set<string>>();
    for (const hit of candidate.all) {
      const ids = groups.get(hit.range.story) ?? new Set<string>();
      ids.add(hit.range.start.paraId); groups.set(hit.range.story, ids);
    }
    const found = new Map<string, string>();
    for (const [story, ids] of groups) {
      const result = await controller.readAt({ kind: 'readParagraphs', request: { story, paraIds: [...ids], view: 'accepted' } }, candidate.token);
      if (!result.ok) throw new Error('Search text is unavailable');
      for (const paragraph of result.paragraphs) found.set(`${story}\n${paragraph.paraId}`, paragraph.text);
    }
    return found;
  };
  const highlights = async (candidate: Run, scroll: boolean) => {
    const geometry = controller.getSnapshot().view?.geometry;
    if (!geometry || !current(candidate)) return;
    const result = await controller.readAt({ kind: 'rangeHighlights', ranges: candidate.hits.map(hit => hit.range),
      expectVersion: candidate.token, expectFrameEpoch: geometry.frameEpoch, expectLayoutRevision: geometry.layoutRevision }, candidate.token);
    if (!current(candidate) || geometry !== controller.getSnapshot().view?.geometry || !result || result.version !== candidate.token ||
      result.frameEpoch !== geometry.frameEpoch || result.layoutRevision !== geometry.layoutRevision) return;
    if (candidate.starting) {
      const pageIndex = controller.ref.getCurrentPage() - 1;
      const first = candidate.hits.findIndex((hit, index) => result.rects[index]?.some(rect => rect.pageIndex >= pageIndex) ||
        geometry.pages.some(page => page.page.pageIndex >= pageIndex && page.displayPage?.positionSpan &&
          hit.displayFrom >= page.displayPage.positionSpan[0] && hit.displayFrom <= page.displayPage.positionSpan[1]));
      if (first >= 0) candidate.current = first;
      candidate.starting = false;
    }
    const rects = result.rects.flatMap((rects, index) => rects.flatMap(rect => {
      const page = geometry.pages.find(page => page.page.pageId === rect.pageId && page.page.pageIndex === rect.pageIndex);
      return page ? [{ x: page.bounds.x + rect.x, y: page.bounds.y + rect.y, width: rect.width, height: rect.height, current: index === candidate.current }] : [];
    }));
    state = { ...state, highlights: rects }; publish();
    const target = rects.find(rect => rect.current);
    if (scroll && target) reveal(target);
    else if (scroll) {
      const hit = candidate.hits[candidate.current];
      const pages = hit ? geometry.pages.filter(page => page.displayPage?.unbuilt && page.displayPage.positionSpan &&
        hit.displayFrom >= page.displayPage.positionSpan[0] && hit.displayFrom <= page.displayPage.positionSpan[1]) : [];
      if (pages.length && hit) {
        const low = Math.max(...pages.map(page => page.displayPage!.positionSpan![0]));
        const high = Math.min(...pages.map(page => page.displayPage!.positionSpan![1]));
        const share = high > low ? (hit.displayFrom - low) / (high - low + 1) : 1;
        const page = pages[Math.min(pages.length - 1, Math.floor(share * pages.length))]!;
        reveal({ x: page.bounds.x + (page.displayPage?.contentBounds?.x ?? 0), y: page.bounds.y + (page.displayPage?.contentBounds?.y ?? 0), width: 0, height: 0 });
      }
    }
  };
  const remember = async (candidate: Run) => {
    const index = candidate.current;
    const hit = candidate.hits[index];
    if (!hit) return;
    try {
      const result = await controller.readAt({ kind: 'stickyAnchors', locs: [{ story: hit.range.story, ...hit.range.start }], version: candidate.token }, candidate.token);
      if (current(candidate) && candidate.current === index) candidate.anchor = result[0] ?? null;
    } catch {}
  };
  const find = async (query: string, options: FindOptions, carry = false): Promise<DocxSearchState> => {
    const request = ++generation;
    const previous = run;
    run = null; state = { ...state, highlights: [] }; publish();
    if (!query) return { query, options: { caseSensitive: options.matchCase }, total: 0, current: -1 };
    const token = await controller.ref.readParagraphs({ view: 'accepted', paraIds: [] });
    if (!token.ok) throw new Error('Search version is unavailable');
    const view = controller.getSnapshot().view;
    if (!view) throw new Error('Document was replaced');
    const result = await controller.readAt({ kind: 'searchText', query, caseSensitive: options.matchCase, carry: carry ? previous?.anchor : undefined }, token.version);
    if (request !== generation || disposed) throw new Error('Search was superseded');
    const counts = new Map<string, number>();
    const hits = result.matches.map((hit, index): Hit => {
      const key = `${hit.story}\n${hit.paraId}`;
      const ordinal = counts.get(key) ?? 0; counts.set(key, ordinal + 1);
      return { ordinal, displayFrom: hit.displayFrom, range: { story: hit.story, start: { paraId: hit.paraId, offset: hit.start }, end: { paraId: hit.paraId, offset: hit.start + query.length } },
        match: { paragraphIndex: index, contentIndex: 0, startOffset: hit.start, endOffset: hit.start + query.length, text: query } };
    });
    const candidate: Run = { query, options, token: token.version, version: view.version, session: view.identity.session, hits, all: hits, current: hits.length ? 0 : -1, starting: !carry };
    if (options.matchWholeWord) {
      const text = await paragraphs(candidate);
      candidate.hits = hits.filter(hit => {
        const value = text.get(`${hit.range.story}\n${hit.range.start.paraId}`) ?? '';
        const plain = findAllMatches(value, query, { ...options, matchWholeWord: false });
        const native = hits.filter(other => other.range.story === hit.range.story && other.range.start.paraId === hit.range.start.paraId);
        const at = plain[hit.ordinal];
        return plain.length === native.length && at && findAllMatches(value, query, options).some(word => word.start === at.start && word.end === at.end);
      });
    }
    if (request !== generation || disposed) throw new Error('Search was superseded');
    const old = carry ? previous?.hits[previous.current] : undefined;
    const carried = old ? candidate.hits.findIndex(hit => hit.range.story === old.range.story && hit.range.start.paraId === old.range.start.paraId && hit.range.start.offset >= old.range.start.offset) : -1;
    const anchored = carry && previous?.anchor && result.carried >= 0 ? candidate.hits.indexOf(candidate.all[result.carried]!) : -1;
    candidate.current = candidate.hits.length ? Math.max(0, anchored >= 0 ? anchored : carried) : -1;
    run = candidate; publish();
    await highlights(candidate, true);
    await remember(candidate);
    if (request !== generation || disposed) throw new Error('Search was superseded');
    return hostState()!;
  };
  const goTo = (index: number) => {
    if (!run?.hits.length || !Number.isInteger(index)) return hostState();
    run.current = ((index % run.hits.length) + run.hits.length) % run.hits.length;
    run.anchor = null;
    state = { ...state, highlights: [] }; publish();
    void remember(run);
    void highlights(run, true).catch(() => {});
    return hostState();
  };
  const host: DocxHostSearch = {
    search: (query, options) => find(query, { matchCase: options?.caseSensitive ?? false, matchWholeWord: false }),
    searchNext: () => goTo((run?.current ?? -1) + 1), searchPrevious: () => goTo((run?.current ?? 0) - 1), searchGoTo: goTo,
    clearSearch: clear, getSearchState: hostState,
    onSearchChange(listener) { hostListeners.add(listener); return () => hostListeners.delete(listener); },
  };
  let geometry = controller.getSnapshot().view?.geometry;
  const unsubscribe = controller.subscribe(() => {
    const view = controller.getSnapshot().view;
    if (!run) return;
    if (!view || view.identity.session !== run.session) return clear();
    if (view.version !== run.version) {
      const previous = run;
      void find(previous.query, previous.options, true).catch(() => {});
    } else if (view.geometry !== geometry) {
      geometry = view.geometry; state = { ...state, highlights: [] }; publish();
      void highlights(run, false).catch(() => {});
    }
  });
  return {
    host, getSnapshot: () => state,
    async find(query: string, options: FindOptions): Promise<DocxSearchState> {
      if (!state.isOpen) return { query, options: { caseSensitive: options.matchCase }, total: 0, current: -1 };
      return find(query, options);
    },
    subscribe(listener: () => void) { listeners.add(listener); return () => listeners.delete(listener); },
    open(replaceMode: boolean) { state = { ...state, isOpen: true, replaceMode }; publish(); },
    close(query: string, options: FindOptions) {
      if (!pendingReplaces && run && run.query === query && run.options.matchCase === options.matchCase && run.options.matchWholeWord === options.matchWholeWord) select(run);
      state = { ...state, isOpen: false }; clear();
    },
    async replace(text: string, all = false): Promise<number> {
      const candidate = run;
      if (!candidate || !current(candidate) || controller.environment().readOnly || controller.environment().mode === 'viewing') return 0;
      const request = generation;
      pendingReplaces++;
      try {
        await controller.ref.flushPendingInput();
        if (!current(candidate)) return 0;
        const projection = await paragraphs(candidate);
        const targets: DocxTextRange[] = [];
        for (const hit of all ? candidate.hits : candidate.hits.slice(candidate.current, candidate.current + 1)) {
          const value = projection.get(`${hit.range.story}\n${hit.range.start.paraId}`) ?? '';
          const matches = findAllMatches(value, candidate.query, { ...candidate.options, matchWholeWord: false });
          const native = candidate.all.filter(other => other.range.story === hit.range.story && other.range.start.paraId === hit.range.start.paraId);
          const match = matches[hit.ordinal];
          if (!match || matches.length !== native.length) return 0;
          const result = await controller.readAt({ kind: 'findText', request: { text: value.slice(match.start, match.end),
            within: { kind: 'paragraph', story: hit.range.story, paraId: hit.range.start.paraId }, view: 'accepted', limit: 10000 } }, candidate.token);
          if (!result.ok || result.truncated) return 0;
          const target = result.matches.find(item => item.range.start.offset === match.start && item.range.end.offset === match.end)?.range;
          if (!target) return 0;
          targets.push(target);
        }
        if (!current(candidate)) return 0;
        const result = await controller.submit({ kind: 'replaceMatches', payload: { targets, text, expectVersion: candidate.token } });
        if (request === generation && !disposed) await find(candidate.query, candidate.options, true);
        return result.ok && result.status === 'executed' ? targets.length : 0;
      } finally {
        pendingReplaces--;
      }
    },
    dispose() { disposed = true; generation++; unsubscribe(); listeners.clear(); hostListeners.clear(); },
  };
}
