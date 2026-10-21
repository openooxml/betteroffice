import {
  createDisplayListQueries, displayPageRevision, displayPageNoteAnchorRevision,
  loadedRustDisplayListQueryEngine, type DisplayPage, type DisplayListQueries, type RustDisplayListQueryEngine,
} from '@betteroffice/docx/layout/render';

export function createPresentedPageQueries() {
  type PageQueries = Pick<DisplayListQueries, 'hitTestRegions' | 'caretRect'>;
  let presented: DisplayPage[] = [];
  let cached: {
    revisions: number[]; anchors: number[]; engine: RustDisplayListQueryEngine;
    queries: DisplayListQueries; pages: Map<DisplayPage, PageQueries>;
  } | null = null;
  const release = () => { cached?.queries.dispose(); cached = null; };
  const get = (page: DisplayPage): PageQueries | null => {
    const engine = loadedRustDisplayListQueryEngine();
    const index = presented.indexOf(page);
    if (!engine || page.unbuilt || index < 0) return null;
    if (cached && (cached.engine !== engine || presented.some((entry, slot) =>
      cached!.revisions[slot] !== displayPageRevision(entry) || cached!.anchors[slot] !== displayPageNoteAnchorRevision(entry)))) release();
    if (!cached) cached = {
      revisions: presented.map(displayPageRevision), anchors: presented.map(displayPageNoteAnchorRevision), engine,
      queries: createDisplayListQueries({ pages: presented }, engine), pages: new Map(),
    };
    const existing = cached.pages.get(page);
    if (existing) return existing;
    const queries = cached.queries;
    const local: PageQueries = presented.length === 1 ? queries : {
      hitTestRegions: (_pageIndex, x, y) => queries.hitTestRegions(index, x, y),
      caretRect(pos) {
        const forward = queries.rangeRectsOnPages!(pos, pos + 1, index, index);
        const first = forward[0];
        if (first) return { ...first, pageIndex: 0, width: 0 };
        if (pos <= 0) return null;
        const backward = queries.rangeRectsOnPages!(pos - 1, pos, index, index);
        const last = backward[backward.length - 1];
        return last ? { ...last, pageIndex: 0, x: last.x + last.width, width: 0 } : null;
      },
    };
    cached.pages.set(page, local);
    return local;
  };
  const retain = (pages: readonly DisplayPage[]) => {
    const next = [...new Set(pages)].filter((page) => !page.unbuilt);
    if (next.length === presented.length && next.every((page, index) => page === presented[index])) return;
    release();
    presented = next;
  };
  return { get, retain, dispose: () => { release(); presented = []; } };
}
