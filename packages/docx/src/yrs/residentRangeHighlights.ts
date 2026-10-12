import type { YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import { DisplayPositionIndex } from './displayPositionIndex';
import { commandSelectionRange } from './residentCommandAnchors';
import { createDisplayListQueries, type DisplayListRect } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import type { RetainedFrame } from '../layout/render/frameDelta';

export interface ResidentRangeHighlights {
  version: string;
  layoutRevision: number;
  frameEpoch: number;
  rects: Array<Array<DisplayListRect & { pageId: string }>>;
  unbuiltPages: number[];
}

export async function readResidentRangeHighlights(
  engine: ResidentEngineSession, frame: RetainedFrame, ranges: readonly YrsStoryRange[], layoutRevision: number
): Promise<ResidentRangeHighlights> {
  const native = await loadRustDisplayListQueryEngine();
  const index = new DisplayPositionIndex({ ...engine.geometryReader, selectionText: engine.selectionText });
  const queries = createDisplayListQueries(frame.displayList, native);
  try {
    const rects = ranges.map(range => {
      const resolved = commandSelectionRange(engine, {
        anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end },
      });
      if (!resolved) return [];
      const root = range.story === 'body' || range.story.startsWith('body:') ? 'body'
        : range.story.split(':').slice(0, 2).join(':');
      const from = index.positionOf({ ...resolved.start, story: range.story }, root);
      const to = index.positionOf({ ...resolved.end, story: range.story }, root);
      if (from === null || to === null || from >= to) return [];
      const [kind, id] = root.split(':');
      const found = root === 'body' ? queries.rangeRects(from, to)
        : kind === 'hf' ? [...queries.hfRangeRects('header', id!, from, to), ...queries.hfRangeRects('footer', id!, from, to)]
        : kind === 'fn' || kind === 'en' ? queries.noteRangeRects(kind === 'fn' ? 'footnote' : 'endnote', Number(id), from, to) : [];
      return found.map(rect => ({ ...rect, pageId: frame.pages[rect.pageIndex]!.pageId.toString() }));
    });
    const state = queries.sourceState();
    if (state.status === 'error') throw new Error(`Range highlight layout is unavailable: ${String(state.error)}`);
    return { version: engine.proposalEngine.version(), layoutRevision, frameEpoch: frame.frameEpoch, rects,
      unbuiltPages: frame.pages.filter(page => page.page.unbuilt).map(page => page.pageIndex) };
  } finally { queries.dispose(); }
}
