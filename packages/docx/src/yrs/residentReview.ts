import type { YrsCommentInfo, YrsLoc, YrsRevisionInfo, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxProposalRegistry, DocxProposalSnapshot } from './proposals';
import { computeProposalGeometryMirror, type ProposalGeometryMirror } from './proposalGeometry';
import { DisplayPositionIndex } from './displayPositionIndex';
import { createDisplayListQueries, type DisplayListQueries, type DisplayListRect } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { Comment } from '../types/content';

export interface DocxCommentPatch {
  body?: unknown;
  author?: string;
  date?: string;
  done?: boolean;
  parentId?: string | null;
}

/** A comment, by the id saves and the host use, or a revision, by its session id. */
export type ResidentReviewTarget = { kind: 'comment'; id: number } | { kind: 'revision'; revisionId: string };

/** The session id and story ranges of a review target, or null when it does not exist. */
export interface ResidentReviewTargetRanges {
  id: string;
  ranges: YrsStoryRange[];
}

export interface ResidentReviewSnapshot {
  version: string;
  previewVersion: number;
  layoutRevision: number;
  frameEpoch: number;
  comments: YrsCommentInfo[];
  uiComments?: Comment[];
  commentIds?: Record<string, number>;
  revisions: YrsRevisionInfo[];
  proposals: DocxProposalSnapshot;
  proposalGeometry: ProposalGeometryMirror;
  commentGeometry: Record<string, DisplayListRect[]>;
  revisionGeometry: Record<string, DisplayListRect[]>;
  proposalHighlights: Record<string, DisplayListRect[]>;
}

function rangeHighlightGeometry(
  index: DisplayPositionIndex, queries: DisplayListQueries, ranges: readonly YrsStoryRange[], anchorCollapsed = false
): DisplayListRect[][] {
  return ranges.map(range => {
    const root = range.story === 'body' || range.story.startsWith('body:') ? 'body'
      : range.story.split(':').slice(0, 2).join(':');
    const from = index.positionOf({ story: range.story, ...range.start }, root);
    const to = index.positionOf({ story: range.story, ...range.end }, root);
    if (from === null || to === null || from > to || (from === to && !anchorCollapsed)) return [];
    const [kind, id] = root.split(':');
    if (from === to) {
      if (root === 'body') return [queries.anchorRect(from)].filter((rect): rect is DisplayListRect => !!rect);
      if (kind === 'hf') return [...queries.hfAnchorRects('header', id!, from), ...queries.hfAnchorRects('footer', id!, from)];
      if (kind === 'fn' || kind === 'en') return queries.noteCaretRects(kind === 'fn' ? 'footnote' : 'endnote', Number(id), from);
      return [];
    }
    if (root === 'body') return queries.rangeRects(from, to);
    if (kind === 'hf') return [...queries.hfRangeRects('header', id!, from, to), ...queries.hfRangeRects('footer', id!, from, to)];
    if (kind === 'fn' || kind === 'en') return queries.noteRangeRects(kind === 'fn' ? 'footnote' : 'endnote', Number(id), from, to);
    return [];
  });
}

function locationAt(engine: ResidentEngineSession, story: string, offset: number): YrsLoc | null {
  for (const paragraph of engine.geometryReader.paragraphSpans(story)) {
    const span = engine.geometryReader.locateParagraph(story, paragraph.paraId);
    if (offset >= span.start && offset <= span.end) return { story, paraId: paragraph.paraId, offset: offset - span.start };
  }
  return null;
}

export function readReviewTarget(engine: ResidentEngineSession, target: ResidentReviewTarget): ResidentReviewTargetRanges | null {
  if (target.kind === 'revision') {
    const ranges = engine.geometryReader.listRevisions().filter(revision => revision.revisionId === target.revisionId)
      .map(({ story, range }) => ({ story, start: range.start, end: range.end }));
    return ranges.length ? { id: target.revisionId, ranges } : null;
  }
  const projection = engine.commentProjection();
  const id = projection?.sessionIds.get(target.id) ?? String(target.id);
  if (!engine.listComments().some(comment => comment.id === id)) return null;
  let anchors: ReturnType<ResidentEngineSession['resolveComment']> = [];
  const visited = new Set<number>();
  let publicId: number | undefined = target.id;
  while (publicId !== undefined && !visited.has(publicId)) {
    visited.add(publicId);
    try { anchors = engine.resolveComment(projection?.sessionIds.get(publicId) ?? String(publicId)); }
    catch (error) { if (error instanceof WebAssembly.RuntimeError) throw error; anchors = []; }
    if (anchors.length) break;
    publicId = projection?.comments.find(comment => comment.id === publicId)?.parentId;
  }
  return { id, ranges: anchors.flatMap(anchor => {
    const start = locationAt(engine, anchor.story, anchor.start);
    const end = locationAt(engine, anchor.story, anchor.end);
    return start && end ? [{ story: anchor.story, start: { paraId: start.paraId, offset: start.offset }, end: { paraId: end.paraId, offset: end.offset } }] : [];
  }) };
}

const reviewSnapshots = new WeakMap<RetainedFrame, { engine: ResidentEngineSession; key: string; snapshot: ResidentReviewSnapshot }>();

export async function readResidentReviewSnapshot(
  engine: ResidentEngineSession, registry: DocxProposalRegistry, frame: RetainedFrame, layoutRevision: number,
  reads?: { queries?: (frame: RetainedFrame, version: string) => DisplayListQueries | undefined; index?: DisplayPositionIndex }
): Promise<ResidentReviewSnapshot> {
  const native = await loadRustDisplayListQueryEngine();
  const keyOf = () => engine.commentRecoveryState && JSON.stringify([
    engine.proposalEngine.version(), frame.frameEpoch, layoutRevision, registry.snapshot(), engine.commentRecoveryState(),
  ]);
  const cached = reviewSnapshots.get(frame);
  if (cached?.engine === engine && cached.key === keyOf()) return cached.snapshot;
  const borrowed = reads?.queries?.(frame, engine.proposalEngine.version());
  const queries = borrowed ?? createDisplayListQueries(frame.displayList, native);
  const index = reads?.index ?? new DisplayPositionIndex({ ...engine.geometryReader, selectionText: engine.selectionText });
  try {
    const proposals = registry.snapshot();
    const comments = engine.listComments();
    const projection = engine.commentProjection();
    const revisions = engine.geometryReader.listRevisions();
    const commentGeometry = Object.fromEntries(comments.map(({ id }) => {
      let anchors: ReturnType<ResidentEngineSession['resolveComment']>;
      try { anchors = engine.resolveComment(id); }
      catch (error) { if (error instanceof WebAssembly.RuntimeError) throw error; anchors = []; }
      const ranges = anchors.flatMap(anchor => {
        const start = locationAt(engine, anchor.story, anchor.start);
        const end = locationAt(engine, anchor.story, anchor.end);
        return start && end ? [{ story: anchor.story, start, end }] : [];
      });
      return [id, rangeHighlightGeometry(index, queries, ranges).flat()];
    }));
    const revisionGeometry: Record<string, DisplayListRect[]> = {};
    for (const revision of revisions) {
      (revisionGeometry[revision.revisionId] ??= []).push(...rangeHighlightGeometry(index, queries, [revision.range], true).flat());
    }
    const proposalGeometry = computeProposalGeometryMirror(engine.geometryReader, proposals);
    const proposalHighlights = Object.fromEntries(Object.entries(proposalGeometry.targets).map(([id, target]) => [
      id, target.ok ? target.ranges.flatMap(range => queries.rangeRects(range.from, range.to)) : [],
    ]));
    const state = queries.sourceState();
    if (state.status === 'error') throw new Error(`Review highlight layout is unavailable: ${String(state.error)}`);
    const snapshot: ResidentReviewSnapshot = {
      version: proposals.version, previewVersion: proposals.previewVersion, layoutRevision,
      frameEpoch: frame.frameEpoch, comments, ...(projection ? { uiComments: projection.comments, commentIds: projection.ids } : {}),
      revisions, proposals, proposalGeometry,
      commentGeometry, revisionGeometry, proposalHighlights,
    };
    // Keyed after the compute: the comment projection assigns aliases for new comments.
    const key = keyOf();
    if (key) reviewSnapshots.set(frame, { engine, key, snapshot });
    return snapshot;
  } finally { if (!borrowed) queries.dispose(); }
}
