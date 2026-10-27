import { yrsIdToNumericId, type DisplayListRect } from '@betteroffice/docx/layout/render';
import type { DocxDocumentRead, DocxDocumentReadValues, DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import type { DocxRevisionEnvironment } from '../../../commands/evaluate';
import type { CommentCallbacks } from '../../../hooks/useCommentSidebarItems';
import type { TrackedChangeEntry } from '../../sidebar/cardUtils';

export interface WorkerRevisionState {
  ids: ReadonlySet<string>;
  navigation: readonly DocxRevisionEnvironment[];
  entries: TrackedChangeEntry[];
  anchors: ReadonlyMap<string, DisplayListRect>;
}

const proposalRevisionIds = (review: DocxReviewSnapshot) => new Set(review.proposals.proposals.flatMap(proposal => proposal.revisionIds));

export function workerRevisionRequest(review: { version: string; value: DocxReviewSnapshot | null } | null): (DocxDocumentRead & { kind: 'sidebar' }) | null {
  const value = review?.value;
  if (!value) return null;
  const proposals = proposalRevisionIds(value);
  return value.revisions.some(revision => !proposals.has(revision.revisionId)) ? { kind: 'sidebar', commentIds: [], expectVersion: review.version } : null;
}

export function workerRevisions(review: DocxReviewSnapshot, sidebar: DocxDocumentReadValues['sidebar']): WorkerRevisionState | null {
  const proposals = proposalRevisionIds(review);
  const indexed = review.revisions.map((site, index) => ({ site, index })).filter(({ site }) => !proposals.has(site.revisionId));
  if (indexed.length && (!sidebar || sidebar.revisions.length !== review.revisions.length)) return null;
  const numeric = new Map(indexed.map(({ site }) => [yrsIdToNumericId(site.revisionId), site.revisionId]));
  const navigation = new Map<string, DocxRevisionEnvironment>();
  for (const { site, index } of indexed) {
    const point = sidebar!.revisions[index]!;
    if (point.key !== `revision-${yrsIdToNumericId(site.revisionId)}`) return null;
    if (point.start && point.start.hfRid === undefined && !navigation.has(site.revisionId)) navigation.set(site.revisionId, { revisionId: site.revisionId, position: point.start.position });
  }
  const anchors = new Map<string, DisplayListRect>();
  for (const [key, revisionId] of numeric) {
    const rect = review.revisionGeometry[revisionId]?.[0];
    if (rect) anchors.set(`revision-${key}`, rect);
  }
  return {
    ids: new Set(review.revisions.map(site => site.revisionId)),
    navigation: [...navigation.values()].sort((first, second) => first.position - second.position),
    entries: (sidebar?.trackedChanges.entries ?? []).filter(entry => numeric.has(entry.revisionId) &&
      (entry.insertionRevisionId == null || numeric.has(entry.insertionRevisionId))),
    anchors,
  };
}

export function workerRevisionCallbacks(controller: {
  trackedChanges(): WorkerRevisionState | null;
  commands: { store: { execute(id: 'reviewAccept' | 'reviewReject', args: { revisionId: string }): unknown } };
}): Pick<CommentCallbacks, 'onAcceptChangeById' | 'onRejectChangeById'> {
  const decide = (id: 'reviewAccept' | 'reviewReject', numeric: number) => {
    const revisionId = [...controller.trackedChanges()?.ids ?? []].find(candidate => yrsIdToNumericId(candidate) === numeric);
    if (revisionId) void controller.commands.store.execute(id, { revisionId });
  };
  return { onAcceptChangeById: numeric => decide('reviewAccept', numeric), onRejectChangeById: numeric => decide('reviewReject', numeric) };
}
