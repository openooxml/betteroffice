import type { YrsRevisionInfo } from '../../yrs';
import type { TrackedChangeEntry } from '../../utils/comments';
import {
  type YrsSidebarProjection,
  yrsIdToNumericId,
} from './yrsSidebarProjection';

/** Tracked-change sidebar data derived from the authoritative Yrs session. */
export interface TrackedChangesResult {
  entries: TrackedChangeEntry[];
  commentToRevision: Map<number, number>;
}

/** Project Yrs revisions onto the display-position sidebar contract. */
export function extractTrackedChangesFromYrs(
  revisions: readonly YrsRevisionInfo[],
  projection: YrsSidebarProjection
): TrackedChangesResult {
  const mapped: TrackedChangeEntry[] = revisions.map((revision) => {
    const start = projection.locToDisplayPoint({
      story: revision.range.story,
      ...revision.range.start,
    });
    const end = projection.locToDisplayPoint({
      story: revision.range.story,
      ...revision.range.end,
    });
    const from = start?.position ?? 0;
    const to = Math.max(from, end?.position ?? from);
    const type: TrackedChangeEntry['type'] =
      revision.kind === 'pPrIns'
        ? 'paragraphMarkInsertion'
        : revision.kind === 'pPrDel'
          ? 'paragraphMarkDeletion'
          : revision.kind === 'pPrChange'
            ? 'paragraphPropertiesChanged'
            : revision.kind === 'trIns'
              ? 'rowInserted'
              : revision.kind === 'trDel'
                ? 'rowDeleted'
                : revision.kind === 'tableIns'
                  ? 'tableInserted'
                  : revision.kind === 'tableDel'
                    ? 'tableDeleted'
                    : revision.kind;

    return {
      type,
      text: revision.preview,
      author: revision.author,
      date: revision.date || undefined,
      from,
      to,
      revisionId: yrsIdToNumericId(revision.revisionId),
      ...(start?.hfRid ? { hfRid: start.hfRid } : {}),
    } as TrackedChangeEntry;
  });

  const byRevision = new Map<number, TrackedChangeEntry[]>();
  for (const entry of mapped) {
    const sites = byRevision.get(entry.revisionId) ?? [];
    sites.push(entry);
    byRevision.set(entry.revisionId, sites);
  }
  const coalesced = [...byRevision.values()].map((sites): TrackedChangeEntry => {
    const from = Math.min(...sites.map((site) => site.from));
    const to = Math.max(...sites.map((site) => site.to));
    const insertions = sites.filter((site) => site.type === 'insertion');
    const deletions = sites.filter((site) => site.type === 'deletion');
    if (insertions.length > 0 && deletions.length > 0) {
      const insertion = insertions[0]!;
      return {
        ...insertion,
        type: 'replacement',
        text: insertions.map((site) => site.text).join(''),
        deletedText: deletions.map((site) => site.text).join(''),
        from,
        to,
      };
    }
    const primary =
      sites.find((site) => site.type === 'insertion' || site.type === 'deletion') ?? sites[0]!;
    const matchingText = sites
      .filter((site) => site.type === primary.type)
      .map((site) => site.text)
      .join('');
    return { ...primary, text: matchingText, from, to };
  });

  const entries: TrackedChangeEntry[] = [];
  for (let index = 0; index < coalesced.length; index += 1) {
    const current = coalesced[index]!;
    const next = coalesced[index + 1];
    if (
      current.type === 'deletion' &&
      next?.type === 'insertion' &&
      current.author === next.author &&
      current.date === next.date &&
      current.to === next.from &&
      (current as { hfRid?: string }).hfRid === (next as { hfRid?: string }).hfRid
    ) {
      entries.push({
        type: 'replacement',
        text: next.text,
        deletedText: current.text,
        author: current.author,
        date: current.date,
        from: current.from,
        to: next.to,
        revisionId: current.revisionId,
        insertionRevisionId: next.revisionId,
        ...((current as { hfRid?: string }).hfRid
          ? { hfRid: (current as { hfRid?: string }).hfRid }
          : {}),
      } as TrackedChangeEntry);
      index += 1;
    } else {
      entries.push(current);
    }
  }

  return { entries, commentToRevision: new Map() };
}
