import {
  anchorPositionsFromPoints,
  type DisplayListQueries,
  type DisplayListRect,
  type TrackedChangesResult,
  type YrsHeaderFooterRegions,
  type YrsSidebarDisplayPoint,
} from '@betteroffice/docx/layout/render';
import type { DocxSidebarRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { readAt } from './viewerReads';

export type ViewerCommentRanges = Map<number, { from: number; to: number } | null>;

export const EMPTY_TRACKED_CHANGES_RESULT: TrackedChangesResult = {
  entries: [],
  commentToRevision: new Map(),
};

export function viewerCommentRanges(value: DocxSidebarRead): ViewerCommentRanges {
  return new Map(value.comments.map(({ id, anchors }): [number, { from: number; to: number } | null] => {
    const starts = anchors.flatMap(({ start }) => start ? [start] : []);
    const ends = anchors.flatMap(({ end }) => end ? [end] : []);
    const start = starts.sort((a, b) => a.position - b.position)[0];
    const end = ends.sort((a, b) => b.position - a.position)[0];
    return [Number(id), start && end && !start.hfRid
      ? { from: start.position, to: end.position }
      : null];
  }));
}

export function viewerSidebarPositions(
  value: DocxSidebarRead,
  queries: DisplayListQueries,
  hfRegions?: YrsHeaderFooterRegions,
  projectY?: (rect: DisplayListRect) => number | null
): Map<string, number> {
  const points: Array<readonly [string, YrsSidebarDisplayPoint | null]> = [];
  for (const { id, anchors } of value.comments) {
    for (const { start } of anchors) points.push([`comment-${id}`, start]);
  }
  for (const { key, start } of value.revisions) points.push([key, start]);
  return anchorPositionsFromPoints(points, queries, hfRegions, projectY);
}

export class ViewerSidebarReads {
  private cached: { key: string; value: Promise<DocxSidebarRead | null> } | null = null;
  private tracked: { version: string; result: TrackedChangesResult } | null = null;
  private deliveredVersion: string | null = null;

  constructor(private readonly read: ResidentEngineWorkerClient['documentRead']) {}

  async sidebar(
    version: string,
    commentIds: readonly string[],
    currentVersion: () => string | null
  ): Promise<DocxSidebarRead | null> {
    const key = JSON.stringify([version, commentIds]);
    if (this.cached?.key !== key) {
      this.cached = {
        key,
        value: readAt(this.read, { kind: 'sidebar', commentIds: [...commentIds], expectVersion: version })
          .then((reply) => reply.status === 'ok' ? reply.value : null),
      };
    }
    const pending = this.cached;
    const value = await pending.value;
    if (!value && this.cached === pending) this.cached = null;
    return currentVersion() === version ? value : null;
  }

  deliver(
    version: string,
    value: DocxSidebarRead,
    to: ((result: TrackedChangesResult) => void) | undefined
  ): void {
    if (this.tracked?.version !== version) {
      this.tracked = {
        version,
        result: value.comments.length === 0 && value.revisions.length === 0
          ? EMPTY_TRACKED_CHANGES_RESULT
          : { entries: value.trackedChanges.entries, commentToRevision: new Map(value.trackedChanges.commentToRevision) },
      };
    }
    if (!to || this.deliveredVersion === version) return;
    this.deliveredVersion = version;
    to(this.tracked.result);
  }
}
