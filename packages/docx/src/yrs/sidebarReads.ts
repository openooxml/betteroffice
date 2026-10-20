import type { YrsLoc, YrsSession } from './index';
import type { TrackedChangeEntry } from '../utils/comments';
import {
  createYrsSidebarProjection,
  yrsIdToNumericId,
  type YrsSidebarDisplayPoint,
} from '../layout/render/yrsSidebarProjection';
import { extractTrackedChangesFromYrs } from '../layout/render/yrsTrackedChanges';

export type DocxSidebarReader = Pick<
  YrsSession,
  | 'version'
  | 'listRevisions'
  | 'resolveComment'
  | 'headings'
  | 'paragraphs'
  | 'storyIds'
  | 'storySegments'
  | 'locateParagraph'
>;

export interface DocxSidebarAnchorPoints {
  start: YrsSidebarDisplayPoint | null;
  end: YrsSidebarDisplayPoint | null;
}

export interface DocxSidebarRead {
  comments: Array<{ id: string; anchors: DocxSidebarAnchorPoints[] }>;
  revisions: Array<{ key: string; start: YrsSidebarDisplayPoint | null }>;
  trackedChanges: { entries: TrackedChangeEntry[]; commentToRevision: Array<[number, number]> };
}

export interface DocxOutlineHeading {
  story: string;
  paraId: string;
  text: string;
  level: number;
  position: number;
}

export function readSidebar(
  reader: DocxSidebarReader,
  commentIds: readonly string[],
  expectVersion: string
): DocxSidebarRead | null {
  if (reader.version() !== expectVersion) return null;
  const projection = createYrsSidebarProjection(reader);
  const revisions = reader.listRevisions();
  const tracked = extractTrackedChangesFromYrs(revisions, projection);
  return {
    comments: commentIds.map((id) => {
      try {
        return {
          id,
          anchors: reader.resolveComment(id).map((anchor) => ({
            start: projection.storyOffsetToDisplayPoint(anchor.story, anchor.start),
            end: projection.storyOffsetToDisplayPoint(anchor.story, anchor.end),
          })),
        };
      } catch {
        return { id, anchors: [] };
      }
    }),
    revisions: revisions.map((revision) => ({
      key: `revision-${yrsIdToNumericId(revision.revisionId)}`,
      start: projection.locToDisplayPoint({ story: revision.story, ...revision.range.start }),
    })),
    trackedChanges: { entries: tracked.entries, commentToRevision: [...tracked.commentToRevision] },
  };
}

/** Hybrid review navigation: the first change after the selection, or the last before it, wrapping at the ends. */
export function readAdjacentRevision(
  reader: DocxSidebarReader,
  selection: { anchor: YrsLoc; focus: YrsLoc } | null,
  direction: 'next' | 'previous',
  excluded: ReadonlySet<string>
): { revisionId: string } | null {
  const projection = createYrsSidebarProjection(reader);
  const positions = new Map<string, number>();
  for (const revision of reader.listRevisions()) {
    if (excluded.has(revision.revisionId) || positions.has(revision.revisionId)) continue;
    const point = projection.locToDisplayPoint({ story: revision.story, ...revision.range.start });
    if (point && !point.hfRid) positions.set(revision.revisionId, point.position);
  }
  const navigation = [...positions].sort((first, second) => first[1] - second[1]);
  if (!navigation.length) return null;
  const points = selection ? [selection.anchor, selection.focus].map((at) => projection.locToDisplayPoint(at)) : [];
  const bounds = points.length && points.every((point) => point && !point.hfRid) ? points.map((point) => point!.position) : null;
  const target = direction === 'next'
    ? navigation.find(([, position]) => position > (bounds ? Math.max(...bounds) : -1)) ?? navigation[0]!
    : [...navigation].reverse().find(([, position]) => position < (bounds ? Math.min(...bounds) : -1)) ?? navigation.at(-1)!;
  return { revisionId: target[0] };
}

export function readOutlineHeadings(
  reader: DocxSidebarReader,
  expectVersion: string
): DocxOutlineHeading[] | null {
  if (reader.version() !== expectVersion) return null;
  const levels = new Map(
    reader.headings('body').map((entry) => [entry.paraId, entry.heading.outlineLevel])
  );
  const headings: DocxOutlineHeading[] = [];
  if (levels.size === 0) return headings;
  const projection = createYrsSidebarProjection(reader);
  // A paragraph's text is its text segments up to its pilcrow, as `paragraphs()` reads it.
  const parts: string[] = [];
  for (const segment of reader.storySegments('body')) {
    if (segment.kind === 'text') {
      parts.push(segment.text);
      continue;
    }
    if (segment.kind !== 'pilcrow') continue;
    const level = levels.get(segment.paraId);
    const text = level == null ? '' : parts.join('').trim();
    parts.length = 0;
    if (level == null || !text) continue;
    const position = projection.locToDisplayPoint({
      story: 'body', paraId: segment.paraId, offset: 0,
    })?.position;
    if (position == null) continue;
    headings.push({ story: 'body', paraId: segment.paraId, text, level, position });
  }
  return headings;
}
