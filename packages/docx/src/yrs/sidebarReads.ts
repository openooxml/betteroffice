import type { YrsSession } from './index';
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

export function readOutlineHeadings(
  reader: DocxSidebarReader,
  expectVersion: string
): DocxOutlineHeading[] | null {
  if (reader.version() !== expectVersion) return null;
  const projection = createYrsSidebarProjection(reader);
  const levels = new Map(
    reader.headings('body').map((entry) => [entry.paraId, entry.heading.outlineLevel])
  );
  const headings: DocxOutlineHeading[] = [];
  for (const paragraph of reader.paragraphs('body')) {
    const level = levels.get(paragraph.paraId);
    const text = paragraph.text.trim();
    if (level == null || !text) continue;
    const position = projection.locToDisplayPoint({
      story: 'body', paraId: paragraph.paraId, offset: 0,
    })?.position;
    if (position == null) continue;
    headings.push({ story: 'body', paraId: paragraph.paraId, text, level, position });
  }
  return headings;
}
