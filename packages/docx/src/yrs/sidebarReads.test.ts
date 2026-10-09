import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { readSidebar, readOutlineHeadings } from './sidebarReads';
import { sidebarDocx } from './__fixtures__/sidebarDocx';
import { createYrsSidebarProjection, yrsIdToNumericId } from '../layout/render/yrsSidebarProjection';
import { extractTrackedChangesFromYrs } from '../layout/render/yrsTrackedChanges';
import { anchorPositionsFromPoints, computeAnchorPositionsFromYrs } from '../layout/render/displayListAnchors';
import type { DisplayListQueries } from '../layout/render/displayListQueries';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

test('sidebar reads preserve session anchor order and tracked changes', async () => {
  const session = await createYrsSession();
  try {
    session.openDocx(sidebarDocx(), true);
    const version = session.version();
    const projection = createYrsSidebarProjection(session);
    const value = readSidebar(session, ['7', 'missing'], version)!;
    expect(value.comments).toEqual([
      { id: '7', anchors: session.resolveComment('7').map((anchor) => ({
        start: projection.storyOffsetToDisplayPoint(anchor.story, anchor.start),
        end: projection.storyOffsetToDisplayPoint(anchor.story, anchor.end),
      })) },
      { id: 'missing', anchors: [] },
    ]);
    expect(value.comments[0]!.anchors.length).toBeGreaterThan(0);
    const revisions = session.listRevisions();
    expect(revisions.map(({ kind }) => kind).sort()).toEqual(['deletion', 'insertion']);
    expect(value.revisions).toEqual(revisions.map((revision) => ({
      key: `revision-${yrsIdToNumericId(revision.revisionId)}`,
      start: projection.locToDisplayPoint({ story: revision.story, ...revision.range.start }),
    })));
    const tracked = extractTrackedChangesFromYrs(revisions, projection);
    expect(value.trackedChanges).toEqual({ entries: tracked.entries, commentToRevision: [...tracked.commentToRevision] });
    const queries = {
      pageCount: () => 1,
      pageSize: () => ({ width: 600, height: 800 }),
      anchorRect: (position: number) => ({ pageIndex: 0, x: 0, y: position, width: 1, height: 10 }),
    } as unknown as DisplayListQueries;
    const points = [
      ...value.comments.flatMap(({ id, anchors }) => anchors.map(({ start }) => [`comment-${id}`, start] as const)),
      ...value.revisions.map(({ key, start }) => [key, start] as const),
    ];
    expect(anchorPositionsFromPoints(points, queries)).toEqual(
      computeAnchorPositionsFromYrs(session, ['7', 'missing'], revisions, projection, queries)
    );
    expect(readSidebar(session, ['7'], 'stale')).toBeNull();
    const throwing = Object.create(session) as typeof session;
    throwing.resolveComment = () => { throw new Error('missing comment'); };
    expect(readSidebar(throwing, ['7'], version)!.comments).toEqual([{ id: '7', anchors: [] }]);
  } finally {
    session.destroy();
  }
});

test('outline headings have trimmed text and the editor display positions', async () => {
  const session = await createYrsSession();
  try {
    session.openDocx(sidebarDocx(), true);
    const headings = readOutlineHeadings(session, session.version())!;
    expect(headings.map(({ text, level }) => [text, level])).toEqual([
      ['First heading', 0], ['Second heading', 1],
    ]);
    const levels = new Set(session.headings('body').map(({ paraId }) => paraId));
    expect(headings.map(({ paraId, text }) => [paraId, text])).toEqual(session.paragraphs('body')
      .filter(({ paraId, text }) => levels.has(paraId) && text.trim())
      .map(({ paraId, text }) => [paraId, text.trim()]));
    const projection = createYrsSidebarProjection(session);
    const editorProjection = createYrsPositionProjection(session, 'body');
    for (const heading of headings) {
      const loc = { story: heading.story, paraId: heading.paraId, offset: 0 };
      expect(heading.position).toBe(projection.locToDisplayPoint(loc)!.position);
      expect<number | null>(heading.position).toBe(editorProjection!.positionForLoc(loc));
    }
    expect(readOutlineHeadings(session, 'stale')).toBeNull();
  } finally {
    session.destroy();
  }
});
