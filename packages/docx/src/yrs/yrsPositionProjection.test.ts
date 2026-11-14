import { describe, expect, test } from 'bun:test';
import type { PointPosition } from '../plugin-api';
import {
  createYrsInputPositionMap,
  displayPositionToYrsLoc,
  type YrsSession,
  type YrsStorySegment,
} from './index';
import {
  createYrsPositionProjection,
  projectYrsDisplayPosition,
  YrsPositionProjection,
} from './yrsPositionProjection';

const segments: Record<string, YrsStorySegment[]> = {
  body: [
    { kind: 'text', text: 'before', attributes: {} },
    { kind: 'pilcrow', paraId: 'body:p0', properties: {}, attributes: {} },
    {
      kind: 'embed',
      embedKind: 'table',
      payload: {
        tblPr: {},
        grid: [100],
        rows: [{ cells: [{ story: 'body:t0:r0c0', tcPr: {} }] }],
      },
      attributes: {},
    },
    { kind: 'text', text: 'after', attributes: {} },
    { kind: 'pilcrow', paraId: 'body:p1', properties: {}, attributes: {} },
  ],
  'body:t0:r0c0': [
    { kind: 'text', text: 'cell', attributes: {} },
    { kind: 'pilcrow', paraId: 'cell:p0', properties: {}, attributes: {} },
  ],
};

const session = {
  storySegments: (story: string) => segments[story] ?? [],
} as unknown as YrsSession;

describe('YrsPositionProjection', () => {
  test('maps point hits into table cells in their own body, header, footer or note', () => {
    const cases: Array<[PointPosition, string]> = [
      [{ position: 12, pageIndex: 1, region: 'body' }, 'body'],
      [{ position: 12, pageIndex: 1, region: 'header', rId: 'rIdHeader' }, 'hf:rIdHeader'],
      [{ position: 12, pageIndex: 1, region: 'footer', rId: 'rIdFooter' }, 'hf:rIdFooter'],
      [{ position: 12, pageIndex: 1, region: 'footnote', noteId: 2 }, 'fn:2'],
      [{ position: 12, pageIndex: 1, region: 'endnote', noteId: 3 }, 'en:3'],
    ];
    const stories: Record<string, YrsStorySegment[]> = {};
    for (const [, root] of cases) {
      stories[root] = segments.body.map((segment) =>
        segment.kind === 'embed'
          ? {
              ...segment,
              payload: {
                tblPr: {},
                grid: [100],
                rows: [{ cells: [{ story: `${root}:t0:r0c0`, tcPr: {} }] }],
              },
            }
          : segment
      );
      stories[`${root}:t0:r0c0`] = segments['body:t0:r0c0'];
    }
    const replica = {
      storyIds: () => Object.keys(stories),
      hasStory: (story: string) => story in stories,
      storySegments: (story: string) => stories[story] ?? [],
    } as unknown as YrsSession;
    const getProjection = (root: string) => createYrsPositionProjection(replica, root);
    for (const [hit, root] of cases) {
      const target = projectYrsDisplayPosition(hit, getProjection)!;
      expect(target.story).toBe(`${root}:t0:r0c0`);
      const map = createYrsInputPositionMap(target.story, [{ paraId: 'cell:p0', length: 4 }]);
      const loc = displayPositionToYrsLoc(map, target.displayPosition);
      expect(loc).toEqual({ story: `${root}:t0:r0c0`, paraId: 'cell:p0', offset: 0 });
      expect(getProjection(root)!.positionForLoc(loc!)).toBe(hit.position);
    }
    expect(projectYrsDisplayPosition(12, getProjection)?.story).toBe('body:t0:r0c0');
    expect(
      projectYrsDisplayPosition({ position: 1, pageIndex: 0, region: 'header' }, getProjection)
    ).toBeNull();
    expect(
      projectYrsDisplayPosition({ position: 1, pageIndex: 0, region: 'footnote' }, getProjection)
    ).toBeNull();
    expect(
      projectYrsDisplayPosition(
        { position: 1, pageIndex: 0, region: 'endnote', noteId: 99 },
        getProjection
      )
    ).toBeNull();
    expect(projectYrsDisplayPosition(NaN, getProjection)).toBeNull();
    expect(projectYrsDisplayPosition(10000, getProjection)).toBeNull();
  });

  test('returns null without reading a missing story', () => {
    let readStory = false;
    const missingStorySession = {
      storyIds: () => ['body'],
      hasStory: (story: string) => story === 'body',
      storySegments: () => {
        readStory = true;
        throw new Error('missing story');
      },
    } as unknown as YrsSession;
    let projection: YrsPositionProjection | null = null;

    expect(() => {
      projection = createYrsPositionProjection(missingStorySession, 'fn:2');
    }).not.toThrow();
    expect(projection).toBeNull();
    expect(readStory).toBe(false);
  });

  test('maps post-table positions back to the root story input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    // `paragraph_spans` measures a paragraph from the previous pilcrow, so the
    // table embed ahead of `body:p1` counts as one of its six units.
    const map = createYrsInputPositionMap('body', [
      { paraId: 'body:p0', length: 6 },
      { paraId: 'body:p1', length: 6 },
    ]);
    const target = projection.targetAt(23);
    const loc = displayPositionToYrsLoc(map, target.displayPosition);

    expect(target).toEqual({ story: 'body', displayPosition: 12 });
    expect(loc).toEqual({
      story: 'body',
      paraId: 'body:p1',
      offset: 3,
    });
    expect(projection.positionForLoc(loc!)).toBe(23);
  });

  test('keeps table cell positions scoped to the cell input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const map = createYrsInputPositionMap('body:t0:r0c0', [{ paraId: 'cell:p0', length: 4 }]);
    const target = projection.targetAt(12);

    expect(target).toMatchObject({
      story: 'body:t0:r0c0',
      displayPosition: 1,
    });
    const loc = displayPositionToYrsLoc(map, target.displayPosition);
    expect(loc).toEqual({
      story: 'body:t0:r0c0',
      paraId: 'cell:p0',
      offset: 0,
    });
    expect(projection.positionForLoc(loc!)).toBe(12);
  });

  test('resolves every position to the deepest story holding it, as a scan of all stories does', () => {
    const text = (value: string): YrsStorySegment => ({ kind: 'text', text: value, attributes: {} });
    const pilcrow = (paraId: string): YrsStorySegment => ({ kind: 'pilcrow', paraId, properties: {}, attributes: {} });
    const table = (cells: string[][]): YrsStorySegment => ({
      kind: 'embed',
      embedKind: 'table',
      payload: { tblPr: {}, grid: [100, 100], rows: cells.map((row) => ({ cells: row.map((story) => ({ story, tcPr: {} })) })) },
      attributes: {},
    });
    const nested: Record<string, YrsStorySegment[]> = {
      body: [
        text('before'), pilcrow('p0'),
        table([['c00', 'c01'], ['c10', 'c11']]),
        { kind: 'embed', embedKind: 'blockSdt', payload: { story: 'sdt' }, attributes: {} },
        table([['d00']]),
        text('after'), pilcrow('p1'),
      ],
      c00: [text('a'), pilcrow('c00p')],
      // a cell opening with a table, so a story starts where its parent's content does
      c01: [table([['e00', 'e01']]), text('tail'), pilcrow('c01p')],
      c10: [pilcrow('c10p')],
      c11: [text('xy'), pilcrow('c11p0'), text('z'), pilcrow('c11p1')],
      e00: [table([['f00']]), pilcrow('e00p')],
      e01: [text('deep'), pilcrow('e01p')],
      f00: [text('deepest'), pilcrow('f00p')],
      sdt: [text('control'), pilcrow('sdtp'), table([['g00']])],
      g00: [text('g'), pilcrow('g00p')],
      d00: [text('last'), pilcrow('d00p')],
    };
    const projection = new YrsPositionProjection(
      { storySegments: (story: string) => nested[story] ?? [] } as unknown as YrsSession,
      'body'
    );
    type Story = { story: string; contentStart: number; size: number; depth: number };
    const stories = [...(projection as unknown as { stories: Map<string, Story> }).stories.values()];
    expect(stories.length).toBe(Object.keys(nested).length);
    for (let position = -1; position <= projection.size + 1; position += 1) {
      let expected = stories.find((story) => story.story === 'body')!;
      for (const story of stories) {
        if (position >= story.contentStart && position <= story.contentStart + story.size && story.depth > expected.depth) {
          expected = story;
        }
      }
      expect([position, projection.targetAt(position).story]).toEqual([position, expected.story]);
    }
  });
});
