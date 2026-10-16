import { expect, test } from 'bun:test';
import type { YrsRevisionInfo } from '../../yrs';
import { extractTrackedChangesFromYrs } from './yrsTrackedChanges';

const projection = {
  locToDisplayPoint: ({ story, offset }: { story: string; offset: number }) => ({
    story,
    position: offset,
  }),
  storyOffsetToDisplayPoint: (story: string, offset: number) => ({
    story,
    position: offset,
  }),
};

const revision = (kind: YrsRevisionInfo['kind'], start: number): YrsRevisionInfo => ({
  revisionId: 'B',
  author: 'Ann',
  date: '2026-09-30T00:00:00Z',
  kind,
  story: 'body',
  preview: '',
  range: {
    story: 'body',
    start: { paraId: 'p', offset: start },
    end: { paraId: 'p', offset: start + 1 },
  },
});

test('a paragraph mark kept before an inserted table is listed as the table insertion', () => {
  const { entries } = extractTrackedChangesFromYrs(
    [revision('pPrIns', 3), revision('tableIns', 4)],
    projection
  );
  expect(entries).toHaveLength(1);
  expect(entries[0]).toMatchObject({ type: 'tableInserted', from: 3, to: 5 });
});
