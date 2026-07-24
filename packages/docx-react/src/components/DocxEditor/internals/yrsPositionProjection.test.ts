import { describe, expect, test } from 'bun:test';
import {
  createYrsInputPositionMap,
  displayPositionToYrsLoc,
  type YrsSession,
  type YrsStorySegment,
} from '@betteroffice/docx/yrs';
import { YrsPositionProjection } from './yrsPositionProjection';

const segments: Record<string, YrsStorySegment[]> = {
  body: [
    { kind: 'text', text: 'before', attributes: {} },
    { kind: 'pilcrow', paraId: 'body:p0', properties: {}, attributes: {} },
    {
      kind: 'embed',
      embedKind: 'table',
      payload: {
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
  test('maps post-table positions back to the root story input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const map = createYrsInputPositionMap('body', [
      { paraId: 'body:p0', length: 6 },
      { paraId: 'body:p1', length: 5 },
    ]);
    const target = projection.targetAt(23);
    const loc = displayPositionToYrsLoc(map, target.displayPosition);

    expect(target).toEqual({ story: 'body', displayPosition: 11 });
    expect(loc).toEqual({
      story: 'body',
      paraId: 'body:p1',
      offset: 2,
    });
    expect(projection.positionForLoc(loc!)).toBe(23);
  });

  test('keeps table cell positions scoped to the cell input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const map = createYrsInputPositionMap('body:t0:r0c0', [
      { paraId: 'cell:p0', length: 4 },
    ]);
    const target = projection.targetAt(12);

    expect(target).toMatchObject({
      story: 'body:t0:r0c0',
      displayPosition: 1,
    });
    expect(displayPositionToYrsLoc(map, target.displayPosition)).toEqual({
      story: 'body:t0:r0c0',
      paraId: 'cell:p0',
      offset: 0,
    });
  });
});
