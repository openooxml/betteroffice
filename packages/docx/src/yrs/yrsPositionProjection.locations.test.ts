import { describe, expect, test } from 'bun:test';
import type { YrsStorySegment } from './index';
import {
  createYrsLocProjectionFromOutline,
  YrsPositionProjection,
  yrsLocToProjectedDisplayPosition,
} from './yrsPositionProjection';

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

const session = { storySegments: (story: string) => segments[story] ?? [] };

describe('yrsLocToProjectedDisplayPosition', () => {
  const reader = {
    hasStory: (story: string) => story in segments,
    paragraphSpans: (story: string) =>
      story === 'body'
        ? [{ paraId: 'body:p0', length: 6 }, { paraId: 'body:p1', length: 6 }]
        : [{ paraId: 'cell:p0', length: 4 }],
  };

  test('uses the body projection for nested stories and adjusts leading block offsets', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const roots: string[] = [];
    const projectionFor = (root: string) => {
      roots.push(root);
      return projection;
    };
    expect(
      yrsLocToProjectedDisplayPosition(
        reader,
        projectionFor,
        { story: 'body:t0:r0c0', paraId: 'cell:p0', offset: 0 },
        'hf:rIdHeader'
      )
    ).toBe(12);
    expect(
      yrsLocToProjectedDisplayPosition(reader, projectionFor, {
        story: 'body',
        paraId: 'body:p1',
        offset: 3,
      })
    ).toBe(23);
    expect(roots).toEqual(['body', 'body']);
  });

  test('keeps the root-local fallback and refuses an unmapped nested story', () => {
    expect(
      yrsLocToProjectedDisplayPosition(reader, () => null, {
        story: 'body',
        paraId: 'body:p1',
        offset: 3,
      })
    ).toBe(12);
    expect(
      yrsLocToProjectedDisplayPosition(reader, () => null, {
        story: 'body:t0:r0c0',
        paraId: 'cell:p0',
        offset: 0,
      })
    ).toBeNull();
    expect(
      yrsLocToProjectedDisplayPosition(reader, () => null, {
        story: 'body',
        paraId: 'missing',
        offset: 0,
      })
    ).toBe(0);
  });

  test('selects the active non-body root and respects the editor input-map gate', () => {
    const root = 'hf:rIdHeader';
    const rootReader = {
      hasStory: () => true,
      paragraphSpans: () => [{ paraId: 'p', length: 4 }],
    };
    const roots: string[] = [];
    const projectionFor = (story: string) => {
      roots.push(story);
      return null;
    };
    expect(
      yrsLocToProjectedDisplayPosition(
        rootReader,
        projectionFor,
        { story: root, paraId: 'p', offset: 2 },
        root
      )
    ).toBe(3);
    expect(roots).toEqual([root]);
    expect(
      yrsLocToProjectedDisplayPosition(
        rootReader,
        projectionFor,
        { story: root, paraId: 'p', offset: 2 },
        root,
        () => null
      )
    ).toBeNull();
    expect(roots).toEqual([root]);
    expect(
      yrsLocToProjectedDisplayPosition(
        reader,
        () => {
          throw new Error('missing input map must not read a projection');
        },
        { story: 'absent', paraId: 'p', offset: 0 }
      )
    ).toBeNull();
  });

  test('the flat outline preserves first paragraph lookup, clamps and wrapper fallbacks', () => {
    const old = new YrsPositionProjection(session, 'body');
    const fast = createYrsLocProjectionFromOutline({
      body: { contentStart: 0, size: 27, paragraphs: [
        { paraId: 'body:p0', displayStart: 0, length: 6, leading: 0 },
        { paraId: 'body:p1', displayStart: 20, length: 5, leading: 1 },
        { paraId: 'body:p1', displayStart: 100, length: 20, leading: 0 },
      ] },
      'body:t0:r0c0': { contentStart: 11, size: 6, paragraphs: [
        { paraId: 'cell:p0', displayStart: 0, length: 4, leading: 0 },
      ] },
    });
    for (const [story, paraId] of [
      ['body', 'body:p0'], ['body', 'body:p1'], ['body:t0:r0c0', 'cell:p0'],
      ['body', 'missing'], ['body:t0:r0c0', 'missing'], ['absent', 'missing'],
    ]) {
      for (const offset of [-Infinity, -1, 0, 1, 3, 6, 100, Infinity, NaN]) {
        const loc = { story, paraId, offset };
        expect(fast.positionForLoc(loc)).toBe(old.positionForLoc(loc));
        expect(yrsLocToProjectedDisplayPosition(reader, () => fast, loc)).toBe(
          yrsLocToProjectedDisplayPosition(reader, () => old, loc)
        );
      }
    }
    expect(yrsLocToProjectedDisplayPosition(reader, () => fast,
      { story: 'body', paraId: 'body:p0', offset: 0 }, 'body', () => null
    )).toBeNull();
  });
});
