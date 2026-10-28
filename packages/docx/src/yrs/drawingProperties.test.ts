import { describe, expect, test } from 'bun:test';
import type { DocxDrawing } from '../types/content/run';
import type { DisplayList } from '../layout/render/displayList';
import { findDrawingRectsByDocPos } from '../layout/render/displayListImages';
import { resolveAnchorTarget, type AnchorReader } from './proposalGeometry';
import type { DocxTextAtom } from './edits';
import type { DocxExportInline } from './structuredExport';

const DRAWING: DocxDrawing = {
  docPrId: '7',
  name: 'N-inline',
  title: 'T-inline',
  descr: 'D-inline',
  hidden: false,
  decorative: true,
  placement: 'inline',
  type: 'picture',
};

const PARAGRAPH = {
  kind: 'session',
  sessionId: 'session',
  story: 'body',
  paraId: 'p',
} as const;

function reader(): AnchorReader {
  return {
    resolveParagraphAnchor: () => ({ status: 'found', anchor: PARAGRAPH }),
    hasStory: () => true,
    paragraphSpans: () => [{ paraId: 'p', length: 3 }],
    storySegments: () => [
      { kind: 'text', text: 'ab', attributes: {} },
      { kind: 'embed', embedKind: 'image', payload: { docPr: DRAWING }, attributes: {} },
      { kind: 'pilcrow', paraId: 'p', properties: {}, attributes: {} },
    ],
    findText: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    listRevisions: () => [],
  } as unknown as AnchorReader;
}

describe('drawing properties plumbing', () => {
  test('atoms and structured inlines carry the shared drawing shape', () => {
    const atom: DocxTextAtom = { offset: 2, kind: 'image', drawing: DRAWING };
    expect(JSON.parse(JSON.stringify(atom)).drawing).toEqual(DRAWING);
    const inline: DocxExportInline = {
      id: 'i',
      anchor: { kind: 'range', story: 'body', start: { paraId: 'p', offset: 2 }, end: { paraId: 'p', offset: 3 }, view: 'accepted' } as never,
      marks: null,
      link: null,
      revisions: [],
      kind: 'image',
      altText: 'D-inline',
      relationshipId: 'rId1',
      part: 'word/media/image.png',
      externalTarget: null,
      drawing: DRAWING,
    };
    expect(inline.kind === 'image' && inline.drawing).toEqual(DRAWING);
    const unsupported: DocxExportInline = {
      id: 'u',
      anchor: inline.anchor,
      marks: null,
      link: null,
      revisions: [],
      kind: 'unsupported',
      element: 'wps:wsp',
      altText: null,
      drawing: DRAWING,
    };
    expect(unsupported.kind === 'unsupported' && unsupported.drawing?.type).toBe('picture');
  });

  test('a drawing target resolves the atom at its accepted-view offset', () => {
    const resolved = resolveAnchorTarget(reader(), { kind: 'drawing', paragraph: PARAGRAPH, offset: 2 }, 'v1');
    expect(resolved).toMatchObject({
      ok: true,
      drawing: true,
      ranges: [{ start: { story: 'body', paraId: 'p', offset: 2 }, end: { offset: 3 } }],
    });
  });

  test('a drawing target refuses offsets that are not on a drawing atom', () => {
    for (const offset of [0, 1, 3, -1, 1.5]) {
      expect(
        resolveAnchorTarget(reader(), { kind: 'drawing', paragraph: PARAGRAPH, offset }, 'v1')
      ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    }
  });
});

describe('findDrawingRectsByDocPos', () => {
  const list = {
    pages: [
      {
        pageIndex: 0,
        width: 100,
        height: 200,
        primitives: [
          { kind: 'image', relId: 'r1', x: 30, y: 40, w: 50, h: 20, docStart: 3 },
          { kind: 'text', text: 'ab', x: 10, y: 40, docStart: 1 },
        ],
      },
      {
        pageIndex: 1,
        width: 100,
        height: 200,
        unbuilt: true,
        primitives: [{ kind: 'image', relId: 'r1', x: 1, y: 2, w: 3, h: 4, docStart: 3 }],
      },
      {
        pageIndex: 2,
        width: 100,
        height: 200,
        primitives: [
          { kind: 'shape', x: 10, y: 10, w: 20, h: 10, geometryPath: [], docStart: 7 },
          { kind: 'shape', x: 30, y: 15, w: 20, h: 10, geometryPath: [], docStart: 7 },
        ],
        header: {
          rId: 'rHead',
          kind: 'header',
          y: 0,
          height: 30,
          primitives: [{ kind: 'image', relId: 'r2', x: 5, y: 6, w: 70, h: 10, docStart: 3 }],
        },
      },
    ],
  } as unknown as DisplayList;

  test('returns the image frame, not the text around it, and skips unbuilt pages', () => {
    expect(findDrawingRectsByDocPos(list, 3)).toEqual([
      { pageIndex: 0, x: 30, y: 40, width: 50, height: 20 },
    ]);
  });

  test('unions the primitives of a group or chart at one atom position', () => {
    expect(findDrawingRectsByDocPos(list, 7)).toEqual([
      { pageIndex: 2, x: 10, y: 10, width: 40, height: 15 },
    ]);
  });

  test('finds header frames in the band that paints them', () => {
    expect(findDrawingRectsByDocPos(list, 3, 'header', 'rHead')).toEqual([
      { pageIndex: 2, x: 5, y: 6, width: 70, height: 10 },
    ]);
    expect(findDrawingRectsByDocPos(list, 3, 'header', 'other')).toEqual([]);
  });

  test('answers a chart its placed frame, not the bounds of what it paints', () => {
    const chart = { label: 'Chart', frame: { x: 20, y: 30, w: 60, h: 40 } };
    const charted = {
      pages: [
        {
          pageIndex: 0,
          width: 100,
          height: 200,
          primitives: [
            { kind: 'text', text: 'Q1', x: 25, y: 60, docStart: 5, chart },
            { kind: 'rect', x: 40, y: 50, w: 5, h: 10, fill: '#000', docStart: 5, chart },
          ],
        },
      ],
    } as unknown as DisplayList;
    expect(findDrawingRectsByDocPos(charted, 5)).toEqual([
      { pageIndex: 0, x: 20, y: 30, width: 60, height: 40 },
    ]);
  });

  test('unions every picture one drawing paints', () => {
    const pictures = {
      pages: [
        {
          pageIndex: 0,
          width: 100,
          height: 200,
          primitives: [
            { kind: 'image', relId: 'r1', x: 10, y: 10, w: 20, h: 10, docStart: 9 },
            { kind: 'image', relId: 'r2', x: 50, y: 40, w: 20, h: 30, docStart: 9 },
          ],
        },
      ],
    } as unknown as DisplayList;
    expect(findDrawingRectsByDocPos(pictures, 9)).toEqual([
      { pageIndex: 0, x: 10, y: 10, width: 60, height: 60 },
    ]);
  });
});
