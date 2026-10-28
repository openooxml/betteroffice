import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { storiesDocx } from './__fixtures__/storiesDocx';
import type { DocxDrawing } from '../types/content/run';
import type { DisplayList } from '../layout/render/displayList';
import { findDrawingRectsByDocPos } from '../layout/render/displayListImages';
import { computeAnchorTargetGeometry, resolveAnchorTarget, type AnchorReader, type ProposalGeometryReader } from './proposalGeometry';
import type { DocxTextAtom } from './edits';
import type { DocxExportInline } from './structuredExport';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

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

  test('unions chart rects and shapes when a display list carries no chart frame', () => {
    const chart = { label: 'Chart' };
    const pie = {
      pages: [
        {
          pageIndex: 0,
          width: 100,
          height: 200,
          primitives: [
            { kind: 'rect', x: 10, y: 10, w: 80, h: 60, fill: '#fff', docStart: 5, chart },
            { kind: 'shape', x: 30, y: 20, w: 40, h: 40, geometryPath: [], docStart: 5, chart },
          ],
        },
      ],
    } as unknown as DisplayList;
    expect(findDrawingRectsByDocPos(pie, 5)).toEqual([
      { pageIndex: 0, x: 10, y: 10, width: 80, height: 60 },
    ]);
  });
});

test('readStories carries control drawings in revision view order', async () => {
  const session = await createYrsSession({ clientId: 76203 });
  try {
    session.openDocx(storiesDocx(), true);
    const image = (id: string, attrs = {}) => ({
      kind: 'image',
      payload: { docPr: { ...DRAWING, docPrId: id }, title: 'Live title', alt: 'Live alt', wrapType: 'square' },
      attrs,
    });
    session.applyRawOps('body', [{ op: 'insertEmbed', index: 0, kind: 'sdt', payload: {
      content: [
        image('1'),
        { kind: 'sdt', payload: { content: [image('2', { hyperlink: { href: '#logo' } })] }, attrs: {} },
        image('3', { del: { id: 3, author: 'A' } }),
        image('4', { ins: { id: 4, author: 'A' } }),
      ],
    } }]);
    for (const view of ['accepted', 'original'] as const) {
      const read = session.readStories({ stories: ['body'], view });
      if (!read.ok) throw new Error(read.failure.message);
      const story = read.stories[0]!;
      if (!story.ok) throw new Error(story.failure.message);
      const atom = story.paragraphs[0]!.atoms[0]!;
      expect(atom.kind).toBe('contentControl');
      expect(atom.drawings?.map(({ docPrId }) => docPrId)).toEqual(view === 'accepted' ? ['1', '2', '4'] : ['1', '2', '3']);
      expect(atom.drawings?.[0]).toEqual({ ...DRAWING, docPrId: '1', title: 'Live title', descr: 'Live alt', placement: 'anchor' });
    }
  } finally {
    session.destroy();
  }
});

test('indexed drawings reuse expanded control positions including hidden and nested content', () => {
  const nested = { kind: 'sdt', payload: { content: [
    { kind: 'text', text: 'xy', attrs: {} },
    { kind: 'image', payload: { docPr: DRAWING }, attrs: { del: { id: 1 } } },
    { kind: 'sdt', payload: { content: [{ kind: 'image', payload: { docPr: DRAWING }, attrs: {} }] }, attrs: {} },
    { kind: 'image', payload: { docPr: DRAWING }, attrs: {} },
  ] }, attributes: {} };
  const base = reader();
  const controlReader = { ...base, version: () => 'v1', storySegments: () => [
    { kind: 'text', text: 'ab', attributes: {} },
    { ...nested, kind: 'embed', embedKind: 'sdt' },
    { kind: 'pilcrow', paraId: 'p', properties: {}, attributes: {} },
  ] } as unknown as ProposalGeometryReader;
  const resolved = computeAnchorTargetGeometry(controlReader, [
    { kind: 'drawing', paragraph: PARAGRAPH, offset: 2, index: 0 },
    { kind: 'drawing', paragraph: PARAGRAPH, offset: 2, index: 1 },
  ]);
  expect(resolved).toMatchObject([
    { ok: true, drawing: true, ranges: [{ from: 8, to: 9 }] },
    { ok: true, drawing: true, ranges: [{ from: 10, to: 11 }] },
  ]);
});
