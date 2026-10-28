import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { DocxSessionParagraphAnchor, YrsSession, YrsStorySegment } from '@betteroffice/docx/yrs';
import { createPluginGeometry } from './geometry';
import type { DocxAnchorGeometryResult, DocxGeometryTarget } from './types';

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PAGE = { width: 100, height: 200 };

function rectAt(left: number, top: number, width: number, height: number): DOMRect {
  return {
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    x: left,
    y: top,
    toJSON() {},
  } as DOMRect;
}

const BODY_ANCHOR: DocxSessionParagraphAnchor = {
  kind: 'session',
  sessionId: 'session',
  story: 'body',
  paraId: 'p',
};
const FOOTER_ANCHOR: DocxSessionParagraphAnchor = {
  kind: 'session',
  sessionId: 'session',
  story: 'hf:rFoot',
  paraId: 'fp',
};
const HEADER_ANCHOR: DocxSessionParagraphAnchor = {
  kind: 'session',
  sessionId: 'session',
  story: 'hf:rHead',
  paraId: 'hp',
};

function segments(docPr: Record<string, unknown>): YrsStorySegment[] {
  return [
    { kind: 'text', text: 'ab', attributes: {} },
    { kind: 'embed', embedKind: 'image', payload: { docPr }, attributes: {} },
    { kind: 'pilcrow', paraId: 'p', properties: {}, attributes: {} },
  ];
}

function harness({
  docPr = { name: 'N-inline', placement: 'inline', type: 'picture' },
  painted = true,
  complete = true,
  unbuilt = false,
  control = false,
  cachedField = false,
}: {
  docPr?: Record<string, unknown>;
  painted?: boolean;
  complete?: boolean;
  unbuilt?: boolean;
  control?: boolean;
  cachedField?: boolean;
} = {}) {
  const pages = document.createElement('div');
  const layer = document.createElement('div');
  pages.getBoundingClientRect = () => rectAt(0, 0, 100, 400);
  layer.getBoundingClientRect = () => rectAt(0, 0, 100, 400);
  const canvases = [0, 1].map((index) => {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    canvas.getBoundingClientRect = () => rectAt(0, index * 200, PAGE.width, PAGE.height);
    pages.appendChild(canvas);
    return canvas;
  });
  void canvases;
  const frame = { pageIndex: 0, x: 30, y: 40, width: 50, height: 20 };
  const drawingPosition = control ? (cachedField ? 7 : 6) : 3;
  const headerFrame = { x: 5, y: 6, w: 70, h: 10, docStart: drawingPosition };
  const source = {
    pageSize: () => PAGE,
    pageCount: () => 2,
    rangeRects: (from: number, to: number) => [
      { pageIndex: 0, x: 10, y: 20, width: to - from, height: 12 } as DisplayListRect,
    ],
    caretRect: () => null,
    anchorRect: () => ({ pageIndex: 0, x: 10, y: 20, width: 0, height: 12 }),
    hitTestRegions: () => null,
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, ...PAGE }),
    displayList: {
      pages: [0, 1].map((pageIndex) => ({
        pageIndex,
        ...PAGE,
        ...(unbuilt && pageIndex === 0 ? { unbuilt: true, positionSpan: [0, 10] } : {}),
        primitives:
          painted && !unbuilt && pageIndex === 0
            ? [
                {
                  kind: 'image',
                  relId: 'r1',
                  x: frame.x,
                  y: frame.y,
                  w: frame.width,
                  h: frame.height,
                  docStart: drawingPosition,
                },
              ]
            : [],
        footer: {
          rId: 'rFoot',
          kind: 'footer',
          y: 170,
          height: 30,
          primitives: painted ? [{ kind: 'image', relId: 'r3', ...headerFrame }] : [],
        },
        header: {
          rId: 'rHead',
          kind: 'header',
          y: 0,
          height: 30,
          primitives: painted ? [{ kind: 'image', relId: 'r2', ...headerFrame }] : [],
        },
      })),
    },
  } as unknown as DisplayListQueries;
  const storySegments = segments(docPr);
  if (control) storySegments[1] = {
    kind: 'embed', embedKind: 'sdt', attributes: {},
    payload: { content: [
      { kind: 'text', text: 'xy', attrs: {} },
      ...(cachedField ? [{ kind: 'field', payload: { fieldData: JSON.stringify({
        type: 'simpleField', content: [{ type: 'run', content: [{ type: 'drawing', drawing: docPr }] }],
      }) }, attrs: {} }] : []),
      { kind: 'image', payload: { docPr }, attrs: {} },
    ] },
  };
  const stories: Record<string, YrsStorySegment[]> = {
    body: storySegments,
    'hf:rFoot': storySegments.map((segment) =>
      segment.kind === 'pilcrow' ? { ...segment, paraId: 'fp' } : segment
    ),
    'hf:rHead': storySegments.map((segment) =>
      segment.kind === 'pilcrow' ? { ...segment, paraId: 'hp' } : segment
    ),
  };
  const session = {
    version: () => 'v1',
    hasStory: (story: string) => story in stories,
    getProposals: () => ({ version: 'v1', previewVersion: 0, proposals: [] }),
    resolveParagraphAnchor: (anchor: DocxSessionParagraphAnchor) => ({
      status: 'found',
      anchor,
    }),
    paragraphSpans: (story: string) => [
      { paraId: story === 'body' ? 'p' : story === 'hf:rFoot' ? 'fp' : 'hp', length: 3 },
    ],
    storySegments: (story: string) => stories[story] ?? [],
    findText: () => ({ ok: true, version: 'v1', matches: [], truncated: false }),
    listRevisions: () => [],
  } as unknown as YrsSession;
  const editor = {
    hasPendingInput: () => false,
    yrsLocToDisplayPosition: (loc: { offset: number }): number | null => loc.offset + 1,
  };
  const dom = createRenderedDomContext(pages, 1, {
    displayListQueries: source,
    projector: createCanvasHostProjector(pages, source, 1),
  });
  const layout = {
    id: 'layout',
    version: 'v1',
    previewVersion: 0,
    reservationsVersion: null,
    zoom: 1,
    pageCount: 2,
  };
  const geometry = createPluginGeometry(
    layout,
    dom,
    layer,
    () => true,
    () => null,
    source,
    () => ({ session, editor, presented: true }),
    () => false,
    undefined,
    () => complete
  );
  return { geometry, frame };
}

function anchored(result: DocxAnchorGeometryResult) {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

describe('drawing geometry targets', () => {
  test('an inline picture answers with its image primitive rect, not its paragraph text', () => {
    const { geometry, frame } = harness();
    const target: DocxGeometryTarget = { kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2 };
    const result = anchored(geometry.getAnchorGeometry(target));
    expect(result.rects).toEqual([{ ...frame, pageIndex: 0 }]);
    expect(result.anchor).toEqual(result.rects[0]);
    expect(result.pageRect).toMatchObject({ width: PAGE.width, height: PAGE.height });
  });

  test('an anchored picture answers with its floating frame', () => {
    const { geometry, frame } = harness({
      docPr: { name: 'N-anchor', placement: 'anchor', type: 'picture' },
    });
    const result = anchored(
      geometry.getAnchorGeometry({ kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2 })
    );
    expect(result.rects).toEqual([{ ...frame, pageIndex: 0 }]);
  });

  test('a header drawing answers on every page that paints it', () => {
    const { geometry } = harness();
    const result = anchored(
      geometry.getAnchorGeometry({ kind: 'drawing', paragraph: HEADER_ANCHOR, offset: 2 })
    );
    expect(result.rects.map((rect) => rect.pageIndex)).toEqual([0, 1]);
    expect(result.rects[0]).toMatchObject({ x: 5, width: 70, height: 10 });
    expect(result.anchor).toEqual(result.rects[0]);
  });

  test('an offset that is not on a drawing atom is refused', () => {
    const { geometry } = harness();
    expect(
      geometry.getAnchorGeometry({ kind: 'drawing', paragraph: BODY_ANCHOR, offset: 0 })
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
  });

  test('a hidden drawing is missing once the layout is complete', () => {
    const { geometry } = harness({
      docPr: { name: 'N-hidden', hidden: true, placement: 'inline', type: 'picture' },
      painted: false,
      complete: true,
    });
    expect(
      geometry.getAnchorGeometry({ kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2 })
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
  });

  test('readAnchorGeometries answers drawing targets in order', async () => {
    const { geometry, frame } = harness();
    const results = await geometry.readAnchorGeometries([
      { kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2 },
      { kind: 'drawing', paragraph: BODY_ANCHOR, offset: 1 },
    ]);
    expect(results[0]).toMatchObject({ ok: true, rects: [{ ...frame, pageIndex: 0 }] });
    expect(results[1]).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
  });

  test('a drawing on a page not built yet waits instead of going missing', () => {
    const { geometry } = harness({ unbuilt: true, complete: true });
    expect(
      geometry.getAnchorGeometry({ kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2 })
    ).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
  });
});

test('an indexed control drawing answers with its inner frame', () => {
  const { geometry, frame } = harness({ control: true });
  const result = anchored(geometry.getAnchorGeometry({
    kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2, index: 0,
  }));
  expect(result.rects).toEqual([{ ...frame, pageIndex: 0 }]);
});

test('invalid control drawing indices are missing targets', () => {
  for (const control of [false, true]) {
    const { geometry } = harness({ control });
    for (const index of control ? [1, -1, 0.5] : [0, 1]) {
      expect(geometry.getAnchorGeometry({
        kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2, index,
      })).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    }
    if (control) expect(geometry.getAnchorGeometry({
      kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2,
    })).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
  }
});

test('a footer control drawing answers on every page that paints it', () => {
  const { geometry } = harness({ control: true });
  const result = anchored(geometry.getAnchorGeometry({
    kind: 'drawing', paragraph: FOOTER_ANCHOR, offset: 2, index: 0,
  }));
  expect(result.rects.map((rect) => rect.pageIndex)).toEqual([0, 1]);
  expect(result.rects[0]).toMatchObject({ x: 5, y: 6, width: 70, height: 10 });
});

test('a cached field drawing is missing while a later control drawing keeps its index', () => {
  const { geometry, frame } = harness({ control: true, cachedField: true });
  expect(geometry.getAnchorGeometry({ kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2, index: 0 })).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
  expect(anchored(geometry.getAnchorGeometry({ kind: 'drawing', paragraph: BODY_ANCHOR, offset: 2, index: 1 })).rects).toEqual([{ ...frame, pageIndex: 0 }]);
});
