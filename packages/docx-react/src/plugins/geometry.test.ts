import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, describe, expect, spyOn, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import type { RenderedDomContext } from '@betteroffice/docx/plugin-api';
import {
  proposalRevisionPreview,
  type DocxSessionParagraphAnchor,
  type DocxTextRange,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import {
  revisionPreviewKey,
  stampRevisionPreviewKey,
  stampSourceVersion,
  UNKNOWN_REVISION_PREVIEW_KEY,
} from '../components/DocxEditor/internals/layoutProvenance';
import { createPluginGeometry, pluginLayout, toOverlayRect } from './geometry';
import * as proposalPreview from './proposalPreview';
import type { DocxProposalSnapshot } from './proposalPreview';
import type { DocxAnchorGeometryResult, DocxGeometryTarget } from './types';

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

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

function place(
  element: HTMLElement,
  rect: DOMRect,
  box: Partial<Record<'clientLeft' | 'clientTop' | 'scrollLeft' | 'scrollTop', number>> = {}
) {
  element.getBoundingClientRect = () => rect;
  for (const [key, value] of Object.entries(box)) {
    Object.defineProperty(element, key, { value, configurable: true });
  }
}

const PAGE = { width: 100, height: 200 };
const RANGE: DisplayListRect = {
  pageIndex: 0,
  x: 10,
  y: 20,
  width: 30,
  height: 40,
} as DisplayListRect;

function queries(): DisplayListQueries {
  return {
    pageSize: () => PAGE,
    pageCount: () => 1,
    rangeRects: () => [RANGE],
    caretRect: () => null,
    pageBounds: () => ({ pageIndex: 0, x: 0, y: 0, ...PAGE }),
  } as unknown as DisplayListQueries;
}

describe('plugin overlay geometry', () => {
  test('moves a context rectangle into the layer once, at every zoom', () => {
    for (const zoom of [0.5, 1, 2]) {
      const pages = document.createElement('div');
      const canvas = document.createElement('canvas');
      canvas.dataset.pageIndex = '0';
      pages.appendChild(canvas);
      const layer = document.createElement('div');
      place(pages, rectAt(130, 60, 800, 2000));
      place(canvas, rectAt(180, 84, PAGE.width * zoom, PAGE.height * zoom));
      place(layer, rectAt(20, 10, 900, 2100), {
        clientLeft: 2,
        clientTop: 3,
        scrollLeft: 5,
        scrollTop: 7,
      });
      const source = queries();
      const dom = createRenderedDomContext(pages, zoom, {
        displayListQueries: source,
        projector: createCanvasHostProjector(pages, source, zoom),
      });
      const [rect] = dom.getRectsForRange(0, 1);
      const overlay = toOverlayRect(pages, layer, zoom, rect);
      const originX = -20 - 2 + 5;
      const originY = -10 - 3 + 7;
      expect(overlay.x).toBeCloseTo(180 + originX + RANGE.x * zoom);
      expect(overlay.y).toBeCloseTo(84 + originY + RANGE.y * zoom);
      expect(overlay.width).toBeCloseTo(RANGE.width * zoom);
      expect(overlay.height).toBeCloseTo(RANGE.height * zoom);

      const layout = { id: 'layout', version: 'v', previewVersion: 0, zoom, pageCount: 1 };
      let current = true;
      const geometry = createPluginGeometry(
        layout,
        dom,
        layer,
        () => current,
        () => null,
        source,
        () => null
      );
      expect(geometry.toOverlayRect(rect)).toEqual(overlay);
      current = false;
      expect(geometry.toOverlayRect(rect)).toBeNull();
    }
  });

  test('answers point queries with null through a context without getPositionAtPoint', () => {
    const pages = document.createElement('div');
    const layer = document.createElement('div');
    const dom = createRenderedDomContext(pages);
    const custom: Omit<RenderedDomContext, 'getPositionAtPoint'> = new Proxy(dom, {
      get: (target, key) => (key === 'getPositionAtPoint' ? undefined : Reflect.get(target, key)),
    });
    const layout = { id: 'layout', version: 'v', previewVersion: 0, zoom: 1, pageCount: 1 };
    const geometry = createPluginGeometry(
      layout,
      custom,
      layer,
      () => true,
      () => null,
      queries(),
      () => null
    );
    expect(geometry.getPositionAtPoint(1, 1)).toBeNull();
  });

  test('the mirror fallback answers in container pixels under an ancestor CSS zoom', () => {
    for (const ancestor of [0.713, 1.25]) {
      const parent = document.createElement('div');
      const pages = document.createElement('div');
      const page = document.createElement('div');
      page.className = 'layout-page';
      page.dataset.pageIndex = '0';
      pages.appendChild(page);
      parent.appendChild(pages);
      for (const element of [parent, pages, page]) {
        Object.defineProperty(element, 'currentCSSZoom', { value: ancestor });
      }
      place(parent, rectAt(20, 10, 900 * ancestor, 2100 * ancestor));
      place(pages, rectAt(20 + 40 * ancestor, 10 + 16 * ancestor, 800 * ancestor, 2000 * ancestor));
      const pageLeft = 20 + (40 + 30) * ancestor;
      place(page, rectAt(pageLeft, 10 + (16 + 24) * ancestor, 100 * ancestor, 200 * ancestor));
      const dom = createRenderedDomContext(pages);
      const bounds = dom.getPageBounds(0)!;
      expect(bounds.x).toBeCloseTo(30, 9);
      expect(bounds.y).toBeCloseTo(24, 9);
      expect(bounds.width).toBeCloseTo(100, 9);
      expect(bounds.height).toBeCloseTo(200, 9);
      const offset = dom.getContainerOffset();
      expect(offset.x).toBeCloseTo(40, 9);
      expect(offset.y).toBeCloseTo(16, 9);
    }
  });

  test('measures origins when called, so moved pages move the overlay', () => {
    const pages = document.createElement('div');
    const layer = document.createElement('div');
    place(layer, rectAt(0, 0, 500, 500));
    place(pages, rectAt(40, 0, 400, 400));
    const rect = { x: 10, y: 10, width: 5, height: 5 };
    expect(toOverlayRect(pages, layer, 1, rect).x).toBe(50);
    place(pages, rectAt(-80, 0, 400, 400));
    expect(toOverlayRect(pages, layer, 1, rect).x).toBe(-70);
  });

  test('a layout exists only while its pixels show the requested version', () => {
    const source = queries();
    const preview = { key: '', previewVersion: 4 };
    expect(pluginLayout(source, 'v1', 1, preview)).toBeNull();
    stampSourceVersion(source, 'v1');
    const layout = pluginLayout(source, 'v1', 1.5, preview);
    expect(layout).toMatchObject({ version: 'v1', previewVersion: 4, zoom: 1.5, pageCount: 1 });
    expect(pluginLayout(source, 'v1', 1.5, preview)?.id).toBe(layout!.id);
    expect(pluginLayout(source, 'v2', 1.5, preview)).toBeNull();
    expect(pluginLayout(source, 'v1', 1, { key: 'pending', previewVersion: 5 })).toBeNull();
    const next = queries();
    stampSourceVersion(next, 'v1');
    expect(pluginLayout(next, 'v1', 1.5, preview)?.id).not.toBe(layout!.id);
  });
});

const PARAGRAPH: DocxSessionParagraphAnchor = {
  kind: 'session',
  sessionId: 'session',
  story: 'body',
  paraId: 'p',
};
const TEXT_RANGE: DocxTextRange = {
  story: 'body',
  start: { paraId: 'p', offset: 0 },
  end: { paraId: 'p', offset: 4 },
  view: 'accepted',
};
const PARAGRAPH_TARGET: DocxGeometryTarget = { kind: 'paragraph', paragraph: PARAGRAPH };

/**
 * Pages under a layer at client (20, 10). An `ancestor` CSS zoom scales every client distance
 * from the layer; the elements' own pixels, borders and scroll offsets stay as they are.
 */
function semanticGeometry(zoom = 1, ancestor = 1, { stableSession = false } = {}) {
  const pages = document.createElement('div');
  const layer = document.createElement('div');
  const canvases = [0, 1].map((index) => {
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    pages.appendChild(canvas);
    return canvas;
  });
  if (ancestor !== 1) {
    for (const element of [pages, layer, ...canvases]) {
      Object.defineProperty(element, 'currentCSSZoom', { value: ancestor });
    }
  }
  const client = (x: number, y: number) => ({
    x: 20 + (x - 20) * ancestor,
    y: 10 + (y - 10) * ancestor,
  });
  const movePages = (x: number, y: number) => {
    const at = client(x, y);
    place(pages, rectAt(at.x, at.y, 500 * zoom * ancestor, 1000 * zoom * ancestor));
    canvases.forEach((canvas, index) => {
      place(
        canvas,
        rectAt(
          at.x + 30 * ancestor,
          at.y + (40 + index * 240 * zoom) * ancestor,
          PAGE.width * zoom * ancestor,
          PAGE.height * zoom * ancestor
        )
      );
    });
  };
  movePages(130, 60);
  place(layer, rectAt(20, 10, 900 * ancestor, 2100 * ancestor), {
    clientLeft: 2,
    clientTop: 3,
    scrollLeft: 5,
    scrollTop: 7,
  });
  const calls: [number, number][] = [];
  const source = {
    pageSize: () => PAGE,
    pageCount: () => 2,
    rangeRects: (from: number, to: number) => {
      const previous = calls.at(-1);
      const probe = !!previous && to <= previous[1] && from >= previous[0];
      if (!probe) calls.push([from, to]);
      return [{ ...RANGE, x: 10 + from, width: to - from }];
    },
    caretRect: () => ({ ...RANGE, x: 8, width: 1 }),
    anchorRect: () => ({ ...RANGE, x: 3, width: 5 }),
    hitTestRegions: () => null,
    pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, ...PAGE }),
  } as unknown as DisplayListQueries;
  let snapshot: DocxProposalSnapshot = {
    version: 'v1',
    previewVersion: 0,
    proposals: [
      {
        id: 'proposal',
        state: 'proposed',
        paragraph: PARAGRAPH,
        revisionIds: ['r1', 'r2'],
        changed: true,
      },
    ],
  };
  const session = {
    version: () => 'v1',
    hasStory: () => true,
    getProposals: () => snapshot,
    resolveParagraphAnchor: () => ({ status: 'found', anchor: PARAGRAPH }),
    paragraphSpans: () => [{ paraId: 'p', length: 4 }],
    storySegments: () => [
      { kind: 'text', text: 'aaaa', attributes: {} },
      { kind: 'pilcrow', paraId: 'p', properties: {}, attributes: {} },
    ],
    findText: () => ({
      ok: true,
      version: 'v1',
      truncated: false,
      matches: [0, 1, 2].map((offset) => ({
        text: 'aa',
        range: {
          ...TEXT_RANGE,
          start: { paraId: 'p', offset },
          end: { paraId: 'p', offset: offset + 2 },
        },
      })),
    }),
    listRevisions: () => [
      {
        revisionId: 'r2',
        kind: 'insertion',
        story: 'body',
        range: { start: { paraId: 'p', offset: 2 }, end: { paraId: 'p', offset: 4 } },
      },
      {
        revisionId: 'r1',
        kind: 'deletion',
        story: 'body',
        range: { start: { paraId: 'p', offset: 0 }, end: { paraId: 'p', offset: 1 } },
      },
    ],
  } as unknown as YrsSession;
  const editor = {
    hasPendingInput: () => false,
    yrsLocToDisplayPosition: (loc: { offset: number }): number | null => loc.offset,
  };
  const dom = createRenderedDomContext(pages, zoom, {
    displayListQueries: source,
    projector: createCanvasHostProjector(pages, source, zoom),
  });
  let current = true;
  let available = true;
  let presented = true;
  const layout = { id: 'layout', version: 'v1', previewVersion: 0, zoom, pageCount: 2 };
  const geometry = createPluginGeometry(
    layout,
    dom,
    layer,
    () => current,
    () => null,
    source,
    // Tests swap the session's reads without a version change, which a real document cannot do;
    // a copy per access keeps per-version reads from carrying across those swaps.
    () =>
      available ? { session: stableSession ? session : { ...session }, editor, presented } : null
  );
  return {
    geometry,
    dom,
    source,
    session,
    editor,
    calls,
    layer,
    canvases,
    movePages,
    setCurrent: (value: boolean) => {
      current = value;
    },
    setAvailable: (value: boolean) => {
      available = value;
    },
    setPresented: (value: boolean) => {
      presented = value;
    },
    /** Updates the registry; `rendered` also stamps the queries as showing its preview. */
    setSnapshot: (value: Partial<DocxProposalSnapshot>, rendered = true) => {
      snapshot = { ...snapshot, ...value };
      if (rendered) {
        stampRevisionPreviewKey(source, revisionPreviewKey(proposalRevisionPreview(snapshot)));
      }
    },
  };
}

function anchored(result: DocxAnchorGeometryResult) {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

function refused(result: DocxAnchorGeometryResult, code: string) {
  expect(result).toMatchObject({ ok: false, failure: { code } });
}

/**
 * Lays units 0..3 out one pixel wide from x=10 on one line, left to right or right to left,
 * dropping `hidden` units as a preview does; with all of them hidden the line keeps a 4px mark
 * at position 0. Hit tests answer the caret stop nearest the point.
 */
function textLine(source: DisplayListQueries, hidden: number[] = [], rtl = false) {
  const x = (unit: number) => 10 + (rtl ? 3 - unit : unit);
  const blank = hidden.length === 4;
  source.rangeRects = (from, to) => {
    if (blank) return from <= 0 && to > 0 ? [{ ...RANGE, x: 10, width: 4 }] : [];
    const rects: DisplayListRect[] = [];
    for (let unit = Math.max(0, from); unit < Math.min(4, to); unit += 1) {
      if (!hidden.includes(unit)) rects.push({ ...RANGE, x: x(unit), width: 1 });
    }
    return rects;
  };
  source.hitTestRegions = (_page, at) => {
    if (blank) return { region: 'body', pos: 0, target: 'text' };
    const unit = [0, 1, 2, 3].find((candidate) => at >= x(candidate) && at < x(candidate) + 1);
    if (unit === undefined) return null;
    const leftHalf = at < x(unit) + 0.5;
    return { region: 'body', pos: leftHalf !== rtl ? unit : unit + 1, target: 'text' };
  };
}

describe('semantic anchor geometry', () => {
  test('resolves every target kind and returns versioned, page-aware fragments', () => {
    const { geometry } = semanticGeometry();
    const targets: DocxGeometryTarget[] = [
      PARAGRAPH_TARGET,
      { kind: 'range', version: 'v1', range: TEXT_RANGE },
      { kind: 'search', paragraph: PARAGRAPH, text: 'aa' },
      { kind: 'revision', revisionId: 'r1' },
      { kind: 'proposal', id: 'proposal' },
    ];
    for (const target of targets) {
      const result = anchored(geometry.getAnchorGeometry(target));
      expect(result).toMatchObject({ version: 'v1', previewVersion: 0, layoutId: 'layout' });
      expect(result.rects.every((rect) => rect.pageIndex === 0 && rect.width > 0)).toBe(true);
      const last = result.rects.at(-1)!;
      expect(result.anchor).toEqual({ ...last, x: last.x + last.width, width: 0 });
    }
  });

  test('keeps both pages and places the anchor and page rectangle on the last page', () => {
    const { geometry, source } = semanticGeometry();
    source.rangeRects = () => [RANGE, { ...RANGE, pageIndex: 1, x: 4, width: 9 }];
    const result = anchored(
      geometry.getAnchorGeometry({ kind: 'range', version: 'v1', range: TEXT_RANGE })
    );
    expect(result.rects.map((rect) => rect.pageIndex)).toEqual([0, 1]);
    expect(result.anchor).toEqual({ pageIndex: 1, x: 156, y: 354, width: 0, height: 40 });
    expect(result.pageRect).toEqual({ x: 143, y: 334, width: 100, height: 200 });
  });

  test('keeps only visible fragments of a partly hidden range', () => {
    const { geometry, source } = semanticGeometry();
    source.rangeRects = () => [
      { ...RANGE, width: 0 },
      { ...RANGE, x: 12, width: 2 },
    ];
    const result = anchored(geometry.getAnchorGeometry(PARAGRAPH_TARGET));
    expect(result.rects).toEqual([{ pageIndex: 0, x: 155, y: 114, width: 2, height: 40 }]);
    expect(result.anchor).toEqual({ pageIndex: 0, x: 157, y: 114, width: 0, height: 40 });
  });

  test('measures zoom, moving pages and layer scroll at call time', () => {
    for (const zoom of [0.5, 1, 2]) {
      const { geometry, layer, movePages } = semanticGeometry(zoom);
      const result = () => anchored(geometry.getAnchorGeometry(PARAGRAPH_TARGET));
      expect(result().rects[0]).toEqual({
        pageIndex: 0,
        x: 143 + 10 * zoom,
        y: 94 + 20 * zoom,
        width: 4 * zoom,
        height: 40 * zoom,
      });
      place(layer, rectAt(20, 10, 900, 2100), { scrollLeft: 25, scrollTop: 37 });
      expect(result().rects[0]).toMatchObject({ x: 163 + 10 * zoom, y: 124 + 20 * zoom });
      movePages(80, 20);
      expect(result().rects[0]).toMatchObject({ x: 113 + 10 * zoom, y: 84 + 20 * zoom });
    }
  });

  test('answers in layer pixels under an ancestor CSS zoom', () => {
    for (const ancestor of [0.713, 1.25]) {
      for (const zoom of [1, 1.5]) {
        const { geometry, dom } = semanticGeometry(zoom, ancestor);
        const result = anchored(geometry.getAnchorGeometry(PARAGRAPH_TARGET));
        const expected = {
          pageIndex: 0,
          x: 143 + 10 * zoom,
          y: 94 + 20 * zoom,
          width: 4 * zoom,
          height: 40 * zoom,
        };
        for (const key of ['x', 'y', 'width', 'height'] as const) {
          expect(result.rects[0]![key]).toBeCloseTo(expected[key], 9);
        }
        expect(result.pageRect.x).toBeCloseTo(143, 9);
        expect(result.pageRect.y).toBeCloseTo(94, 9);
        expect(result.pageRect.width).toBeCloseTo(PAGE.width * zoom, 9);
        expect(result.pageRect.height).toBeCloseTo(PAGE.height * zoom, 9);

        const [contextRect] = dom.getRectsForRange(0, 4);
        expect(contextRect!.x).toBeCloseTo(30 / zoom + 10, 9);
        expect(contextRect!.width).toBeCloseTo(4, 9);
        const overlay = geometry.toOverlayRect(contextRect!)!;
        expect(overlay.x).toBeCloseTo(expected.x, 9);
        expect(overlay.y).toBeCloseTo(expected.y, 9);
        expect(overlay.width).toBeCloseTo(expected.width, 9);
      }
    }
  });

  test('resolves a client point to page pixels under an ancestor CSS zoom', () => {
    for (const ancestor of [0.713, 1.25]) {
      for (const zoom of [1, 1.5]) {
        const { dom, source, canvases } = semanticGeometry(zoom, ancestor);
        const hits: [number, number, number][] = [];
        source.hitTestRegions = (pageIndex, x, y) => {
          hits.push([pageIndex, x, y]);
          return { region: 'body', pos: 3, target: 'text' } as ReturnType<
            DisplayListQueries['hitTestRegions']
          >;
        };
        const page = canvases[1]!.getBoundingClientRect();
        const scale = zoom * ancestor;
        expect(
          dom.getPositionAtPoint!(page.left + 10 * scale, page.top + 20 * scale)
        ).toMatchObject({ position: 3, pageIndex: 1 });
        expect(hits[0]![0]).toBe(1);
        expect(hits[0]![1]).toBeCloseTo(10, 9);
        expect(hits[0]![2]).toBeCloseTo(20, 9);
      }
    }
  });

  test('anchors a hidden target at its boundary, then falls back to its paragraph', () => {
    const { geometry, source } = semanticGeometry();
    const target: DocxGeometryTarget = { kind: 'revision', revisionId: 'r2' };
    textLine(source, [2, 3]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [],
      anchor: { pageIndex: 0, x: 155, width: 0 },
    });
    textLine(source, [0, 1, 2, 3]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [],
      anchor: { pageIndex: 0, x: 153, width: 0 },
    });
    source.rangeRects = () => [];
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [],
      anchor: { pageIndex: 0, x: 146, width: 0 },
    });
    source.anchorRect = () => null;
    refused(geometry.getAnchorGeometry(target), 'layout-unavailable');
  });

  test('anchors adjacent hidden revisions at the visible edge they share', () => {
    const { geometry, session, source, setSnapshot } = semanticGeometry();
    session.listRevisions = () =>
      [
        ['cd', 1, 2],
        ['ef', 2, 4],
      ].map(([revisionId, start, end]) => ({
        revisionId,
        kind: 'deletion',
        story: 'body',
        range: { start: { paraId: 'p', offset: start }, end: { paraId: 'p', offset: end } },
      })) as ReturnType<YrsSession['listRevisions']>;
    setSnapshot({
      proposals: [
        {
          id: 'proposal',
          paragraph: PARAGRAPH,
          state: 'accepted',
          changed: true,
          revisionIds: ['cd', 'ef'],
        },
      ],
    });
    textLine(source, [1, 2, 3]);
    for (const revisionId of ['cd', 'ef']) {
      expect(anchored(geometry.getAnchorGeometry({ kind: 'revision', revisionId }))).toMatchObject({
        rects: [],
        anchor: { x: 154, width: 0 },
      });
    }
    textLine(source, [0, 1, 2, 3]);
    expect(
      anchored(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'ef' })).anchor
    ).toMatchObject({ x: 153 });
  });

  test('finds the last drawn unit behind a long hidden suffix', () => {
    for (const rtl of [false, true]) {
      const { geometry, session, source } = semanticGeometry();
      session.paragraphSpans = () => [{ paraId: 'p', length: 100 }];
      const x = (unit: number) => 10 + (rtl ? 2 - unit : unit);
      source.rangeRects = (from, to) => {
        const rects: DisplayListRect[] = [];
        for (let unit = Math.max(0, from); unit < Math.min(3, to); unit += 1) {
          rects.push({ ...RANGE, x: x(unit), width: 1 });
        }
        return rects;
      };
      source.hitTestRegions = (_page, at) => {
        const unit = [0, 1, 2].find((candidate) => at >= x(candidate) && at < x(candidate) + 1);
        if (unit === undefined) return null;
        const leftHalf = at < x(unit) + 0.5;
        return { region: 'body', pos: leftHalf !== rtl ? unit : unit + 1, target: 'text' };
      };
      expect(anchored(geometry.getAnchorGeometry(PARAGRAPH_TARGET)).anchor).toMatchObject({
        x: rtl ? 153 : 156,
        width: 0,
      });
    }
  });

  test('puts collapsed boundaries on the caret stop, in either direction', () => {
    const { geometry, source } = semanticGeometry();
    const at = (offset: number) =>
      anchored(
        geometry.getAnchorGeometry({
          kind: 'range',
          version: 'v1',
          range: {
            ...TEXT_RANGE,
            start: { paraId: 'p', offset },
            end: { paraId: 'p', offset },
          },
        })
      );
    textLine(source);
    expect([0, 2, 4].map((offset) => at(offset).anchor.x)).toEqual([153, 155, 157]);
    textLine(source, [], true);
    expect([0, 2, 4].map((offset) => at(offset).anchor.x)).toEqual([157, 155, 153]);
    expect(at(2).rects).toEqual([]);
  });

  test('anchors at the logical end of the last unit, whatever order the fragments come in', () => {
    const { geometry, source } = semanticGeometry();
    const target: DocxGeometryTarget = { kind: 'range', version: 'v1', range: TEXT_RANGE };
    const text = { ...RANGE, x: 30, width: 40 };
    const image = { ...RANGE, x: 10, width: 20 };
    const probe = (unit: DisplayListRect[], line: DisplayListRect[], stops: [number, number]) => {
      source.rangeRects = (from, to) => (to - from === 1 ? unit : line);
      source.hitTestRegions = (_page, at) => ({
        region: 'body',
        pos: at < unit[0]!.x + unit[0]!.width / 2 ? stops[0] : stops[1],
        target: 'text',
      });
      return anchored(geometry.getAnchorGeometry(target)).anchor;
    };
    const last = { ...text, x: 66, width: 4 };
    expect(probe([last], [text, image], [3, 4])).toMatchObject({ x: 213 });
    source.hitTestRegions = () => ({ region: 'body', pos: 3, target: 'image' });
    expect(anchored(geometry.getAnchorGeometry(target)).anchor).toMatchObject({ x: 213 });
    expect(probe([last], [text], [4, 3])).toMatchObject({ x: 209 });
    source.hitTestRegions = () => null;
    source.rangeRects = (from, to) => (to - from === 1 ? [last] : [text]);
    expect(anchored(geometry.getAnchorGeometry(target)).anchor).toMatchObject({ x: 213 });
    source.rangeRects = (from, to) =>
      to - from === 1 ? [] : [text, { ...text, y: 70, x: 10, width: 25 }, image];
    expect(anchored(geometry.getAnchorGeometry(target)).anchor).toMatchObject({
      x: 178,
      y: 164,
    });
  });

  test('refuses while the rendered context is at another zoom than the layout', () => {
    const pages = document.createElement('div');
    const layer = document.createElement('div');
    const source = queries();
    const dom = createRenderedDomContext(pages, 1, {
      displayListQueries: source,
      projector: createCanvasHostProjector(pages, source, 1),
    });
    const layout = { id: 'layout', version: 'v1', previewVersion: 0, zoom: 1.5, pageCount: 1 };
    const geometry = createPluginGeometry(
      layout,
      dom,
      layer,
      () => true,
      () => null,
      source,
      () => {
        throw new Error('a stale context must not reach the editor');
      }
    );
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    expect(geometry.toOverlayRect({ x: 0, y: 0, width: 1, height: 1 })).toBeNull();
    expect(geometry.getPositionAtPoint(0, 0)).toBeNull();
  });

  test('answers only once the pages show the layout', () => {
    const state = semanticGeometry();
    state.setPresented(false);
    refused(state.geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    state.setPresented(true);
    expect(state.geometry.getAnchorGeometry(PARAGRAPH_TARGET).ok).toBe(true);
  });

  test('refuses a range in a story the document no longer has', () => {
    const { geometry, session } = semanticGeometry();
    session.hasStory = () => false;
    refused(
      geometry.getAnchorGeometry({
        kind: 'range',
        version: 'v1',
        range: { ...TEXT_RANGE, story: 'body:t9:r0c0' },
      }),
      'missing-target'
    );
  });

  test('refuses unavailable or stale pixels, pending input and preview mismatches', () => {
    const state = semanticGeometry();
    const { geometry, session, editor } = state;
    state.setCurrent(false);
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    state.setCurrent(true);
    state.setAvailable(false);
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    state.setAvailable(true);
    session.version = () => 'v2';
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'stale-version');
    session.version = () => 'v1';
    refused(
      geometry.getAnchorGeometry({ kind: 'range', version: 'v0', range: TEXT_RANGE }),
      'stale-version'
    );
    editor.hasPendingInput = () => true;
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    editor.hasPendingInput = () => false;
    state.setSnapshot({ previewVersion: 1 });
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    state.setSnapshot({ previewVersion: 0 });
    const key = spyOn(proposalPreview, 'currentPreviewKey').mockReturnValue('new-preview');
    try {
      refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    } finally {
      key.mockRestore();
    }
  });

  test('follows a preview change only once the queries show the new preview', () => {
    const { geometry, setSnapshot, source, session } = semanticGeometry();
    const accepted = {
      id: 'proposal',
      paragraph: PARAGRAPH,
      state: 'accepted' as const,
      changed: true,
      revisionIds: ['r1', 'r2'],
    };
    const target: DocxGeometryTarget = { kind: 'proposal', id: 'proposal' };
    stampSourceVersion(source, 'v1');
    expect(proposalPreview.currentPreviewKey(session)).toBe('');
    expect(proposalPreview.renderedPreviewKey(source)).toBe('');
    anchored(geometry.getAnchorGeometry(target));
    expect(pluginLayout(source, 'v1', 1, { key: '', previewVersion: 0 })).not.toBeNull();

    setSnapshot({ proposals: [accepted] }, false);
    const key = revisionPreviewKey({ r1: 'accepted', r2: 'accepted' });
    expect(proposalPreview.currentPreviewKey(session)).toBe(key);
    expect(proposalPreview.renderedPreviewKey(source)).toBe('');
    refused(geometry.getAnchorGeometry(target), 'layout-unavailable');
    expect(pluginLayout(source, 'v1', 1, { key, previewVersion: 0 })).toBeNull();

    stampRevisionPreviewKey(source, key);
    expect(proposalPreview.renderedPreviewKey(source)).toBe(key);
    anchored(geometry.getAnchorGeometry(target));
    expect(pluginLayout(source, 'v1', 1, { key, previewVersion: 0 })).not.toBeNull();

    stampRevisionPreviewKey(source, UNKNOWN_REVISION_PREVIEW_KEY);
    refused(geometry.getAnchorGeometry(target), 'layout-unavailable');
    stampRevisionPreviewKey(source, key);

    setSnapshot({ proposals: [{ ...accepted, state: 'proposed' }] }, false);
    refused(geometry.getAnchorGeometry(target), 'layout-unavailable');
  });

  test('refuses missing page bounds or unmappable display positions', () => {
    const { geometry, source, editor } = semanticGeometry();
    source.pageBounds = () => null;
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'layout-unavailable');
    editor.yrsLocToDisplayPosition = () => null;
    refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'unsupported');
  });

  test('unions proposal revisions in document order and ignores revisions resolved elsewhere', () => {
    const { geometry, calls, session } = semanticGeometry();
    const result = anchored(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' }));
    expect(result.rects).toHaveLength(2);
    expect(calls).toEqual([
      [0, 1],
      [2, 4],
    ]);
    session.listRevisions = () => [];
    const fallback = anchored(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' }));
    expect(fallback).toMatchObject({ rects: [], anchor: { x: 153, width: 0 } });
  });

  test('draws only the side a decision keeps and anchors a wholly hidden one at its boundary', () => {
    const { geometry, setSnapshot, source } = semanticGeometry();
    const decide = (state: 'accepted' | 'rejected', revisionIds = ['r1', 'r2']) =>
      setSnapshot({
        proposals: [{ id: 'proposal', paragraph: PARAGRAPH, state, changed: true, revisionIds }],
      });
    const target: DocxGeometryTarget = { kind: 'proposal', id: 'proposal' };
    decide('accepted');
    textLine(source, [0]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [{ x: 155 }, { x: 156 }],
      anchor: { x: 157, width: 0 },
    });
    decide('rejected');
    textLine(source, [2, 3]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [{ x: 153 }],
      anchor: { x: 154, width: 0 },
    });
    expect(
      anchored(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'r2' })).rects
    ).toEqual([]);
    decide('accepted', ['r1']);
    textLine(source, [0]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [],
      anchor: { x: 154, width: 0 },
    });
  });

  test('one session reads the document once across preview decisions at a version', () => {
    const { geometry, setSnapshot, source, session } = semanticGeometry(1, 1, { stableSession: true });
    const reads = { listRevisions: 0, paragraphSpans: 0 };
    for (const name of ['listRevisions', 'paragraphSpans'] as const) {
      const read = session[name].bind(session) as () => unknown;
      (session as unknown as Record<string, () => unknown>)[name] = () => {
        reads[name] += 1;
        return read();
      };
    }
    const decide = (state: 'proposed' | 'accepted' | 'rejected') =>
      setSnapshot({
        proposals: [
          { id: 'proposal', paragraph: PARAGRAPH, state, changed: true, revisionIds: ['r1', 'r2'] },
        ],
      });
    const target: DocxGeometryTarget = { kind: 'proposal', id: 'proposal' };
    decide('accepted');
    textLine(source, [0]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [{ x: 155 }, { x: 156 }],
      anchor: { x: 157, width: 0 },
    });
    decide('rejected');
    textLine(source, [2, 3]);
    expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
      rects: [{ x: 153 }],
      anchor: { x: 154, width: 0 },
    });
    decide('proposed');
    textLine(source);
    expect(anchored(geometry.getAnchorGeometry(target)).rects).toHaveLength(3);
    expect(reads).toEqual({ listRevisions: 1, paragraphSpans: 1 });

    session.version = () => 'v2';
    refused(geometry.getAnchorGeometry(target), 'stale-version');
  });

  test('hides the same text from range and search targets as from its revision', () => {
    const { geometry, setSnapshot, source } = semanticGeometry();
    const decide = (state: 'accepted' | 'rejected') =>
      setSnapshot({
        proposals: [
          { id: 'proposal', paragraph: PARAGRAPH, state, changed: true, revisionIds: ['r1', 'r2'] },
        ],
      });
    decide('accepted');
    textLine(source, [0]);
    const deleted: DocxGeometryTarget = {
      kind: 'range',
      version: 'v1',
      range: { ...TEXT_RANGE, end: { paraId: 'p', offset: 1 } },
    };
    for (const target of [deleted, { kind: 'revision', revisionId: 'r1' } as const]) {
      expect(anchored(geometry.getAnchorGeometry(target))).toMatchObject({
        rects: [],
        anchor: { x: 154, width: 0 },
      });
    }
    decide('rejected');
    textLine(source, [2, 3]);
    expect(
      anchored(
        geometry.getAnchorGeometry({
          kind: 'search',
          paragraph: PARAGRAPH,
          text: 'aa',
          occurrence: 2,
        })
      )
    ).toMatchObject({ rects: [], anchor: { x: 155 } });
    expect(
      anchored(geometry.getAnchorGeometry({ kind: 'paragraph', paragraph: PARAGRAPH }))
    ).toMatchObject({ rects: [{ x: 153 }, { x: 154 }], anchor: { x: 155 } });
  });

  test('coalesces overlapping revision ranges', () => {
    const { geometry, session, calls } = semanticGeometry();
    const revisions = session.listRevisions();
    revisions[1]!.range.end.offset = 3;
    session.listRevisions = () => revisions;
    anchored(geometry.getAnchorGeometry({ kind: 'proposal', id: 'proposal' }));
    expect(calls).toEqual([[0, 4]]);
  });

  test('anchors no-op proposals at their paragraph and refuses unknown ids', () => {
    const { geometry, setSnapshot, session } = semanticGeometry();
    setSnapshot({
      proposals: [
        { id: 'noop', paragraph: PARAGRAPH, state: 'proposed', changed: false, revisionIds: [] },
      ],
    });
    expect(anchored(geometry.getAnchorGeometry({ kind: 'proposal', id: 'noop' }))).toMatchObject({
      rects: [],
      anchor: { x: 153, width: 0 },
    });
    refused(geometry.getAnchorGeometry({ kind: 'proposal', id: 'missing' }), 'unknown-proposal');
    refused(
      geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'missing' }),
      'missing-target'
    );
    Reflect.deleteProperty(session, 'getProposals');
    refused(geometry.getAnchorGeometry({ kind: 'proposal', id: 'noop' }), 'unknown-proposal');
  });

  test('refuses ambiguous persisted paragraphs, missing paragraphs and source-only anchors', () => {
    const { geometry, session, calls } = semanticGeometry();
    const target: DocxGeometryTarget = {
      kind: 'paragraph',
      paragraph: {
        kind: 'persisted',
        story: { kind: 'body', partUri: '/word/document.xml' },
        paraId: '00000001',
      },
    };
    session.resolveParagraphAnchor = () => ({
      status: 'ambiguous',
      candidates: [PARAGRAPH, PARAGRAPH],
    });
    refused(geometry.getAnchorGeometry(target), 'ambiguous-target');
    expect(calls).toEqual([]);
    session.resolveParagraphAnchor = () => ({ status: 'missing' });
    refused(geometry.getAnchorGeometry(target), 'missing-target');
    session.resolveParagraphAnchor = () => ({ status: 'unsupported', reason: 'foreign-session' });
    refused(geometry.getAnchorGeometry(target), 'unsupported');
    session.resolveParagraphAnchor = () => ({
      status: 'found',
      anchor: {
        kind: 'source',
        packageSha256: 'hash',
        partUri: '/word/header1.xml',
        paragraphOrdinal: 0,
      },
    });
    refused(geometry.getAnchorGeometry(target), 'unsupported');
  });

  test('allows body-rooted stories and refuses other stories', () => {
    const { geometry, session } = semanticGeometry();
    for (const story of ['body', 'body:t0:r0c0', 'body:sdt0']) {
      session.resolveParagraphAnchor = () => ({ status: 'found', anchor: { ...PARAGRAPH, story } });
      expect(geometry.getAnchorGeometry(PARAGRAPH_TARGET).ok).toBe(true);
    }
    for (const story of ['header:rId1', 'footer:rId2', 'footnote:1', 'comment:1', 'bodyOther']) {
      session.resolveParagraphAnchor = () => ({ status: 'found', anchor: { ...PARAGRAPH, story } });
      refused(geometry.getAnchorGeometry(PARAGRAPH_TARGET), 'unsupported');
      refused(
        geometry.getAnchorGeometry({
          kind: 'range',
          version: 'v1',
          range: { ...TEXT_RANGE, story },
        }),
        'unsupported'
      );
      const revisions = session.listRevisions().map((revision) => ({ ...revision, story }));
      session.listRevisions = () => revisions;
      refused(geometry.getAnchorGeometry({ kind: 'revision', revisionId: 'r1' }), 'unsupported');
    }
  });

  test('selects first, nth and all non-overlapping search occurrences', () => {
    const { geometry, calls } = semanticGeometry();
    for (const [occurrence, expected] of [
      [undefined, [[0, 2]]],
      ['first', [[0, 2]]],
      [1, [[0, 2]]],
      [2, [[2, 4]]],
      [
        'all',
        [
          [0, 2],
          [2, 4],
        ],
      ],
    ] as const) {
      calls.length = 0;
      anchored(
        geometry.getAnchorGeometry({ kind: 'search', paragraph: PARAGRAPH, text: 'aa', occurrence })
      );
      expect(calls).toEqual(expected.map(([from, to]): [number, number] => [from, to]));
    }
    for (const occurrence of [0, -1, 1.5, 3, NaN, Infinity]) {
      refused(
        geometry.getAnchorGeometry({
          kind: 'search',
          paragraph: PARAGRAPH,
          text: 'aa',
          occurrence,
        }),
        'missing-target'
      );
    }
    refused(
      geometry.getAnchorGeometry({ kind: 'search', paragraph: PARAGRAPH, text: '' }),
      'missing-target'
    );
  });

  test('answers occurrences a truncated search covers and refuses the rest', () => {
    const { geometry, session } = semanticGeometry();
    const target: DocxGeometryTarget = { kind: 'search', paragraph: PARAGRAPH, text: 'aa' };
    const full = session.findText({
      text: 'aa',
      within: { kind: 'paragraph', story: 'body', paraId: 'p' },
      view: 'accepted',
    });
    session.findText = () => ({ ...(full as Extract<typeof full, { ok: true }>), truncated: true });
    expect(geometry.getAnchorGeometry(target).ok).toBe(true);
    expect(geometry.getAnchorGeometry({ ...target, occurrence: 2 }).ok).toBe(true);
    refused(geometry.getAnchorGeometry({ ...target, occurrence: 3 }), 'unsupported');
    refused(geometry.getAnchorGeometry({ ...target, occurrence: 'all' }), 'unsupported');
  });

  test('refuses truncated searches and preserves navigation refusal codes', () => {
    const { geometry, session } = semanticGeometry();
    const target: DocxGeometryTarget = { kind: 'search', paragraph: PARAGRAPH, text: 'aa' };
    session.findText = () => ({ ok: true, version: 'v1', matches: [], truncated: true });
    refused(geometry.getAnchorGeometry(target), 'unsupported');
    session.findText = () => ({ ok: true, version: 'v1', matches: [], truncated: false });
    refused(geometry.getAnchorGeometry(target), 'missing-target');
    for (const code of [
      'stale-version',
      'missing-target',
      'ambiguous-target',
      'limit-exceeded',
    ] as const) {
      session.findText = () => ({
        ok: false,
        version: 'v1',
        failure: { code, message: 'refused' },
      });
      refused(geometry.getAnchorGeometry(target), code === 'limit-exceeded' ? 'unsupported' : code);
    }
  });
});
