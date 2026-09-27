import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayListQueries, DisplayListRect } from '@betteroffice/docx/layout/render';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import { stampSourceVersion } from '../components/DocxEditor/internals/layoutProvenance';
import { createPluginGeometry, pluginLayout, toOverlayRect } from './geometry';

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

      const layout = { id: 'layout', version: 'v', zoom, pageCount: 1 };
      let current = true;
      const geometry = createPluginGeometry(layout, dom, layer, () => current);
      expect(geometry.toOverlayRect(rect)).toEqual(overlay);
      current = false;
      expect(geometry.toOverlayRect(rect)).toBeNull();
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
    expect(pluginLayout(source, 'v1', 1)).toBeNull();
    stampSourceVersion(source, 'v1');
    const layout = pluginLayout(source, 'v1', 1.5);
    expect(layout).toMatchObject({ version: 'v1', zoom: 1.5, pageCount: 1 });
    expect(pluginLayout(source, 'v1', 1.5)?.id).toBe(layout!.id);
    expect(pluginLayout(source, 'v2', 1.5)).toBeNull();
    const next = queries();
    stampSourceVersion(next, 'v1');
    expect(pluginLayout(next, 'v1', 1.5)?.id).not.toBe(layout!.id);
  });
});
