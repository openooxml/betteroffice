import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { effectiveZoom } from '@betteroffice/docx/layout/render';
import { scrollIntoViewDelta, scrollViewport, viewportColumnBand } from './viewportBand';

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** An element drawn at `zoom`: its client rect scales, its layout sizes do not. */
function zoomed(zoom: number, top: number, layoutHeight: number, clientHeight = layoutHeight) {
  const element = document.createElement('div');
  Object.defineProperties(element, {
    currentCSSZoom: { value: zoom },
    clientHeight: { value: clientHeight },
    offsetHeight: { value: Math.round(layoutHeight) },
  });
  element.getBoundingClientRect = () =>
    ({ top, height: layoutHeight * zoom, bottom: top + layoutHeight * zoom } as DOMRect);
  return element;
}

describe('effectiveZoom', () => {
  test('is exactly 1 without a zoom, whatever the rounding of layout sizes', () => {
    expect(effectiveZoom(zoomed(1, 0, 300.484375))).toBe(1);
  });

  test('multiplies nested zooms where currentCSSZoom is unavailable', () => {
    const outer = document.createElement('div');
    const inner = document.createElement('div');
    const leaf = document.createElement('div');
    outer.style.setProperty('zoom', '0.64');
    inner.style.setProperty('zoom', '1.1');
    outer.appendChild(inner);
    inner.appendChild(leaf);
    document.body.appendChild(outer);
    expect(effectiveZoom(leaf)).toBeCloseTo(0.704);
    outer.remove();
  });
});

describe('viewportColumnBand', () => {
  test('reads the band in column layout pixels under an ancestor CSS zoom', () => {
    for (const zoom of [0.8, 1, 1.25, 0.64 * 1.1]) {
      const scrollTop = 3_000.5;
      const scroller = zoomed(zoom, 40.25, 1_000.4);
      const column = zoomed(zoom, 40.25 + (8 - scrollTop) * zoom, 13_696.3);
      const band = viewportColumnBand(scroller, column);
      expect(band.top).toBeCloseTo(scrollTop - 8, 9);
      expect(band.bottom).toBeCloseTo(scrollTop - 8 + 1_000.4, 9);
    }
  });

  test('uses the window for the root scroller, whose rect moves with the page', () => {
    const column = zoomed(0.8, -480, 10_000);
    for (const scroller of [null, document.documentElement]) {
      const band = viewportColumnBand(scroller, column);
      expect(band.top).toBeCloseTo(600);
      expect(band.bottom).toBeCloseTo(600 + window.innerHeight / 0.8);
    }
    expect(scrollViewport(document.documentElement)).toEqual({
      top: 0,
      bottom: window.innerHeight,
      zoom: 1,
    });
  });
});

describe('scrollIntoViewDelta', () => {
  test('keeps the margin in layout pixels under an ancestor CSS zoom', () => {
    const viewport = { top: 100, bottom: 500, zoom: 0.5 };
    // 20 client pixels below the top is 40 layout pixels in: inside the margin already
    expect(scrollIntoViewDelta(viewport, 120, 130, 24)).toBe(0);
    expect(scrollIntoViewDelta(viewport, 105, 115, 24)).toBe(10 - 24);
    expect(scrollIntoViewDelta(viewport, 480, 495, 24)).toBe(790 - 800 + 24);
  });
});
