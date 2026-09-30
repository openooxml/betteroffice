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
function zoomed(
  zoom: number,
  top: number,
  boxHeight: number,
  clientHeight = Math.round(boxHeight)
) {
  const element = document.createElement('div');
  Object.defineProperties(element, {
    currentCSSZoom: { value: zoom },
    clientHeight: { value: clientHeight },
  });
  element.getBoundingClientRect = () =>
    ({ top, height: boxHeight * zoom, bottom: top + boxHeight * zoom } as DOMRect);
  return element;
}

describe('effectiveZoom', () => {
  test('is exactly 1 without a zoom', () => {
    expect(effectiveZoom(zoomed(1, 0, 300.484375))).toBe(1);
  });

  test('multiplies the zoom of flat-tree ancestors where currentCSSZoom is unavailable', () => {
    const outer = document.createElement('div');
    const inner = document.createElement('div');
    const leaf = document.createElement('div');
    outer.style.setProperty('zoom', '0.64');
    inner.style.setProperty('zoom', '1.1');
    outer.appendChild(inner);
    inner.appendChild(leaf);
    document.body.appendChild(outer);
    expect(effectiveZoom(leaf)).toBeCloseTo(0.704);

    const host = document.createElement('div');
    host.style.setProperty('zoom', '0.8');
    const shadow = host.attachShadow({ mode: 'open' });
    const shadowChild = document.createElement('div');
    shadow.appendChild(shadowChild);
    document.body.appendChild(host);
    expect(effectiveZoom(shadowChild)).toBeCloseTo(0.8);

    const slotHost = document.createElement('div');
    slotHost.style.setProperty('zoom', '0.6');
    const slot = document.createElement('slot');
    slot.style.setProperty('zoom', '0.5');
    slotHost.attachShadow({ mode: 'open' }).appendChild(slot);
    const slotted = document.createElement('div');
    slotHost.appendChild(slotted);
    if (slotted.assignedSlot !== slot)
      Object.defineProperty(slotted, 'assignedSlot', { value: slot });
    document.body.appendChild(slotHost);
    expect(effectiveZoom(slotted)).toBeCloseTo(0.3);
    slotHost.remove();
    outer.remove();
    host.remove();
  });
});

describe('scrollViewport', () => {
  test('treats only the scrolling element as the root', () => {
    expect(scrollViewport(document.documentElement)).toEqual({
      top: 0,
      bottom: window.innerHeight,
      height: window.innerHeight,
      zoom: 1,
    });
    const body = scrollViewport(document.body);
    expect(body.height).toBe(document.body.clientHeight);
    expect(body.top).toBe(document.body.getBoundingClientRect().top);
  });
});

describe('viewportColumnBand', () => {
  test('reads the band in column layout pixels under an ancestor CSS zoom', () => {
    for (const zoom of [0.8, 1, 1.25, 0.64 * 1.1]) {
      const scrollTop = 3_000.5;
      const scroller = zoomed(zoom, 40.25, 1_000.4, 1_000);
      const column = zoomed(zoom, 40.25 + (8 - scrollTop) * zoom, 13_696.3);
      const band = viewportColumnBand(scroller, column);
      expect(band.top - band.columnTop).toBeCloseTo(scrollTop - 8, 9);
      expect(band.height).toBeCloseTo(1_000, 9);
    }
  });

  test('passes unzoomed client rects and clientHeight through unchanged', () => {
    const scroller = zoomed(1, 40.3, 300.484375, 300);
    const column = zoomed(1, -2_000.1, 13_696.3);
    expect(viewportColumnBand(scroller, column)).toEqual({
      columnTop: -2_000.1,
      top: 40.3,
      height: 300,
    });
  });

  test('uses the window for the root scroller, whose rect moves with the page', () => {
    const column = zoomed(0.8, -480, 10_000);
    for (const scroller of [null, document.documentElement]) {
      const band = viewportColumnBand(scroller, column);
      expect(band.top - band.columnTop).toBeCloseTo(600);
      expect(band.height).toBeCloseTo(window.innerHeight / 0.8);
    }
  });
});

describe('scrollIntoViewDelta', () => {
  test('keeps the margin in layout pixels under an ancestor CSS zoom', () => {
    const viewport = { top: 100, bottom: 500, height: 800, zoom: 0.5 };
    // 20 client pixels below the top is 40 layout pixels in: inside the margin already
    expect(scrollIntoViewDelta(viewport, 120, 130, 24)).toBe(0);
    expect(scrollIntoViewDelta(viewport, 105, 115, 24)).toBe(10 - 24);
    expect(scrollIntoViewDelta(viewport, 480, 495, 24)).toBe(-10 + 24);
  });

  test('scrolls an unzoomed caret exactly as the box edges did before', () => {
    const scroller = zoomed(1, 40.25, 300.484375, 300);
    const viewport = scrollViewport(scroller);
    const box = scroller.getBoundingClientRect();
    const bottom = box.bottom - 24 + 0.2;
    expect(scrollIntoViewDelta(viewport, bottom - 16, bottom, 24)).toBe(bottom - box.bottom + 24);
    expect(scrollIntoViewDelta(viewport, box.top + 3.5, box.top + 19.5, 24)).toBe(
      box.top + 3.5 - box.top - 24
    );
  });
});
