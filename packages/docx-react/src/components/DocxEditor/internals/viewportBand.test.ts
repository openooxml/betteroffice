import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { renderedScale } from '@betteroffice/docx/layout/render';
import { viewportColumnBand } from './viewportBand';

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function box(
  rect: { top: number; height: number; width?: number },
  layout: { offsetHeight: number; offsetWidth?: number; clientHeight?: number }
): HTMLElement {
  return {
    getBoundingClientRect: () => ({ width: 0, ...rect }),
    offsetWidth: 0,
    clientHeight: layout.offsetHeight,
    ...layout,
  } as unknown as HTMLElement;
}

describe('viewportColumnBand', () => {
  test('reads the band in column layout pixels under an ancestor CSS zoom', () => {
    for (const zoom of [0.8, 1, 1.25, 0.62 * 1.1]) {
      const scrollTop = 3_000;
      const scroller = box({ top: 40, height: 1_000 * zoom }, { offsetHeight: 1_000 });
      const column = box(
        { top: 40 + (8 - scrollTop) * zoom, height: 13_696 * zoom },
        { offsetHeight: 13_696 }
      );
      const band = viewportColumnBand(scroller, column);
      expect(band.top).toBeCloseTo(scrollTop - 8);
      expect(band.bottom).toBeCloseTo(scrollTop - 8 + 1_000);
    }
  });

  test('uses the window for the root scroller', () => {
    const column = box({ top: -480, height: 8_000 }, { offsetHeight: 10_000 });
    const band = viewportColumnBand(null, column);
    expect(band.top).toBeCloseTo(600);
    expect(band.bottom).toBeCloseTo(600 + window.innerHeight / 0.8);
  });
});

describe('renderedScale', () => {
  test('falls back to the width, then to 1, for an element without a layout size', () => {
    const unsized = box({ top: 0, height: 0, width: 50 }, { offsetHeight: 0, offsetWidth: 100 });
    expect(renderedScale(unsized)).toBe(0.5);
    expect(renderedScale(box({ top: 0, height: 0 }, { offsetHeight: 0 }))).toBe(1);
  });
});
