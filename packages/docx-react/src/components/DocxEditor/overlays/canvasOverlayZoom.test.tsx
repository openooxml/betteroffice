import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import { projectPageLocalRect } from '../internals/canvasProjection';
import { CanvasImageSelectionOverlay } from './CanvasImageSelectionOverlay';
import { CanvasSelectionOverlay } from './CanvasSelectionOverlay';

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PAGE = { width: 400, height: 600 };
const IMAGE = { x: 50, y: 80, w: 120, h: 60 };
const IMAGE_POS = 7;

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

/**
 * An overlay target at client (20, 10) holding a page canvas 30 px right and 40 px down, both
 * in the target's own pixels. `ancestor` is a CSS zoom above the target: client rects carry it,
 * the target's pixels do not.
 */
function scene(zoom: number, ancestor: number) {
  const overlayTarget = document.createElement('div');
  const host = document.createElement('div');
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  host.appendChild(canvas);
  overlayTarget.appendChild(host);
  document.body.appendChild(overlayTarget);
  if (ancestor !== 1) {
    for (const element of [overlayTarget, host, canvas]) {
      Object.defineProperty(element, 'currentCSSZoom', { value: ancestor });
    }
  }
  overlayTarget.getBoundingClientRect = () => rectAt(20, 10, 1000 * ancestor, 2000 * ancestor);
  host.getBoundingClientRect = () =>
    rectAt(20, 10, PAGE.width * zoom * ancestor, PAGE.height * zoom * ancestor);
  canvas.getBoundingClientRect = () =>
    rectAt(
      20 + 30 * ancestor,
      10 + 40 * ancestor,
      PAGE.width * zoom * ancestor,
      PAGE.height * zoom * ancestor
    );
  const displayList = {
    pages: [
      {
        pageIndex: 0,
        ...PAGE,
        primitives: [{ kind: 'image', relId: 'rId1', docStart: IMAGE_POS, ...IMAGE }],
      },
    ],
  } as unknown as DisplayList;
  const queries = {
    displayList,
    pageSize: () => PAGE,
  } as unknown as DisplayListQueries;
  const hostRef = createRef<HTMLDivElement>() as React.MutableRefObject<HTMLDivElement | null>;
  hostRef.current = host;
  return { overlayTarget, host, hostRef, displayList, queries };
}

function box(element: Element) {
  const style = (element as HTMLElement).style;
  return {
    left: Number.parseFloat(style.left),
    top: Number.parseFloat(style.top),
    width: Number.parseFloat(style.width),
    height: Number.parseFloat(style.height),
  };
}

function expectBox(actual: ReturnType<typeof box>, expected: ReturnType<typeof box>) {
  expect(actual.left).toBeCloseTo(expected.left, 6);
  expect(actual.top).toBeCloseTo(expected.top, 6);
  expect(actual.width).toBeCloseTo(expected.width, 6);
  expect(actual.height).toBeCloseTo(expected.height, 6);
}

describe('canvas overlays under an ancestor CSS zoom', () => {
  test('project page-local rects into the overlay target pixels', () => {
    for (const ancestor of [1, 0.713, 1.25]) {
      for (const zoom of [1, 1.5]) {
        const { overlayTarget, host, queries } = scene(zoom, ancestor);
        const projected = projectPageLocalRect(host, overlayTarget, queries, 0, 10, 20, 30, 5)!;
        expectBox(projected, {
          left: 30 + 10 * zoom,
          top: 40 + 20 * zoom,
          width: 30 * zoom,
          height: 5 * zoom,
        });
        expect(projected.scaleX).toBeCloseTo(zoom, 9);
        expect(projected.targetZoom).toBe(ancestor);
        overlayTarget.remove();
      }
    }
  });

  test('keeps the zoom 1 arithmetic exact', () => {
    const { overlayTarget, host, queries } = scene(1.3, 1);
    const canvasRect = host.querySelector('canvas')!.getBoundingClientRect();
    const scale = canvasRect.width / PAGE.width;
    expect(projectPageLocalRect(host, overlayTarget, queries, 0, 10.1, 20.3, 30.7, 5.9)).toEqual({
      left: canvasRect.left - 20 + 10.1 * scale,
      top: canvasRect.top - 10 + 20.3 * scale,
      width: 30.7 * scale,
      height: 5.9 * (canvasRect.height / PAGE.height),
      scaleX: scale,
      scaleY: canvasRect.height / PAGE.height,
      targetZoom: 1,
    });
  });

  test('place the selection highlight over the painted text', () => {
    for (const ancestor of [0.713, 1.25]) {
      const zoom = 1.5;
      const { overlayTarget, hostRef, displayList, queries } = scene(zoom, ancestor);
      render(
        <CanvasSelectionOverlay
          selectionRects={[{ x: 10, y: 20, width: 30, height: 5, pageIndex: 0 }]}
          caretPosition={null}
          isFocused
          overlayTarget={overlayTarget}
          canvasHostRef={hostRef}
          displayList={displayList}
          displayListQueries={queries}
          directProjection={false}
          sidebarOpen={false}
          zoom={zoom}
        />
      );
      const rect = overlayTarget.querySelector('[data-testid="selection-rect-0"]')!;
      expectBox(box(rect), { left: 30 + 10 * zoom, top: 40 + 20 * zoom, width: 45, height: 7.5 });
      cleanup();
      overlayTarget.remove();
    }
  });

  test('image handles sit on the image and resize by the pointer distance in page pixels', () => {
    for (const ancestor of [0.713, 1.25]) {
      const zoom = 1.5;
      const { overlayTarget, hostRef, queries } = scene(zoom, ancestor);
      const resized: [number, number, number][] = [];
      render(
        <CanvasImageSelectionOverlay
          pmPos={IMAGE_POS}
          isFocused
          overlayTarget={overlayTarget}
          canvasHostRef={hostRef}
          displayListQueries={queries}
          sidebarOpen={false}
          zoom={zoom}
          onResize={(pos, width, height) => resized.push([pos, width, height])}
        />
      );
      const se = overlayTarget.querySelector<HTMLElement>('[data-handle="se"]')!;
      const left = 30 + (IMAGE.x + IMAGE.w) * zoom;
      const top = 40 + (IMAGE.y + IMAGE.h) * zoom;
      expect(Number.parseFloat(se.style.left) + 5).toBeCloseTo(left, 6);
      expect(Number.parseFloat(se.style.top) + 5).toBeCloseTo(top, 6);

      const clientX = 20 + left * ancestor;
      const clientY = 10 + top * ancestor;
      const grow = 30;
      act(() => {
        fireEvent.mouseDown(se, { clientX, clientY });
      });
      act(() => {
        fireEvent.mouseMove(window, {
          clientX: clientX + grow * zoom * ancestor,
          clientY: clientY + (grow / 2) * zoom * ancestor,
          shiftKey: true,
        });
      });
      act(() => {
        fireEvent.mouseUp(window);
      });
      expect(resized).toEqual([[IMAGE_POS, IMAGE.w + grow, IMAGE.h + grow / 2]]);
      cleanup();
      overlayTarget.remove();
    }
  });
});
