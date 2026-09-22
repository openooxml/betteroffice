import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createElement } from 'react';
import {
  createDisplayListQueries,
  type DisplayListQueries,
  type DisplayPage,
  type TextRunPrimitive,
} from '@betteroffice/docx/layout/render';
import { preloadLayoutWasm } from '@betteroffice/docx/wasm/layout';
import {
  createRenderedDomContext,
  createCanvasHostProjector,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { YrsInputRef } from '../YrsInput';
import type { YrsPositionProjection } from '../internals/yrsPositionProjection';
import { usePagesPointer } from './usePagesPointer';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
let queries: DisplayListQueries;

function run(y: number, start: number): TextRunPrimitive {
  return {
    kind: 'text',
    text: 'abcdefghij',
    x: 100,
    baselineY: y,
    width: 100,
    font: '400 16px Calibri',
    color: '#000000',
    docStart: start,
    docEnd: start + 10,
    blockId: start,
    lineIndex: 0,
  };
}

beforeAll(async () => {
  await preloadLayoutWasm(
    new Uint8Array(
      readFileSync(
        resolve(
          import.meta.dir,
          '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm'
        )
      )
    )
  );
  const page: DisplayPage = {
    pageIndex: 1,
    width: 800,
    height: 1000,
    contentBounds: { x: 80, y: 80, width: 640, height: 760 },
    primitives: [
      run(150, 101),
      {
        kind: 'image',
        relId: 'rIdImage',
        x: 300,
        y: 120,
        w: 60,
        h: 40,
        docStart: 111,
        docEnd: 112,
      },
    ],
    header: { kind: 'header', rId: 'rIdHeader', y: 0, height: 80, primitives: [run(40, 1)] },
    footer: { kind: 'footer', rId: 'rIdFooter', y: 940, height: 60, primitives: [run(975, 1)] },
    noteAreas: [
      {
        kind: 'footnote',
        y: 850,
        height: 60,
        noteIds: [2],
        primitives: [{ ...run(880, 1), groupId: 'footnote-2' }],
      },
    ],
  };
  queries = createDisplayListQueries({
    pages: [{ pageIndex: 0, width: 600, height: 900, primitives: [] }, page],
  });
  await queries.whenReady();
});

afterEach(cleanup);
afterAll(async () => {
  queries?.dispose();
  if (ownsDom) await GlobalRegistrator.unregister();
});

function hostAt(zoom: number, scroll: number) {
  const host = document.createElement('div');
  const left = 40;
  const top = 60 - scroll;
  host.getBoundingClientRect = () => new DOMRect(left, top, 848 * zoom, 2020 * zoom);
  Object.defineProperty(host, 'offsetWidth', { value: 848 });
  for (let index = 0; index < 2; index += 1) {
    const size = queries.pageSize(index)!;
    const canvas = document.createElement('canvas');
    canvas.dataset.pageIndex = String(index);
    canvas.getBoundingClientRect = () =>
      new DOMRect(
        left + ((848 - size.width) * zoom) / 2,
        top + (24 + (index === 0 ? 0 : 924)) * zoom,
        size.width * zoom,
        size.height * zoom
      );
    host.append(canvas);
  }
  return host;
}

function clientPoint(host: HTMLElement, x: number, y: number) {
  const page = host.querySelector('canvas[data-page-index="1"]')!.getBoundingClientRect();
  return {
    clientX: page.left + (x * page.width) / 800,
    clientY: page.top + (y * page.height) / 1000,
  };
}

describe('public point query', () => {
  for (const zoom of [0.75, 1.5]) {
    test(`matches a page 2 click at ${zoom * 100}% after scrolling without moving focus`, () => {
      const host = hostAt(zoom, 800);
      const selections: Array<[number, number, string | undefined]> = [];
      let focused = false;
      const projection = {
        size: 200,
        targetAt: (position: number) => ({ story: 'body', displayPosition: position }),
      } as unknown as YrsPositionProjection;
      function Harness() {
        const pointer = usePagesPointer({
          pagesContainerRef: { current: host },
          canvasHostRef: { current: host },
          displayListQueries: queries,
          yrsRootStory: 'body',
          readOnly: false,
          yrsSession: { cellSelection: () => null } as unknown as YrsSession,
          yrsInputRef: {
            current: {
              focus: () => {
                focused = true;
              },
              setSelectionFromDisplay: (a: number, h: number, story?: string) =>
                selections.push([a, h, story]),
            } as unknown as YrsInputRef,
          },
          getYrsPositionProjection: () => projection,
          applyYrsCommand: () => false,
          syncYrsInputState: () => false,
          setSelectionRects: () => {},
          setCaretPosition: () => {},
          setIsFocused: () => {},
          scrollToPositionImpl: () => {},
        });
        return createElement('div', {
          'data-testid': 'pointer',
          onMouseDown: pointer.handlePagesMouseDown,
        });
      }
      const rendered = render(createElement(Harness));
      const context = createRenderedDomContext(host, zoom, {
        displayListQueries: queries,
        projector: createCanvasHostProjector(host, queries, zoom),
      });
      const point = clientPoint(host, 145, 145);
      const active = document.activeElement;
      const hit = context.getPositionAtPoint(point.clientX, point.clientY);
      expect(hit).toMatchObject({ region: 'body', pageIndex: 1 });
      expect(hit!.position).toBeGreaterThan(101);
      expect(hit!.position).toBeLessThan(111);
      expect(selections).toEqual([]);
      expect(focused).toBe(false);
      expect(document.activeElement).toBe(active);
      act(() =>
        rendered.getByTestId('pointer').dispatchEvent(
          new MouseEvent('mousedown', {
            bubbles: true,
            button: 0,
            ...point,
          })
        )
      );
      expect(selections).toEqual([[hit!.position, hit!.position, 'body']]);
      expect(focused).toBe(true);
      const prior = [...selections];
      expect(context.getPositionAtPoint(-100, -100)).toBeNull();
      expect(selections).toEqual(prior);
    });
  }

  test('retains header, footer and note identity', () => {
    const host = hostAt(1.5, 700);
    const context = createRenderedDomContext(host, 1.5, {
      displayListQueries: queries,
      projector: createCanvasHostProjector(host, queries, 1.5),
    });
    for (const [y, expected] of [
      [35, { region: 'header', rId: 'rIdHeader' }],
      [970, { region: 'footer', rId: 'rIdFooter' }],
      [875, { region: 'footnote', noteId: 2 }],
    ] as const) {
      const point = clientPoint(host, 145, y);
      expect(context.getPositionAtPoint(point.clientX, point.clientY)).toMatchObject({
        ...expected,
        pageIndex: 1,
      });
    }
  });

  test('rejects margins, images, page gaps, invalid coordinates and unavailable geometry', () => {
    const host = hostAt(0.75, 600);
    const context = createRenderedDomContext(host, 0.75, {
      displayListQueries: queries,
      projector: createCanvasHostProjector(host, queries, 0.75),
    });
    for (const [x, y] of [
      [10, 145],
      [320, 140],
      [145, -12],
      [NaN, 145],
      [Infinity, 145],
    ]) {
      const point = clientPoint(host, x, y);
      expect(context.getPositionAtPoint(point.clientX, point.clientY)).toBeNull();
    }
    expect(createRenderedDomContext(host).getPositionAtPoint(100, 100)).toBeNull();
  });

  test('projects canvas-free page hosts using the same pagination geometry', () => {
    const host = hostAt(0.75, 500);
    const point = clientPoint(host, 145, 145);
    const context = createRenderedDomContext(host, 0.75, {
      displayListQueries: queries,
      projector: createCanvasHostProjector(host, queries, 0.75),
    });
    const expected = context.getPositionAtPoint(point.clientX, point.clientY);
    host.replaceChildren();
    expect(context.getPositionAtPoint(point.clientX, point.clientY)).toEqual(expected);
  });
});
