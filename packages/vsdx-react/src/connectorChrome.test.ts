import { expect, test } from 'bun:test';
import type { PageDisplayList, PageSnapshot } from '@betteroffice/vsdx';
import { connectorRuns, dragSegmentRoute, hitSegmentDot, paintConnectorChrome, previewChrome, selectedConnectorChrome } from './connectorChrome';
import type { ChromePoint, SelectedConnectorChrome } from './connectorChrome';
import { grabTolerance, sameRoutePoints } from './VsdxEditor';

const page: PageSnapshot = {
  id: 'page:1',
  sourcePartPath: 'visio/pages/page1.xml',
  name: 'Page',
  shapes: [{ id: 'page:1:shape:7', sourceId: 7, name: 'Connector', cells: [], children: [] }],
};

function frameWith(connectorId: string): PageDisplayList {
  return {
    contractVersion: 7,
    width: 816,
    height: 1056,
    printWidth: 816,
    printHeight: 1056,
    paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 1056 },
    primitives: [
      {
        kind: 'shape',
        id: connectorId,
        zOrder: 0,
        path: [
          { type: 'move', x: 1, y: 1 },
          { type: 'line', x: 4, y: 1 },
          { type: 'line', x: 4, y: 3 },
        ],
        transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      },
    ],
    connectors: [{ id: connectorId, begin: 'point', end: 'free', routable: true }],
  } as unknown as PageDisplayList;
}

const frame = frameWith('visio/pages/page1.xml:7');

test('selects the connector route with one segment per straight run', () => {
  const selected = selectedConnectorChrome(frame, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  expect(selected?.start).toEqual({ x: 1, y: 1 });
  expect(selected?.end).toEqual({ x: 4, y: 3 });
  expect(selected?.segments).toEqual([
    { a: { x: 1, y: 1 }, b: { x: 4, y: 1 }, mid: { x: 2.5, y: 1 }, index: 0 },
    { a: { x: 4, y: 1 }, b: { x: 4, y: 3 }, mid: { x: 4, y: 2 }, index: 1 },
  ]);
});

test('ignores selections that are not a painted connector', () => {
  expect(selectedConnectorChrome(frame, page, { pageId: 'page:1', shapeId: 'missing' })).toBeNull();
  expect(selectedConnectorChrome(frame, page, { pageId: 'page:2', shapeId: 'page:1:shape:7' })).toBeNull();
  expect(selectedConnectorChrome({ ...frame, connectors: [] }, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' })).toBeNull();
  expect(selectedConnectorChrome({ ...frame, primitives: [] }, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' })).toBeNull();
});

test('breaks runs at curves and keeps repeated vertices in the run', () => {
  const runs = connectorRuns([
    { type: 'move', x: 0, y: 0 },
    { type: 'line', x: 2, y: 0 },
    { type: 'line', x: 2, y: 0 },
    { type: 'cubic', cp1x: 3, cp1y: 0, cp2x: 3, cp2y: 2, x: 4, y: 2 },
    { type: 'line', x: 6, y: 2 },
  ]);
  expect(runs).toEqual([
    [{ x: 0, y: 0 }, { x: 2, y: 0 }, { x: 2, y: 0 }],
    [{ x: 4, y: 2 }, { x: 6, y: 2 }],
  ]);
});

function paintRecorder(calls: string[]): CanvasRenderingContext2D {
  const context = {
    save: () => calls.push('save'),
    restore: () => calls.push('restore'),
    setTransform: () => {},
    transform: () => {},
    beginPath: () => calls.push('beginPath'),
    arc: (x: number, y: number) => calls.push(`arc:${x},${y}`),
    fill: () => calls.push(`fill:${(context as unknown as Record<string, string>).fillStyle}`),
    stroke: () => calls.push(`stroke:${(context as unknown as Record<string, string>).strokeStyle}`),
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 0,
  };
  return context as unknown as CanvasRenderingContext2D;
}

test('paints glue-coded endpoints and one blue dot per segment', () => {
  const selected = selectedConnectorChrome(frame, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  const calls: string[] = [];
  const context = paintRecorder(calls);
  paintConnectorChrome(context, frame, selected as SelectedConnectorChrome, 1, 1);
  expect(calls).toContain('arc:2.5,1');
  expect(calls).toContain('arc:4,2');
  expect(calls).toContain('fill:#2563eb');
  expect(calls).toContain('stroke:#16a34a');
  expect(calls).toContain('stroke:#64748b');
});

test('ignores a connector painted inside a group', () => {
  const grouped: PageDisplayList = {
    ...frame,
    primitives: [{ kind: 'group', id: 'group', zOrder: 0, primitives: frame.primitives, transform: { a: 1, b: 0, c: 0, d: 1, e: 5, f: 5 } }],
  } as unknown as PageDisplayList;
  expect(selectedConnectorChrome(grouped, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' })).toBeNull();
});

test('offers no handles on a connector the renderer cannot route', () => {
  const frameless: PageDisplayList = { ...frame, connectors: [{ id: 'visio/pages/page1.xml:7', begin: 'point', end: 'free', routable: false }] } as unknown as PageDisplayList;
  const selected = selectedConnectorChrome(frameless, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  expect(selected?.segments.length).toBe(2);
  expect(selected?.draggable).toBeNull();
  const calls: string[] = [];
  const context = paintRecorder(calls);
  paintConnectorChrome(context, frameless, selected as SelectedConnectorChrome, 1, 1);
  expect(calls).not.toContain('fill:#2563eb');
  expect(calls).toContain('stroke:#16a34a');
});

test('exposes a draggable run only for all-straight routes', () => {
  const selected = selectedConnectorChrome(frame, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  expect(selected?.draggable).toEqual([
    { x: 1, y: 1 },
    { x: 4, y: 1 },
    { x: 4, y: 3 },
  ]);
  const curved: PageDisplayList = {
    ...frame,
    primitives: [
      {
        kind: 'shape',
        id: 'visio/pages/page1.xml:7',
        zOrder: 0,
        path: [
          { type: 'move', x: 0, y: 0 },
          { type: 'line', x: 2, y: 0 },
          { type: 'cubic', cp1x: 3, cp1y: 0, cp2x: 3, cp2y: 2, x: 4, y: 2 },
          { type: 'line', x: 6, y: 2 },
        ],
        transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      },
    ],
  } as unknown as PageDisplayList;
  const bent = selectedConnectorChrome(curved, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  expect(bent?.segments.length).toBe(2);
  expect(bent?.draggable).toBeNull();
});

test('hits the nearest segment dot within tolerance', () => {
  const selected = selectedConnectorChrome(frame, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  expect(hitSegmentDot(selected as SelectedConnectorChrome, { x: 2.5, y: 1.05 }, 0.2)?.index).toBe(0);
  expect(hitSegmentDot(selected as SelectedConnectorChrome, { x: 4, y: 2 }, 0.2)?.index).toBe(1);
  expect(hitSegmentDot(selected as SelectedConnectorChrome, { x: 0, y: 0 }, 0.2)).toBeNull();
});

test('keys a segment to its route vertex across a repeated point', () => {
  const repeated = frameWith('visio/pages/page1.xml:7');
  (repeated.primitives[0] as { path: unknown[] }).path = [
    { type: 'move', x: 1, y: 1 },
    { type: 'line', x: 1, y: 1 },
    { type: 'line', x: 4, y: 1 },
    { type: 'line', x: 4, y: 3 },
  ];
  const selected = selectedConnectorChrome(repeated, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' }) as SelectedConnectorChrome;
  const hit = hitSegmentDot(selected, { x: 2.5, y: 1 }, 0.2);
  expect(hit?.index).toBe(1);
  expect(dragSegmentRoute(selected.draggable as ChromePoint[], hit?.index ?? -1, hit?.mid as ChromePoint, { x: 2.5, y: 2 })).toEqual([
    { x: 1, y: 1 },
    { x: 1, y: 2 },
    { x: 4, y: 2 },
    { x: 4, y: 1 },
    { x: 4, y: 3 },
  ]);
});

test('measures the grab radius in model units, not display pixels', () => {
  const canvas = { getBoundingClientRect: () => ({ width: 816 }) } as unknown as HTMLCanvasElement;
  expect(grabTolerance(canvas, frame)).toBeCloseTo(12 / 96, 9);
  const zoomed = { getBoundingClientRect: () => ({ width: 1632 }) } as unknown as HTMLCanvasElement;
  expect(grabTolerance(zoomed, frame)).toBeCloseTo(6 / 96, 9);
});

test('drags a segment perpendicular into a four-segment route', () => {
  const vertices = [{ x: 1, y: 1 }, { x: 4, y: 1 }, { x: 4, y: 3 }];
  const route = dragSegmentRoute(vertices, 0, { x: 2.5, y: 1 }, { x: 2.5, y: 2 });
  expect(route).toEqual([
    { x: 1, y: 1 },
    { x: 1, y: 2 },
    { x: 4, y: 2 },
    { x: 4, y: 1 },
    { x: 4, y: 3 },
  ]);
});

test('discards parallel travel and collapses a drag back', () => {
  const vertices = [{ x: 1, y: 1 }, { x: 4, y: 1 }, { x: 4, y: 3 }];
  expect(dragSegmentRoute(vertices, 0, { x: 2.5, y: 1 }, { x: 9, y: 1 })).toEqual(vertices);
  const moved = dragSegmentRoute(vertices, 0, { x: 2.5, y: 1 }, { x: 2.5, y: 2 });
  const back = dragSegmentRoute(moved, 1, { x: 2.5, y: 2 }, { x: 2.5, y: 1 });
  expect(back).toEqual(vertices);
});

test('drags diagonal segments along their normal', () => {
  const vertices = [{ x: 0, y: 0 }, { x: 2, y: 2 }];
  const route = dragSegmentRoute(vertices, 0, { x: 1, y: 1 }, { x: 2, y: 0 });
  expect(route.length).toBe(4);
  expect(route[0]).toEqual({ x: 0, y: 0 });
  expect(route[3]).toEqual({ x: 2, y: 2 });
  for (const [actual, wanted] of [[route[1], { x: 1, y: -1 }], [route[2], { x: 3, y: 1 }]] as const) {
    expect(actual.x).toBeCloseTo(wanted.x, 9);
    expect(actual.y).toBeCloseTo(wanted.y, 9);
  }
});

test('repaints previews with moved dots and fixed endpoints', () => {
  const selected = selectedConnectorChrome(frame, page, { pageId: 'page:1', shapeId: 'page:1:shape:7' });
  const preview = previewChrome(selected as SelectedConnectorChrome, [
    { x: 1, y: 1 },
    { x: 1, y: 2 },
    { x: 4, y: 2 },
    { x: 4, y: 1 },
    { x: 4, y: 3 },
  ]);
  expect(preview.start).toEqual({ x: 1, y: 1 });
  expect(preview.end).toEqual({ x: 4, y: 3 });
  expect(preview.segments.map((segment) => segment.index)).toEqual([0, 1, 2, 3]);
  expect(preview.segments.map((segment) => segment.mid)).toEqual([
    { x: 1, y: 1.5 },
    { x: 2.5, y: 2 },
    { x: 4, y: 1.5 },
    { x: 4, y: 2 },
  ]);
});

test('commits a segment drag only when the route actually moved', () => {
  const route = [{ x: 1, y: 1 }, { x: 4, y: 1 }, { x: 4, y: 3 }];
  expect(sameRoutePoints(route, route.map((point) => ({ ...point })))).toBe(true);
  expect(sameRoutePoints(route, [{ x: 1, y: 1 }, { x: 4, y: 1.5 }, { x: 4, y: 3 }])).toBe(false);
  expect(sameRoutePoints(route, route.slice(0, 2))).toBe(false);
});
