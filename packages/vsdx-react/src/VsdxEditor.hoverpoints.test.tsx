import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';
import type { DiagramHandle, ModelPoint } from '@betteroffice/vsdx';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const root = resolve(import.meta.dir, '../../..');
let foundation: Uint8Array;

beforeAll(async () => {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  foundation = fixture;
});

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const { VsdxEditor } = await import('./VsdxEditor');
const { standardShapeById } = await import('./components/shapes/shapeLibrary');
const { connectionPointsForShape, connectorRouteFromFrame, modelToPage } = await import('./connector');

afterEach(() => { cleanup(); });

interface Ready { handle: DiagramHandle; refresh: () => void; }

function stubCanvasRect(canvas: HTMLCanvasElement, width: number, height: number): void {
  canvas.getBoundingClientRect = () => ({ x: 0, y: 0, left: 0, top: 0, right: width, bottom: height, width, height, toJSON: () => {} }) as DOMRect;
}

function stubCanvas(): () => void {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  return () => { canvasPrototype.getContext = getContext; };
}

function cssForModel(frame: { width: number; height: number; paintTransform: vsdx.Affine }, model: ModelPoint, scale: number): { x: number; y: number } {
  const page = modelToPage(frame as never, model);
  return { x: page.x * scale, y: page.y * scale };
}

function pointOf(handle: DiagramHandle, shapeId: string, side: 'north' | 'east' | 'south' | 'west' | 'centre'): ModelPoint {
  const shape = handle.snapshot().pages[0].shapes.find((item) => item.id === shapeId)!;
  return connectionPointsForShape(shape).find((point) => point.side === side)!;
}

async function dragPath(canvas: HTMLCanvasElement, from: { x: number; y: number }, to: { x: number; y: number }): Promise<void> {
  fireEvent.pointerDown(canvas, { clientX: from.x, clientY: from.y, pointerId: 1, buttons: 1, bubbles: true });
  const steps = 14;
  for (let index = 1; index <= steps; index++) {
    fireEvent.pointerMove(canvas, {
      clientX: from.x + ((to.x - from.x) * index) / steps,
      clientY: from.y + ((to.y - from.y) * index) / steps,
      pointerId: 1,
      buttons: 1,
      bubbles: true,
    });
  }
  fireEvent.pointerUp(canvas, { clientX: to.x, clientY: to.y, pointerId: 1, bubbles: true });
}

async function renderWithTwoShapes(): Promise<{ ready: Ready; canvas: HTMLCanvasElement; fromId: string; toId: string }> {
  let ready: Ready | undefined;
  render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; }} />);
  await waitFor(() => expect(ready).toBeDefined());
  const rectangle = standardShapeById('rectangle')!;
  let fromId = '';
  let toId = '';
  await act(async () => {
    fromId = ready!.handle.addShape('page:1', rectangle.draft(2, 2, 1, 1)).shapeId;
    toId = ready!.handle.addShape('page:1', rectangle.draft(5, 2, 1, 1)).shapeId;
    ready!.refresh();
  });
  const frame = ready!.handle.layoutPage(0);
  const canvas = document.querySelector('canvas[aria-label]') as HTMLCanvasElement;
  stubCanvasRect(canvas, frame.width, frame.height);
  return { ready: ready!, canvas, fromId, toId };
}

test('hovering an edge with no tool armed offers its connection points, and leaving clears them', async () => {
  const restore = stubCanvas();
  try {
    const { ready, canvas, fromId } = await renderWithTwoShapes();
    expect(canvas.style.cursor).toBe('');
    const frame = ready.handle.layoutPage(0);
    const east = pointOf(ready.handle, fromId, 'east');
    await act(async () => {
      fireEvent.pointerMove(canvas, { clientX: cssForModel(frame, east, 1).x, clientY: cssForModel(frame, east, 1).y, pointerId: 1, bubbles: true });
    });
    expect(canvas.style.cursor).toBe('crosshair');
    await act(async () => {
      fireEvent.pointerMove(canvas, { clientX: 4, clientY: 4, pointerId: 1, bubbles: true });
    });
    expect(canvas.style.cursor).toBe('');
  } finally {
    cleanup();
    restore();
  }
});

test('dragging out of a hover point with no tool armed glues a connector that follows moves and survives save', async () => {
  const restore = stubCanvas();
  try {
    const { ready, canvas, fromId, toId } = await renderWithTwoShapes();
    const frame = ready.handle.layoutPage(0);
    const east = pointOf(ready.handle, fromId, 'east');
    const west = pointOf(ready.handle, toId, 'west');
    const before = ready.handle.snapshot().pages[0].shapes.length;
    await act(async () => { await dragPath(canvas, cssForModel(frame, east, 1), cssForModel(frame, west, 1)); });
    await waitFor(() => expect(ready.handle.snapshot().pages[0].shapes.length).toBe(before + 1));
    const page = ready.handle.snapshot().pages[0];
    const connector = page.shapes[page.shapes.length - 1];
    expect(connector.name).toBe('Dynamic connector');
    const route = connectorRouteFromFrame(ready.handle.layoutPage(0), page.sourcePartPath, connector.sourceId)!;
    expect(route[0]).toEqual({ x: east.x, y: east.y });
    expect(route[route.length - 1]).toEqual({ x: west.x, y: west.y });
    await act(async () => { ready.handle.moveShape('page:1', toId, '7', '4'); ready.refresh(); });
    const moved = connectorRouteFromFrame(ready.handle.layoutPage(0), page.sourcePartPath, connector.sourceId)!;
    expect(moved[moved.length - 1]).toEqual({ x: 6.5, y: 4 });
    expect(moved[0]).toEqual({ x: east.x, y: east.y });
    const saved = ready.handle.save();
    const reopened = vsdx.openDiagram(saved, { clientId: 987001 });
    try {
      const live = reopened.snapshot().pages[0];
      const kept = live.shapes.find((shape) => shape.name === 'Dynamic connector')!;
      const keptRoute = connectorRouteFromFrame(reopened.layoutPage(0), live.sourcePartPath, kept.sourceId)!;
      expect(keptRoute[0]).toEqual({ x: east.x, y: east.y });
      expect(keptRoute[keptRoute.length - 1]).toEqual({ x: 6.5, y: 4 });
    } finally { reopened.dispose(); }
  } finally {
    cleanup();
    restore();
  }
});

test('dragging out of the centre glues dynamically and follows the source', async () => {
  const restore = stubCanvas();
  try {
    const { ready, canvas, fromId, toId } = await renderWithTwoShapes();
    const frame = ready.handle.layoutPage(0);
    const from = pointOf(ready.handle, fromId, 'centre');
    const to = pointOf(ready.handle, toId, 'centre');
    const before = ready.handle.snapshot().pages[0].shapes.length;
    await act(async () => { await dragPath(canvas, cssForModel(frame, from, 1), cssForModel(frame, to, 1)); });
    await waitFor(() => expect(ready.handle.snapshot().pages[0].shapes.length).toBe(before + 1));
    await act(async () => { ready.handle.moveShape('page:1', fromId, '3', '3'); ready.refresh(); });
    const page = ready.handle.snapshot().pages[0];
    const connector = page.shapes[page.shapes.length - 1];
    const moved = connectorRouteFromFrame(ready.handle.layoutPage(0), page.sourcePartPath, connector.sourceId)!;
    expect(moved[0]).toEqual({ x: 3, y: 3 });
    expect(moved[moved.length - 1]).toEqual({ x: to.x, y: to.y });
  } finally {
    cleanup();
    restore();
  }
});

test('dropping a hover drag on empty canvas leaves a free-ended connector', async () => {
  const restore = stubCanvas();
  try {
    const { ready, canvas, fromId } = await renderWithTwoShapes();
    const frame = ready.handle.layoutPage(0);
    const east = pointOf(ready.handle, fromId, 'east');
    const before = ready.handle.snapshot().pages[0].shapes.length;
    await act(async () => { await dragPath(canvas, cssForModel(frame, east, 1), cssForModel(frame, { x: 7, y: 4 }, 1)); });
    await waitFor(() => expect(ready.handle.snapshot().pages[0].shapes.length).toBe(before + 1));
    const page = ready.handle.snapshot().pages[0];
    const connector = page.shapes[page.shapes.length - 1];
    expect(connector.name).toBe('Dynamic connector');
    const route = connectorRouteFromFrame(ready.handle.layoutPage(0), page.sourcePartPath, connector.sourceId)!;
    expect(route[0]).toEqual({ x: east.x, y: east.y });
    expect(route[route.length - 1]).toEqual({ x: 7, y: 4 });
  } finally {
    cleanup();
    restore();
  }
});
