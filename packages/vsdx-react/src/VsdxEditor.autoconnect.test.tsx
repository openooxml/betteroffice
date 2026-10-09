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
const { autoConnectArrowCenter, autoConnectArrowsForShape, connectionPointsForShape, connectorRouteFromFrame, modelToPage } = await import('./connector');

afterEach(() => { cleanup(); });

interface Ready { handle: DiagramHandle; refresh: () => void; }

function stubCanvasRect(canvas: HTMLCanvasElement, width: number, height: number): void {
  canvas.getBoundingClientRect = () => ({ x: 0, y: 0, left: 0, top: 0, right: width, bottom: height, width, height, toJSON: () => {} }) as DOMRect;
}

function cssForModel(frame: { width: number; height: number; paintTransform: vsdx.Affine }, model: ModelPoint, scale: number): { x: number; y: number } {
  const page = modelToPage(frame as never, model);
  return { x: page.x * scale, y: page.y * scale };
}

function centreOf(handle: DiagramHandle, shapeId: string): ModelPoint {
  const shape = handle.snapshot().pages[0].shapes.find((item) => item.id === shapeId)!;
  return connectionPointsForShape(shape).find((point) => point.side === 'centre')!;
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

function stubCanvas(): () => void {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  return () => { canvasPrototype.getContext = getContext; };
}

async function hoverEastArrow(canvas: HTMLCanvasElement, ready: Ready, shapeId: string): Promise<{ x: number; y: number }> {
  const frame = ready.handle.layoutPage(0);
  const from = centreOf(ready.handle, shapeId);
  const overShape = cssForModel(frame, from, 1);
  await act(async () => {
    fireEvent.pointerMove(canvas, { clientX: overShape.x, clientY: overShape.y, pointerId: 1, bubbles: true });
  });
  const shape = ready.handle.snapshot().pages[0].shapes.find((item) => item.id === shapeId)!;
  const east = autoConnectArrowsForShape(shape).find((arrow) => arrow.side === 'east')!;
  const css = cssForModel(frame, autoConnectArrowCenter(east, frame, 1), 1);
  await act(async () => {
    const steps = 12;
    for (let index = 1; index <= steps; index++) {
      fireEvent.pointerMove(canvas, {
        clientX: overShape.x + ((css.x - overShape.x) * index) / steps,
        clientY: overShape.y + ((css.y - overShape.y) * index) / steps,
        pointerId: 1,
        bubbles: true,
      });
    }
  });
  await waitFor(() => expect(document.querySelector('[role="menu"]')).not.toBeNull());
  return css;
}

test('a reviewer drag from an edge arrow draws a connector glued to that edge', async () => {
  const restore = stubCanvas();
  try {
    const { ready, canvas, fromId, toId } = await renderWithTwoShapes();
    const arrow = await hoverEastArrow(canvas, ready, fromId);
    const to = centreOf(ready.handle, toId);
    const frame = ready.handle.layoutPage(0);
    const shape = ready.handle.snapshot().pages[0].shapes.find((item) => item.id === fromId)!;
    const east = autoConnectArrowsForShape(shape).find((item) => item.side === 'east')!;
    const before = ready.handle.snapshot().pages[0].shapes.length;
    await act(async () => { await dragPath(canvas, arrow, cssForModel(frame, to, 1)); });
    await waitFor(() => expect(ready.handle.snapshot().pages[0].shapes.length).toBe(before + 1));
    const page = ready.handle.snapshot().pages[0];
    const connector = page.shapes[page.shapes.length - 1];
    expect(connector.name).toBe('Dynamic connector');
    const route = connectorRouteFromFrame(ready.handle.layoutPage(0), page.sourcePartPath, connector.sourceId)!;
    expect(route[0]).toEqual({ x: east.point.x, y: east.point.y });
    expect(route[route.length - 1]).toEqual({ x: to.x, y: to.y });
    expect(document.querySelector('[role="menu"]')).toBeNull();
  } finally {
    cleanup();
    restore();
  }
});

test('clicking a quick shape inserts it connected as a single undo step', async () => {
  const restore = stubCanvas();
  try {
    const { ready, canvas, fromId } = await renderWithTwoShapes();
    await hoverEastArrow(canvas, ready, fromId);
    const items = document.querySelectorAll('[role="menuitem"]');
    expect(items.length).toBe(5);
    const before = ready.handle.snapshot().pages[0].shapes.length;
    await act(async () => { fireEvent.click(items[0]); });
    await waitFor(() => expect(ready.handle.snapshot().pages[0].shapes.length).toBe(before + 2));
    const page = ready.handle.snapshot().pages[0];
    const names = page.shapes.slice(-2).map((shape) => shape.name);
    expect(names).toContain('rectangle');
    expect(names).toContain('Dynamic connector');
    const added = page.shapes[page.shapes.length - 2];
    const addedCentre = connectionPointsForShape(added).find((point) => point.side === 'centre')!;
    const sourceCentre = centreOf(ready.handle, fromId);
    expect(addedCentre.x).toBeGreaterThan(sourceCentre.x);
    expect(addedCentre.y).toBeCloseTo(sourceCentre.y, 10);
    await act(async () => { ready.handle.undo(); ready.refresh(); });
    expect(ready.handle.snapshot().pages[0].shapes.length).toBe(before);
  } finally {
    cleanup();
    restore();
  }
});
