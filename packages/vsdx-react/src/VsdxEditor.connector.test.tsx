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

function cssForModel(frame: { width: number; height: number; paintTransform: vsdx.Affine }, model: ModelPoint, scale: number): { x: number; y: number } {
  const page = modelToPage(frame as never, model);
  return { x: page.x * scale, y: page.y * scale };
}

function centreOf(handle: DiagramHandle, shapeId: string): ModelPoint {
  const shape = handle.snapshot().pages[0].shapes.find((item) => item.id === shapeId)!;
  return connectionPointsForShape(shape).find((point) => point.side === 'centre')!;
}

/** Arms the tool through the same Insert-tab toggle a reviewer uses. */
async function armConnector(container: HTMLElement): Promise<void> {
  const insert = container.querySelector('#vsdx-ribbon-tab-insert');
  expect(insert).toBeDefined();
  fireEvent.click(insert!);
  const toggle = await waitFor(() => {
    const button = document.querySelector('button[aria-label="Connector (Alt+3)"]') as HTMLButtonElement | null;
    expect(button).toBeDefined();
    return button!;
  });
  expect(toggle.getAttribute('aria-pressed')).toBe('false');
  fireEvent.click(toggle);
  await waitFor(() => expect(toggle.getAttribute('aria-pressed')).toBe('true'));
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

test('a reviewer drag from shape A to shape B creates a connector that follows moves', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  try {
    let ready: Ready | undefined;
    const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; }} />);
    await waitFor(() => expect(ready).toBeDefined());
    const rectangle = standardShapeById('rectangle')!;
    let fromId = '';
    let toId = '';
    await act(async () => {
      fromId = ready!.handle.addShape('page:1', rectangle.draft(2, 2, 1, 1)).shapeId;
      toId = ready!.handle.addShape('page:1', rectangle.draft(5, 2, 1, 1)).shapeId;
      ready!.refresh();
    });
    await armConnector(view.container);
    const canvas = document.querySelector('canvas[aria-label]') as HTMLCanvasElement;
    expect(canvas).not.toBeNull();
    for (const scale of [1, 1.5, 0.5]) {
      await act(async () => { fireEvent.click(view.getByLabelText('Reset zoom to 100%')); });
      if (scale > 1) await act(async () => { fireEvent.click(view.getByLabelText('Zoom in')); });
      if (scale < 1) { await act(async () => { fireEvent.click(view.getByLabelText('Zoom out')); }); await act(async () => { fireEvent.click(view.getByLabelText('Zoom out')); }); }
      const frame = ready!.handle.layoutPage(0);
      stubCanvasRect(canvas, frame.width * scale, frame.height * scale);
      const from = centreOf(ready!.handle, fromId);
      const to = centreOf(ready!.handle, toId);
      const before = ready!.handle.snapshot().pages[0].shapes.length;
      await act(async () => { await dragPath(canvas, cssForModel(frame, from, scale), cssForModel(frame, to, scale)); });
      await waitFor(() => expect(ready!.handle.snapshot().pages[0].shapes.length).toBe(before + 1));
      const page = ready!.handle.snapshot().pages[0];
      const connector = page.shapes[page.shapes.length - 1];
      expect(connector.name).toBe('Dynamic connector');
      const route = connectorRouteFromFrame(ready!.handle.layoutPage(0), page.sourcePartPath, connector.sourceId)!;
      expect(route[0]).toEqual({ x: from.x, y: from.y });
      expect(route[route.length - 1]).toEqual({ x: to.x, y: to.y });
    }
    const live = ready!.handle.snapshot().pages[0];
    const connector = live.shapes[live.shapes.length - 1];
    const part = live.sourcePartPath;
    await act(async () => { ready!.handle.moveShape('page:1', fromId, '3', '3'); ready!.refresh(); });
    const movedFrom = connectorRouteFromFrame(ready!.handle.layoutPage(0), part, connector.sourceId)!;
    expect(movedFrom[0]).toEqual({ x: 3, y: 3 });
    await act(async () => { ready!.handle.moveShape('page:1', toId, '7', '4'); ready!.refresh(); });
    const movedBoth = connectorRouteFromFrame(ready!.handle.layoutPage(0), part, connector.sourceId)!;
    expect(movedBoth[movedBoth.length - 1]).toEqual({ x: 7, y: 4 });
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
  }
});
