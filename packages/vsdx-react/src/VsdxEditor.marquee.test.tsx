import { afterEach, beforeAll, beforeEach, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';
import type { DiagramHandle, ModelPoint, PageDisplayList, PageSnapshot } from '@betteroffice/vsdx';
import { findShapePlacement, numericCellValue } from './components/ribbon/commands';
import { normalizeMarquee } from './interactions';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const { VsdxEditor, marqueeEnclosedShapes, selectionCorners } = await import('./VsdxEditor');

const root = resolve(import.meta.dir, '../../..');
let source: Uint8Array;

beforeAll(async () => {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/grouped-glue.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  source = fixture;
});

const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
const realGetContext = canvasPrototype.getContext;
const overlayCalls: string[] = [];
const mainCalls: string[] = [];

function recorder(calls: string[]): CanvasRenderingContext2D {
  return new Proxy({}, {
    get: (_target, key) => {
      if (key === 'measureText') return () => ({ width: 0 });
      return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); };
    },
    set: () => true,
  }) as unknown as CanvasRenderingContext2D;
}

beforeEach(() => {
  canvasPrototype.getContext = function (this: HTMLCanvasElement) {
    return recorder(this.hasAttribute('aria-hidden') ? overlayCalls : mainCalls);
  } as never;
});

afterEach(() => {
  cleanup();
  canvasPrototype.getContext = realGetContext;
  overlayCalls.length = 0;
  mainCalls.length = 0;
});

/** The rulers add their own canvases, so the drawing pair is found by label, not by index. */
function canvases(): { main: HTMLCanvasElement; overlay: HTMLCanvasElement } {
  const main = document.querySelector<HTMLCanvasElement>('canvas[aria-label]');
  expect(main).not.toBeNull();
  const overlay = main!.parentElement!.querySelector<HTMLCanvasElement>('canvas[aria-hidden]');
  expect(overlay).not.toBeNull();
  return { main: main!, overlay: overlay! };
}

function label(): string | null {
  return canvases().main.getAttribute('aria-label');
}

function expectedLabel(ids: readonly string[]): string {
  if (ids.length === 0) return 'Page 1 of 1';
  if (ids.length === 1) return `Page 1 of 1; selected shape ${ids[0]}`;
  return `Page 1 of 1; ${ids.length} shapes selected`;
}

async function readyEditor(): Promise<{ main: HTMLCanvasElement; handle: DiagramHandle }> {
  let handle: DiagramHandle | undefined;
  render(<VsdxEditor file={source} fonts={[]} onReady={(api) => { handle = api.handle; }} />);
  await waitFor(() => expect(handle).toBeDefined());
  await waitFor(() => expect(label()).toBe('Page 1 of 1'));
  const { main } = canvases();
  const frame = handle!.layoutPage(0);
  main.getBoundingClientRect = () => ({ left: 0, top: 0, width: frame.width, height: frame.height, right: frame.width, bottom: frame.height, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect;
  main.setPointerCapture = () => {};
  main.releasePointerCapture = () => {};
  main.hasPointerCapture = () => false;
  return { main, handle: handle! };
}

/** Click selection resolves a hit to its top-level shape, so the marquee must agree. */
function topLevelId(page: PageSnapshot, shapeId: string): string {
  const top = page.shapes.find((shape) => shape.id === shapeId || findShapePlacement(shape.children, shapeId) !== null);
  return top?.id ?? shapeId;
}

function selectableQuads(page: PageSnapshot, frame: PageDisplayList): Array<{ id: string; corners: ModelPoint[] }> {
  const quads: Array<{ id: string; corners: ModelPoint[] }> = [];
  for (const shape of page.shapes) {
    try {
      const corners = selectionCorners(page, frame, { pageId: page.id, shapeId: shape.id, hit: { kind: 'shape', shapeId: shape.id } });
      if (corners) quads.push({ id: shape.id, corners });
    } catch { void 0; }
  }
  return quads;
}

function down(main: HTMLCanvasElement, at: [number, number], shift = false) {
  fireEvent.pointerDown(main, { button: 0, pointerId: 1, clientX: at[0], clientY: at[1], shiftKey: shift });
}
function move(main: HTMLCanvasElement, to: [number, number], shift = false) {
  fireEvent.pointerMove(main, { button: 0, pointerId: 1, clientX: to[0], clientY: to[1], shiftKey: shift });
}
function strokeRects(calls: string[]): number[][] {
  return calls.filter((entry) => entry.startsWith('strokeRect:')).map((entry) => entry.slice('strokeRect:'.length).split(',').map(Number));
}
function matchesRect(rect: number[], wanted: number[]): boolean {
  return rect.length === 4 && rect.every((value, index) => Math.abs(value - wanted[index]) < 1e-9);
}
function up(main: HTMLCanvasElement, to: [number, number], shift = false) {
  fireEvent.pointerUp(main, { button: 0, pointerId: 1, clientX: to[0], clientY: to[1], shiftKey: shift });
}

test('dragging from empty canvas selects fully enclosed shapes and paints on the overlay', async () => {
  const { main, handle } = await readyEditor();
  const snapshot = handle.snapshot();
  const page = snapshot.pages[0];
  const frame = handle.layoutPage(0);
  const quads = selectableQuads(page, frame);
  expect(quads.length).toBeGreaterThanOrEqual(1);
  const { id: first, corners } = quads[0];
  const pad = 12;
  const from: [number, number] = [Math.min(...corners.map((corner) => corner.x)) - pad, Math.min(...corners.map((corner) => corner.y)) - pad];
  const to: [number, number] = [Math.max(...corners.map((corner) => corner.x)) + pad, Math.max(...corners.map((corner) => corner.y)) + pad];
  const wanted = marqueeEnclosedShapes(page, frame, normalizeMarquee({ x: from[0], y: from[1] }, { x: to[0], y: to[1] })).map((item) => item.shapeId);
  expect(wanted).toContain(first);
  down(main, from);
  move(main, to);
  const wantedRect = [from[0], from[1], to[0] - from[0], to[1] - from[1]];
  await waitFor(() => {
    expect(strokeRects(overlayCalls).some((rect) => matchesRect(rect, wantedRect))).toBe(true);
  });
  expect(strokeRects(mainCalls).some((rect) => matchesRect(rect, wantedRect))).toBe(false);
  await act(async () => { up(main, to); });
  await waitFor(() => expect(label()).toBe(expectedLabel(wanted)));
});

test('a marquee that only clips a shape leaves it unselected', async () => {
  const { main, handle } = await readyEditor();
  const snapshot = handle.snapshot();
  const page = snapshot.pages[0];
  const frame = handle.layoutPage(0);
  const quads = selectableQuads(page, frame);
  expect(quads.length).toBeGreaterThanOrEqual(1);
  const { id: first, corners } = quads[0];
  const xs = corners.map((corner) => corner.x).sort((left, right) => left - right);
  const ys = corners.map((corner) => corner.y).sort((left, right) => left - right);
  const from: [number, number] = [xs[0] - 4, ys[0] - 4];
  const to: [number, number] = [(xs[0] + xs[3]) / 2, (ys[0] + ys[3]) / 2];
  const wanted = marqueeEnclosedShapes(page, frame, normalizeMarquee({ x: from[0], y: from[1] }, { x: to[0], y: to[1] })).map((item) => item.shapeId);
  expect(wanted).not.toContain(first);
  await act(async () => { down(main, from); move(main, to); up(main, to); });
  await waitFor(() => expect(label()).toBe(expectedLabel(wanted)));
});

test('shift-drag adds to the existing selection instead of replacing it', async () => {
  const { main, handle } = await readyEditor();
  const snapshot = handle.snapshot();
  const page = snapshot.pages[0];
  const frame = handle.layoutPage(0);
  const quads = selectableQuads(page, frame);
  expect(quads.length).toBeGreaterThanOrEqual(2);
  const tight = (corners: ModelPoint[], pad: number): [[number, number], [number, number]] => [
    [Math.min(...corners.map((corner) => corner.x)) - pad, Math.min(...corners.map((corner) => corner.y)) - pad],
    [Math.max(...corners.map((corner) => corner.x)) + pad, Math.max(...corners.map((corner) => corner.y)) + pad],
  ];
  const [firstFrom, firstTo] = tight(quads[0].corners, 12);
  const first = marqueeEnclosedShapes(page, frame, normalizeMarquee({ x: firstFrom[0], y: firstFrom[1] }, { x: firstTo[0], y: firstTo[1] })).map((item) => item.shapeId);
  expect(first.length).toBeGreaterThan(0);
  await act(async () => { down(main, firstFrom); move(main, firstTo); up(main, firstTo); });
  await waitFor(() => expect(label()).toBe(expectedLabel(first)));
  const [secondFrom, secondTo] = tight(quads[1].corners, 12);
  const second = marqueeEnclosedShapes(page, frame, normalizeMarquee({ x: secondFrom[0], y: secondFrom[1] }, { x: secondTo[0], y: secondTo[1] })).map((item) => item.shapeId);
  const combined = [...first, ...second.filter((id) => !first.includes(id))];
  await act(async () => { down(main, secondFrom, true); move(main, secondTo, true); up(main, secondTo, true); });
  await waitFor(() => expect(label()).toBe(expectedLabel(combined)));
});

test('escape during the drag cancels with nothing selected', async () => {
  const { main } = await readyEditor();
  down(main, [10, 10]);
  move(main, [200, 200]);
  await act(async () => { fireEvent.keyDown(main, { key: 'Escape' }); });
  up(main, [200, 200]);
  await waitFor(() => expect(label()).toBe('Page 1 of 1'));
});

test('a drag that starts on a shape moves it instead of starting a marquee', async () => {
  const { main, handle } = await readyEditor();
  const before = handle.snapshot();
  const page = before.pages[0];
  const frame = handle.layoutPage(0);
  const quads = selectableQuads(page, frame);
  expect(quads.length).toBeGreaterThanOrEqual(1);
  const target = quads.map((quad) => {
    const quadCentre: [number, number] = [(quad.corners[0].x + quad.corners[2].x) / 2, (quad.corners[0].y + quad.corners[2].y) / 2];
    return { centre: quadCentre, hit: handle.hitTest(quadCentre[0], quadCentre[1]) };
  }).find((entry) => entry.hit !== null);
  expect(target).toBeDefined();
  const { centre, hit } = target!;
  expect(hit).not.toBeNull();
  const selectedId = topLevelId(page, hit!.shapeId);
  const placement = findShapePlacement(page.shapes, selectedId);
  expect(placement).not.toBeNull();
  const pinBefore = numericCellValue(placement!.shape, 'PinX');
  down(main, centre);
  await waitFor(() => expect(label()).toBe(`Page 1 of 1; selected shape ${selectedId}`));
  await act(async () => { move(main, [centre[0] + 48, centre[1] + 48]); up(main, [centre[0] + 48, centre[1] + 48]); });
  const after = handle.snapshot();
  const moved = findShapePlacement(after.pages[0].shapes, selectedId);
  expect(moved).not.toBeNull();
  expect(numericCellValue(moved!.shape, 'PinX')).not.toBe(pinBefore);
  expect(label()).toBe(`Page 1 of 1; selected shape ${selectedId}`);
});
