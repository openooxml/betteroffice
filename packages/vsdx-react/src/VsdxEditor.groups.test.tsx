import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { initWasm } from '@betteroffice/vsdx';
import type { DiagramHandle, PagePrimitive } from '@betteroffice/vsdx';
import { VsdxEditor } from './VsdxEditor';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const root = resolve(import.meta.dir, '../../..');

beforeAll(async () => initWasm(await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm'))));

const GROUP_ID = 'page:1:shape:1';
const CHILD_ID = 'page:1:shape:1:shape:2:shape:3';
const CHILD_PART_ID = 'visio/pages/page1.xml:3';
const INSIDE_CHILD = { x: 823, y: 165 };

function childPoints(primitives: readonly PagePrimitive[]): Array<{ x: number; y: number }> {
  for (const primitive of primitives) {
    if (primitive.kind === 'shape' && primitive.id === CHILD_PART_ID) return primitive.path.filter((command) => Number.isFinite(Number(command.x)) && Number.isFinite(Number(command.y))).map((command) => ({ x: Number(command.x) * 96, y: 1056 - Number(command.y) * 96 }));
    if (primitive.kind === 'group') { const nested = childPoints(primitive.primitives); if (nested.length) return nested; }
  }
  return [];
}

function bounds(points: ReadonlyArray<{ x: number; y: number }>) {
  return { left: Math.min(...points.map((point) => point.x)), right: Math.max(...points.map((point) => point.x)), top: Math.min(...points.map((point) => point.y)), bottom: Math.max(...points.map((point) => point.y)) };
}

function pin(handle: DiagramHandle, shapeId: string, name: 'PinX' | 'PinY'): number {
  const walk = (shapes: ReturnType<DiagramHandle['snapshot']>['pages'][number]['shapes']): number | null => {
    for (const shape of shapes) {
      if (shape.id === shapeId) return Number(shape.cells.find((cell) => cell.name === name && cell.locator.section === null)?.value);
      const nested = walk(shape.children);
      if (nested !== null) return nested;
    }
    return null;
  };
  return walk(handle.snapshot().pages[0].shapes)!;
}

async function clickInsideTheGroup() {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/nested-groups.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  await waitFor(() => expect(ready).toBeDefined());
  const handle = ready!.handle;
  const hit = handle.hitTest(INSIDE_CHILD.x, INSIDE_CHILD.y);
  const child = bounds(childPoints(handle.layoutPage(0).primitives));
  await act(async () => { ready!.refresh(); });
  const canvases = view.container.querySelectorAll('canvas');
  const main = canvases[0] as HTMLCanvasElement;
  const overlay = canvases[1] as HTMLCanvasElement;
  main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 816, height: 1056, right: 816, bottom: 1056, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
  (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
  const calls: string[] = [];
  overlay.getContext = ((() => new Proxy({ canvas: {} }, {
    get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
    set(target, key, value) { Reflect.set(target, key, value); return true; },
  })) as unknown as typeof overlay.getContext);
  fireEvent.pointerDown(main, { pointerId: 1, clientX: INSIDE_CHILD.x, clientY: INSIDE_CHILD.y });
  fireEvent.pointerUp(main, { pointerId: 1, clientX: INSIDE_CHILD.x, clientY: INSIDE_CHILD.y });
  await act(async () => { await new Promise((settle) => setTimeout(settle, 20)); });
  return { handle, main, calls, hit, child, restore: () => { cleanup(); canvasPrototype.getContext = getContext; } };
}

test('a click inside a group selects the group and frames it where the group is drawn', async () => {
  const { main, calls, hit, child, restore } = await clickInsideTheGroup();
  try {
    expect(hit?.shapeId).toBe(CHILD_ID);
    expect(child.left).toBeCloseTo(820.4, 1); expect(child.right).toBeCloseTo(1091.7, 1);
    expect(child.top).toBeCloseTo(24.5, 1); expect(child.bottom).toBeCloseTo(165.5, 1);
    const label = main.getAttribute('aria-label') ?? '';
    expect(label).toContain(GROUP_ID);
    expect(label).not.toContain(CHILD_ID);
    const frame = bounds(calls.filter((entry) => entry.startsWith('moveTo:') || entry.startsWith('lineTo:')).slice(0, 4).map((entry) => { const [x, y] = entry.split(':')[1].split(',').map(Number); return { x, y }; }));
    expect(frame.left).toBeCloseTo(376.3, 0);
    expect(frame.right).toBeCloseTo(1067.1, 0);
    expect(frame.top).toBeCloseTo(-243, 0);
    expect(frame.bottom).toBeCloseTo(377.6, 0);
    expect(INSIDE_CHILD.x > frame.left && INSIDE_CHILD.x < frame.right && INSIDE_CHILD.y > frame.top && INSIDE_CHILD.y < frame.bottom).toBe(true);
  } finally { restore(); }
});

test('an arrow-key nudge moves the selected group in page space', async () => {
  const { handle, main, child, restore } = await clickInsideTheGroup();
  try {
    const groupBefore = pin(handle, GROUP_ID, 'PinY');
    const childBefore = pin(handle, CHILD_ID, 'PinY');
    await act(async () => { fireEvent.keyDown(main, { key: 'ArrowUp' }); });
    expect(pin(handle, GROUP_ID, 'PinY') - groupBefore).toBeCloseTo(1 / 96, 6);
    expect(pin(handle, CHILD_ID, 'PinY')).toBe(childBefore);
    const moved = bounds(childPoints(handle.layoutPage(0).primitives));
    expect(moved.left).toBeCloseTo(child.left, 3);
    expect(moved.top).toBeCloseTo(child.top - 1, 3);
  } finally { restore(); }
});

test('a right-click inside a group selects the group before opening the menu', async () => {
  const { main, restore } = await clickInsideTheGroup();
  try {
    await act(async () => { fireEvent.contextMenu(main, { clientX: INSIDE_CHILD.x, clientY: INSIDE_CHILD.y }); });
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    const label = main.getAttribute('aria-label') ?? '';
    expect(label).toContain(GROUP_ID);
    expect(label).not.toContain(CHILD_ID);
  } finally { restore(); }
});

test('a right-click on empty canvas drops the menu and the selection it targeted', async () => {
  const { main, restore } = await clickInsideTheGroup();
  try {
    await act(async () => { fireEvent.contextMenu(main, { clientX: INSIDE_CHILD.x, clientY: INSIDE_CHILD.y }); });
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    await act(async () => { fireEvent.contextMenu(main, { clientX: 5, clientY: 5 }); });
    expect(document.querySelector('[role="menu"]') === null).toBe(true);
    expect(main.getAttribute('aria-label') ?? '').not.toContain(GROUP_ID);
  } finally { restore(); }
});
