import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { initWasm } from '@betteroffice/vsdx';
import { en } from '@betteroffice/vsdx-i18n';
import type { Affine, DiagramHandle, PagePrimitive } from '@betteroffice/vsdx';
import { VsdxEditor } from './VsdxEditor';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const root = resolve(import.meta.dir, '../../..');

beforeAll(async () => initWasm(await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm'))));

const GROUP_ID = 'page:1:shape:1';
const CHILD_ID = 'page:1:shape:1:shape:2:shape:3';
const CHILD_PART_ID = 'visio/pages/page1.xml:3';
const INSIDE_CHILD = { x: 823, y: 165 };

const IDENTITY: Affine = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

function compose(outer: Affine, inner: Affine): Affine {
  return {
    a: outer.a * inner.a + outer.c * inner.b, b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d, d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e, f: outer.b * inner.e + outer.d * inner.f + outer.f,
  };
}

/** Page-space canvas points for the group child, composed through its ancestor transforms. */
function childPoints(primitives: readonly PagePrimitive[], ancestors: Affine = IDENTITY): Array<{ x: number; y: number }> {
  for (const primitive of primitives) {
    const local = compose(ancestors, ('transform' in primitive ? primitive.transform : undefined) ?? IDENTITY);
    if (primitive.kind === 'shape' && primitive.id === CHILD_PART_ID) {
      return primitive.path
        .filter((command) => Number.isFinite(Number(command.x)) && Number.isFinite(Number(command.y)))
        .map((command) => {
          const x = Number(command.x), y = Number(command.y);
          return { x: (local.a * x + local.c * y + local.e) * 96, y: 1056 - (local.b * x + local.d * y + local.f) * 96 };
        });
    }
    if (primitive.kind === 'group') { const nested = childPoints(primitive.primitives, local); if (nested.length) return nested; }
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
  fireEvent.click(view.getByRole('tab', { name: en.ribbon.tabs.view }));
  fireEvent.click(view.container.querySelector('[data-view-toggle="grid"]') as HTMLElement);
  const canvases = drawingCanvases(view.container);
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

/** The drawing canvas, which the rulers precede in DOM order, and its overlay. */
function drawingCanvases(container: HTMLElement): [HTMLCanvasElement, HTMLCanvasElement] {
  const drawing = container.querySelector<HTMLCanvasElement>('canvas[aria-label]')!;
  return [drawing, drawing.parentElement!.querySelector<HTMLCanvasElement>('canvas[aria-hidden]')!];
}

test('a click inside a group selects the group and frames it where the group is drawn', async () => {
  const { main, calls, hit, child, restore } = await clickInsideTheGroup();
  try {
    expect(hit?.shapeId).toBe(CHILD_ID);
    expect(child.left).toBeCloseTo(764.7, 1); expect(child.right).toBeCloseTo(882.3, 1);
    expect(child.top).toBeCloseTo(60.6, 1); expect(child.bottom).toBeCloseTo(178.2, 1);
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
    expect(pin(handle, GROUP_ID, 'PinY') - groupBefore).toBeCloseTo(1 / 16, 6);
    expect(pin(handle, CHILD_ID, 'PinY')).toBe(childBefore);
    const moved = bounds(childPoints(handle.layoutPage(0).primitives));
    expect(moved.left).toBeCloseTo(child.left, 3);
    expect(moved.top).toBeCloseTo(child.top - 96 / 16, 3);
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

test('a right-click on empty canvas swaps in the canvas menu and drops the selection', async () => {
  const { main, restore } = await clickInsideTheGroup();
  try {
    await act(async () => { fireEvent.contextMenu(main, { clientX: INSIDE_CHILD.x, clientY: INSIDE_CHILD.y }); });
    expect(document.querySelector(`[role="menu"][aria-label="${en.contextMenu.label}"]`)).not.toBeNull();
    await act(async () => { fireEvent.contextMenu(main, { clientX: 5, clientY: 5 }); });
    expect(document.querySelector(`[role="menu"][aria-label="${en.contextMenu.label}"]`)).toBeNull();
    expect(document.querySelector(`[role="menu"][aria-label="${en.contextMenu.canvasLabel}"]`)).not.toBeNull();
    expect(main.getAttribute('aria-label') ?? '').not.toContain(GROUP_ID);
  } finally { restore(); }
});
