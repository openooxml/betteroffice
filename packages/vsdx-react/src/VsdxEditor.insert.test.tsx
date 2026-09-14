import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { initWasm } from '@betteroffice/vsdx';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { VsdxEditor } from './VsdxEditor';
import { STENCIL_DRAG_MIME } from './components/shapes/ShapesPanel';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { act, cleanup, createEvent, fireEvent, render, waitFor } = await import('@testing-library/react');
const root = resolve(import.meta.dir, '../../..');
const PAGE_WIDTH = 816;
const PAGE_HEIGHT = 1056;

let foundation: Uint8Array;
let twoPages: Uint8Array;
let restoreContext: (() => void) | null = null;

beforeAll(async () => {
  const [wasm, fixture, demo] = await Promise.all([
    readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx')),
  ]);
  await initWasm(wasm);
  foundation = fixture;
  twoPages = demo;
});

afterEach(() => { cleanup(); restoreContext?.(); restoreContext = null; });

function stubCanvasContext() {
  const prototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = prototype.getContext;
  prototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  restoreContext = () => { prototype.getContext = getContext; };
}

async function editor(cssScale = 1, file = foundation) {
  stubCanvasContext();
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={file} fonts={[]} onReady={(api) => { ready = api; }} />);
  await waitFor(() => expect(ready).toBeDefined());
  await act(async () => { ready!.refresh(); });
  const main = view.container.querySelectorAll('canvas')[0] as HTMLCanvasElement;
  main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: PAGE_WIDTH * cssScale, height: PAGE_HEIGHT * cssScale, right: PAGE_WIDTH * cssScale, bottom: PAGE_HEIGHT * cssScale, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
  return { view, main, handle: ready!.handle, frame: main.parentElement as HTMLElement };
}

function stencilDrag(shapeId: string) {
  return { types: [STENCIL_DRAG_MIME], getData: (type: string) => type === STENCIL_DRAG_MIME ? shapeId : '', dropEffect: 'none', effectAllowed: 'copy' };
}

/** happy-dom has no DragEvent, so the pointer position has to be pinned onto the dispatched event. */
function dragTo(target: HTMLElement, name: 'dragOver' | 'drop', dataTransfer: object, clientX: number, clientY: number) {
  const event = createEvent[name](target, { dataTransfer });
  Object.defineProperty(event, 'clientX', { value: clientX });
  Object.defineProperty(event, 'clientY', { value: clientY });
  fireEvent(target, event);
  return event as Event & { dataTransfer: { dropEffect: string } };
}

function pins(handle: DiagramHandle, pageIndex = 0): Array<{ id: string; x: number; y: number; width: number; height: number }> {
  return handle.snapshot().pages[pageIndex].shapes.map((shape) => {
    const cell = (name: string) => Number(shape.cells.find((entry) => entry.name === name && entry.locator.section === null)?.value);
    return { id: shape.id, x: cell('PinX'), y: cell('PinY'), width: cell('Width'), height: cell('Height') };
  });
}

test('drops a stencil shape on the inches under the pointer at every zoom', async () => {
  for (const cssScale of [0.5, 1, 2]) {
    const { main, handle, frame } = await editor(cssScale);
    const before = pins(handle).length;
    const dragging = stencilDrag('ellipse');
    let over: ReturnType<typeof dragTo> | undefined;
    await act(async () => {
      over = dragTo(frame, 'dragOver', dragging, 192 * cssScale, 864 * cssScale);
      dragTo(frame, 'drop', dragging, 192 * cssScale, 864 * cssScale);
    });
    expect(over!.defaultPrevented).toBe(true);
    expect(over!.dataTransfer.dropEffect).toBe('copy');
    const added = pins(handle);
    expect(added).toHaveLength(before + 1);
    const dropped = added[added.length - 1];
    expect(dropped.x).toBeCloseTo(2, 6);
    expect(dropped.y).toBeCloseTo(2, 6);
    expect(dropped.width).toBeCloseTo(1.5, 6);
    expect(dropped.height).toBeCloseTo(1, 6);
    expect(main.getAttribute('aria-label')).toContain(dropped.id);
    expect(document.activeElement === main).toBe(true);
    cleanup();
    restoreContext?.();
    restoreContext = null;
  }
});

test('refuses a drag that carries no stencil shape', async () => {
  const { handle, frame } = await editor();
  const before = pins(handle).length;
  const foreign = { types: ['text/plain'], getData: (type: string) => type === 'text/plain' ? 'rectangle' : '', dropEffect: 'none', effectAllowed: 'copy' };
  let over: ReturnType<typeof dragTo> | undefined;
  await act(async () => {
    over = dragTo(frame, 'dragOver', foreign, 192, 864);
    dragTo(frame, 'drop', foreign, 192, 864);
  });
  expect(over!.defaultPrevented).toBe(false);
  expect(over!.dataTransfer.dropEffect).toBe('none');
  expect(pins(handle)).toHaveLength(before);
});

test('cascades repeated tile clicks instead of stacking them on the page centre', async () => {
  const { view, handle } = await editor();
  const rectangle = view.getByRole('button', { name: 'Rectangle' });
  await act(async () => { fireEvent.click(rectangle); });
  await act(async () => { fireEvent.click(rectangle); });
  const added = pins(handle).slice(-2);
  expect(added[1].x - added[0].x).toBeCloseTo(0.25, 6);
  expect(added[0].y - added[1].y).toBeCloseTo(0.25, 6);
  expect(added[0].width).toBeCloseTo(4 / 3, 6);
});

test('keeps a separate insert cascade for each page', async () => {
  const { view, handle } = await editor(1, twoPages);
  const rectangle = view.getByRole('button', { name: 'Rectangle' });
  const visited = pins(handle, 1).length;
  const insertOn = async (pageName: string) => {
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: pageName })); });
    await act(async () => { fireEvent.click(rectangle); });
  };
  await insertOn('Product map');
  await insertOn('Release flow');
  await insertOn('Product map');
  const revisited = pins(handle).slice(-2);
  expect(pins(handle, 1)).toHaveLength(visited + 1);
  expect(revisited[1].x - revisited[0].x).toBeCloseTo(0.25, 6);
  expect(revisited[0].y - revisited[1].y).toBeCloseTo(0.25, 6);
});
