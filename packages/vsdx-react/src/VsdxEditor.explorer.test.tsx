import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { initWasm } from '@betteroffice/vsdx';
import { en } from '@betteroffice/vsdx-i18n';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { VsdxEditor } from './VsdxEditor';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const root = resolve(import.meta.dir, '../../..');

beforeAll(async () => initWasm(await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm'))));

/** The drawing canvas, which the rulers precede in DOM order, and its overlay. */
function drawingCanvases(container: HTMLElement): [HTMLCanvasElement, HTMLCanvasElement] {
  const drawing = container.querySelector<HTMLCanvasElement>('canvas[aria-label]')!;
  return [drawing, drawing.parentElement!.querySelector<HTMLCanvasElement>('canvas[aria-hidden]')!];
}

test('the explorer mirrors the engine snapshot and shares the canvas selection', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  try {
    const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx'));
    let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
    const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
    await waitFor(() => expect(ready).toBeDefined());
    expect(view.getByRole('tree', { name: 'Drawing Explorer' })).toBeDefined();
    expect(view.getByRole('button', { name: 'Page-1' })).toBeDefined();
    const node = view.getByRole('button', { name: 'Shape 1' }) as HTMLButtonElement;
    expect(node.closest('[role="treeitem"]')?.getAttribute('aria-selected')).toBe('false');
    await act(async () => { fireEvent.click(node); });
    const canvases = drawingCanvases(view.container);
    expect(canvases[0].getAttribute('aria-label')).toContain('page:1:shape:1');
    expect(node.closest('[role="treeitem"]')?.getAttribute('aria-selected')).toBe('true');
    expect(view.getByRole('button', { name: 'Geometry' })).toBeDefined();
    cleanup();
  } finally {
    canvasPrototype.getContext = getContext;
  }
});

test('the explorer and the shape data panel stay side by side and share one selection', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  try {
    const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx'));
    let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
    const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
    await waitFor(() => expect(ready).toBeDefined());
    expect(view.getByText(en.shapeData.title)).toBeDefined();
    expect(view.getByText(en.layersPanel.title)).toBeDefined();
    expect(view.getByText(en.shapeData.noSelection)).toBeDefined();
    await act(async () => { fireEvent.click(view.getByRole('button', { name: 'Shape 1' })); });
    expect(view.queryByText(en.shapeData.noSelection)).toBeNull();
    cleanup();
  } finally {
    canvasPrototype.getContext = getContext;
  }
});
