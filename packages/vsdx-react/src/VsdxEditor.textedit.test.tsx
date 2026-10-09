import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';
import type { DiagramHandle } from '@betteroffice/vsdx';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const root = resolve(import.meta.dir, '../../..');

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const { VsdxEditor } = await import('./VsdxEditor');

test('a refused text commit keeps the editor open with the draft intact', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/text-lock.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const pageId = handle.snapshot().pages[0].id;
    const shapeId = handle.snapshot().pages[0].shapes[0].id;
    handle.hitTest = (() => ({ kind: 'shape', shapeId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = view.container.querySelector('canvas[aria-label]') as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    await act(async () => { fireEvent.doubleClick(main, { clientX: 100, clientY: 100 }); });
    const box = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(box).not.toBeNull();
    expect(box!.value).toBe('locked text');
    await act(async () => { fireEvent.change(box!, { target: { value: 'locked text edited' } }); });
    await act(async () => { fireEvent.keyDown(box!, { key: 'Escape' }); });
    const kept = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(kept).not.toBeNull();
    expect(kept!.value).toBe('locked text edited');
    expect(view.container.querySelector('[role="alert"]')?.textContent).toContain('protects');
    expect(handle.shapeText(pageId, shapeId)).toBe('locked text');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('clicking away from a refused text commit keeps the draft', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/text-lock.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const pageId = handle.snapshot().pages[0].id;
    const shapeId = handle.snapshot().pages[0].shapes[0].id;
    handle.hitTest = (() => ({ kind: 'shape', shapeId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = view.container.querySelector('canvas[aria-label]') as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    await act(async () => { fireEvent.doubleClick(main, { clientX: 100, clientY: 100 }); });
    const box = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(box).not.toBeNull();
    await act(async () => { fireEvent.change(box!, { target: { value: 'locked text edited' } }); });

    handle.hitTest = (() => null) as unknown as DiagramHandle['hitTest'];
    await act(async () => { fireEvent.pointerDown(main, { pointerId: 1, clientX: 400, clientY: 400 }); });

    const kept = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(kept).not.toBeNull();
    expect(kept!.value).toBe('locked text edited');
    expect(handle.shapeText(pageId, shapeId)).toBe('locked text');
    await waitFor(() => expect(document.activeElement).toBe(kept));

    await act(async () => { fireEvent.keyDown(kept!, { key: 'Escape' }); });
    expect(view.container.querySelector('textarea')).toBeNull();
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a refused edit does not survive the shape becoming unreachable', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/text-lock.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const pageId = handle.snapshot().pages[0].id;
    const shapeId = handle.snapshot().pages[0].shapes[0].id;
    handle.hitTest = (() => ({ kind: 'shape', shapeId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = view.container.querySelector('canvas[aria-label]') as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    await act(async () => { fireEvent.doubleClick(main, { clientX: 100, clientY: 100 }); });
    const box = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(box).not.toBeNull();
    await act(async () => { fireEvent.change(box!, { target: { value: 'locked text edited' } }); });
    await act(async () => { fireEvent.keyDown(box!, { key: 'Escape' }); });
    expect(view.container.querySelector('textarea')).not.toBeNull();

    const other = handle.snapshot().pages[0].shapes.find(shape => shape.id !== shapeId)!;
    expect(other).toBeDefined();

    handle.deleteShape(pageId, shapeId);
    await act(async () => { ready!.refresh(); });
    expect(view.container.querySelector('textarea')).toBeNull();

    handle.hitTest = (() => ({ kind: 'shape', shapeId: other.id })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { fireEvent.doubleClick(main, { clientX: 100, clientY: 100 }); });
    const reopened = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(reopened).not.toBeNull();
    expect(reopened!.value).toBe('free text');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a refused edit does not survive its layer being hidden', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/text-lock.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const page = handle.snapshot().pages[0];
    const shapeId = page.shapes[0].id;
    handle.hitTest = (() => ({ kind: 'shape', shapeId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = view.container.querySelector('canvas[aria-label]') as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    await act(async () => { fireEvent.doubleClick(main, { clientX: 100, clientY: 100 }); });
    const box = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(box).not.toBeNull();
    await act(async () => { fireEvent.change(box!, { target: { value: 'locked text edited' } }); });
    await act(async () => { fireEvent.keyDown(box!, { key: 'Escape' }); });
    expect(view.container.querySelector('textarea')).not.toBeNull();

    handle.setLayerVisible(page.sourcePartPath, 0, false);
    await act(async () => { ready!.refresh(); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === shapeId)).toBe(true);
    expect(view.container.querySelector('textarea')).toBeNull();

    const other = handle.snapshot().pages[0].shapes.find((shape) => shape.id !== shapeId)!;
    handle.hitTest = (() => ({ kind: 'shape', shapeId: other.id })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { fireEvent.doubleClick(main, { clientX: 100, clientY: 100 }); });
    const reopened = view.container.querySelector('textarea') as HTMLTextAreaElement | null;
    expect(reopened).not.toBeNull();
    expect(reopened!.value).toBe('free text');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});
