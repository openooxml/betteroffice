import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { mock } from 'bun:test';
import { rotationGripPosition, selectionHandlePositions, controlHandleCanvasPositions } from './interactions';

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
const originalOpenDiagram = vsdx.openDiagram;
const originalPaintPage = vsdx.paintPage;
let opens = 0;
let disposals = 0;
let paintPageOverride: typeof vsdx.paintPage | null = null;
mock.module('@betteroffice/vsdx', () => ({
  ...vsdx,
  openDiagram: (...args: Parameters<typeof originalOpenDiagram>) => {
    opens++;
    const handle = originalOpenDiagram(...args);
    const dispose = handle.dispose.bind(handle);
    handle.dispose = () => { disposals++; dispose(); };
    return handle;
  },
  paintPage: (...args: Parameters<typeof originalPaintPage>) => (paintPageOverride ?? originalPaintPage)(...args),
}));
const { VsdxEditor } = await import('./VsdxEditor');
const { useState } = await import('react');

afterEach(() => { cleanup(); paintPageOverride = null; });

/** The drawing canvas, which the rulers precede in DOM order, and its overlay. */
function drawingCanvas(container: HTMLElement): HTMLCanvasElement | null {
  return container.querySelector<HTMLCanvasElement>('canvas[aria-label]');
}

function drawingCanvases(container: HTMLElement): [HTMLCanvasElement, HTMLCanvasElement] {
  const drawing = drawingCanvas(container)!;
  return [drawing, drawing.parentElement!.querySelector<HTMLCanvasElement>('canvas[aria-hidden]')!];
}

/** Snapping is on by default; a drag that asserts its raw release turns it off in the View tab. */
function turnSnapOff(view: ReturnType<typeof render>): void {
  fireEvent.click(view.getByRole('tab', { name: 'View' }));
  fireEvent.click(view.container.querySelector('[data-view-toggle="snap"]') as HTMLElement);
}

test('does not reopen for inline fonts and a state-setting onReady callback', async () => {
  let ready: { handle: DiagramHandle } | undefined;
  opens = 0;
  disposals = 0;
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  function Host() {
    const [, setApi] = useState<unknown>();
    const [, setChanges] = useState(0);
    return <VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; setApi(api); }} onChange={() => setChanges((count) => count + 1)} />;
  }
  render(<Host />);
  await waitFor(() => expect(ready).toBeDefined());
  await act(async () => { ready!.handle.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '17'); });
  await waitFor(() => expect(ready!.handle.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('17'));
  expect(opens).toBe(1);
  cleanup();
  await waitFor(() => expect(disposals).toBe(1));
  canvasPrototype.getContext = getContext;
});

test('a parent re-rendering with a new inline onChange does not reopen the document', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  opens = 0;
  disposals = 0;
  let ready: { handle: DiagramHandle } | undefined;
  let forceRerender: (() => void) | undefined;
  function Host() {
    const [, setTick] = useState(0);
    forceRerender = () => setTick((value) => value + 1);
    return <VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; }} onChange={() => {}} />;
  }
  render(<Host />);
  await waitFor(() => expect(ready).toBeDefined());
  expect(opens).toBe(1);
  act(() => { forceRerender!(); });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(opens).toBe(1);
  expect(disposals).toBe(0);
  cleanup();
  canvasPrototype.getContext = getContext;
});

test('late collaboration opens with its seed and client ID without reopening for equivalent options', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const seedHandle = originalOpenDiagram(foundation, { clientId: 900 });
  seedHandle.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '23');
  const initialUpdate = seedHandle.encodeStateAsUpdate();
  seedHandle.dispose();
  opens = 0;
  let ready: DiagramHandle | undefined;
  let attached: vsdx.CollaborationReplica | null = null;
  const onReady = (api: { handle: DiagramHandle }) => { ready = api.handle; };
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    view.rerender(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} collaboration={{ clientId: 901, initialUpdate, onReplica: (replica) => { attached = replica; } }} />);
    await waitFor(() => expect(attached?.clientId).toBe(901));
    expect(ready!.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('23');
    await act(async () => { ready!.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '31'); });
    const opened = ready!;
    view.rerender(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} collaboration={{ clientId: 901, initialUpdate: initialUpdate.slice(), onReplica: (replica) => { attached = replica; } }} />);
    await waitFor(() => expect(attached).toBe(opened));
    expect(opens).toBe(2);
    expect(ready!.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('31');
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
  }
});

test('attaches collaboration that arrives while initialization is pending', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  const originalFontFace = globalThis.FontFace;
  const originalFonts = document.fonts;
  let finishFontLoad: (() => void) | undefined;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  class DeferredFontFace {
    constructor(_family: string, _source: ArrayBuffer, _descriptors: FontFaceDescriptors) {}
    load() { return new Promise<FontFace>((resolve) => { finishFontLoad = () => resolve(this as unknown as FontFace); }); }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: DeferredFontFace });
  Object.defineProperty(document, 'fonts', { configurable: true, value: { add: () => {} } });
  const fonts = [{ family: 'Deferred', bytes: await readFile(resolve(root, 'packages/fonts/assets/LiberationSans-Regular.ttf')) }];
  let replicas = 0;
  const view = render(<VsdxEditor file={foundation} fonts={fonts} />);
  await waitFor(() => expect(finishFontLoad).toBeDefined());
  view.rerender(<VsdxEditor file={foundation} fonts={fonts} collaboration={{ clientId: 2, onReplica: (replica) => { if (replica) replicas++; } }} />);
  finishFontLoad?.();
  await waitFor(() => expect(replicas).toBe(1));
  cleanup();
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: originalFontFace });
  Object.defineProperty(document, 'fonts', { configurable: true, value: originalFonts });
  canvasPrototype.getContext = getContext;
});

test('a stale paint does not resolve images after a newer paint has taken over', async () => {
  const originalCreateImageBitmap = globalThis.createImageBitmap;
  globalThis.createImageBitmap = async () => ({} as ImageBitmap);
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  opens = 0;
  disposals = 0;
  const mediaBytesCalls: string[] = [];
  let call = 0;
  const pending: Array<{ resolve: () => void; reject: (error: Error) => void; index: number }> = [];
  let refreshApi: (() => void) | undefined;
  const onErrors: unknown[] = [];
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => {
    refreshApi = api.refresh;
    api.handle.mediaBytes = (assetId: string) => { mediaBytesCalls.push(assetId); return new Uint8Array(0); };
  }} onError={(error) => { onErrors.push(error); }} />);
  await waitFor(() => expect(refreshApi).toBeDefined());
  drawingCanvas(view.container)!.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  paintPageOverride = (_context, _list, _dpr, _scale, options) => {
    const index = call++;
    return new Promise<void>((resolve, reject) => {
      pending[index] = {
        index,
        resolve: () => { void Promise.resolve(options?.resolveImage?.(`asset-${index}`)).then(() => resolve(), () => resolve()); },
        reject: (error) => { void Promise.resolve(options?.resolveImage?.(`asset-${index}`)).then(() => reject(error), () => reject(error)); },
      };
    });
  };
  act(() => { refreshApi!(); });
  await waitFor(() => expect(call).toBe(1));
  act(() => { refreshApi!(); });
  await waitFor(() => expect(call).toBe(2));
  pending[1].resolve();
  await waitFor(() => expect(mediaBytesCalls).toEqual(['asset-1']));
  pending[0].reject(new Error('stale paint rejected late'));
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(mediaBytesCalls).toEqual(['asset-1']);
  expect(onErrors).toEqual([]);
  globalThis.createImageBitmap = originalCreateImageBitmap;
  cleanup();
  paintPageOverride = null;
  canvasPrototype.getContext = getContext;
});



test('changing the presence provider does not reattach the same collaboration callback', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const attached: Array<vsdx.CollaborationReplica | null> = [];
  const onReplica = (replica: vsdx.CollaborationReplica | null) => { attached.push(replica); };
  const presence = { peers: [], setCursor: () => {}, onPresence: () => () => {} };
  const view = render(<VsdxEditor file={foundation} fonts={[]} collaboration={{ clientId: 920, onReplica }} />);
  try {
    await waitFor(() => expect(attached).toHaveLength(1));
    view.rerender(<VsdxEditor file={foundation} fonts={[]} collaboration={{ clientId: 920, onReplica, presence }} />);
    await act(async () => {});
    expect(attached).toHaveLength(1);
    view.rerender(<VsdxEditor file={foundation} fonts={[]} collaboration={{ clientId: 920, onReplica, presence: { ...presence } }} />);
    await act(async () => {});
    expect(attached).toHaveLength(1);
    cleanup();
    expect(attached).toHaveLength(2);
    expect(attached[1]).toBeNull();
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
  }
});

test('loading another font preserves unsaved edits and the active replica', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const font = { family: 'Arial', bytes: await readFile(resolve(root, 'packages/fonts/assets/LiberationSans-Regular.ttf')) };
  let ready: DiagramHandle | undefined;
  const onReady = (api: { handle: DiagramHandle }) => { ready = api.handle; };
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const original = ready!;
    await act(async () => { original.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '37'); });
    view.rerender(<VsdxEditor file={foundation} fonts={[font]} onReady={onReady} />);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 0)); });
    expect(ready).toBe(original);
    expect(ready!.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('37');
    expect(ready!.canUndo()).toBe(true);
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
  }
});


test('an undecodable embedded image does not blank the rest of its page', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  const originalCreateImageBitmap = globalThis.createImageBitmap;
  let fills = 0;
  const errors: Error[] = [];
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'fillText' ? () => { fills++; } : () => {}, set: () => true }) as never;
  globalThis.createImageBitmap = async () => { throw new Error('unsupported embedded image'); };
  const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/nested-groups.vsdx'));
  try {
    await act(async () => { render(<VsdxEditor file={fixture} fonts={[]} onError={(error) => { errors.push(error); }} />); });
    await waitFor(() => {
      expect(errors.length).toBeGreaterThan(0);
      expect(fills).toBeGreaterThan(0);
    });
    expect(errors[0].message).toBe('unsupported embedded image');
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
    globalThis.createImageBitmap = originalCreateImageBitmap;
  }
});


test('late collaboration cannot discard unsaved local edits or attach the wrong replica', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const seed = originalOpenDiagram(foundation, { clientId: 930 });
  const initialUpdate = seed.encodeStateAsUpdate();
  seed.dispose();
  let ready: DiagramHandle | undefined;
  const attached: Array<vsdx.CollaborationReplica | null> = [];
  const errors: Error[] = [];
  const onReady = (api: { handle: DiagramHandle }) => { ready = api.handle; };
  const onError = (error: Error) => { errors.push(error); };
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} onError={onError} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const original = ready!;
    await act(async () => { original.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '41'); });
    view.rerender(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} onError={onError} collaboration={{ clientId: 931, initialUpdate, onReplica: (replica) => { attached.push(replica); } }} />);
    await act(async () => {});
    expect(ready).toBe(original);
    expect(original.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('41');
    expect(attached).toEqual([]);
    expect(errors[errors.length - 1]?.message).toBe('Save your changes before switching collaboration sessions.');
    const reopened = originalOpenDiagram(original.save(), { clientId: 932 });
    expect(reopened.snapshot().pages[0].shapes[0].cells.find((cell) => cell.name === 'Both')?.formula).toBe('41');
    reopened.dispose();
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
  }
});

test('layout failure cannot hide unsaved edits from the session guard', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  let ready: DiagramHandle | undefined;
  const errors: Error[] = [];
  const onReady = (api: { handle: DiagramHandle }) => { ready = api.handle; };
  const onError = (error: Error) => { errors.push(error); };
  const attached: Array<vsdx.CollaborationReplica | null> = [];
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} onError={onError} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const original = ready!;
    original.layoutPage = () => { throw new Error('layout budget exceeded'); };
    await act(async () => { original.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '43'); });
    await waitFor(() => expect(errors.some(error => error.message === 'layout budget exceeded')).toBe(true));
    expect(original.snapshot().pages[0].shapes[0].cells.find(cell => cell.name === 'Both')?.formula).toBe('43');
    view.rerender(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} onError={onError} collaboration={{ clientId: 941, onReplica: replica => attached.push(replica) }} />);
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)); });
    expect(ready).toBe(original);
    expect(attached).toEqual([]);
    expect(ready!.snapshot().pages[0].shapes[0].cells.find(cell => cell.name === 'Both')?.formula).toBe('43');
  } finally {
    cleanup();
    canvasPrototype.getContext = getContext;
  }
});

test('restoring earlier bytes for the same font face registers them again', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const first = { family: 'Arial', bytes: await readFile(resolve(root, 'packages/fonts/assets/LiberationSans-Regular.ttf')) };
  const second = { family: 'Arial', bytes: await readFile(resolve(root, 'packages/fonts/assets/LiberationSerif-Regular.ttf')) };
  const registered: vsdx.VsdxFontFace[] = [];
  let ready: DiagramHandle | undefined;
  const onReady = (api: { handle: DiagramHandle }) => {
    ready = api.handle;
    const register = ready.registerFont;
    ready.registerFont = (face) => { registered.push(face); return register(face); };
  };
  const view = render(<VsdxEditor file={foundation} fonts={[first]} onReady={onReady} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const original = ready!;
    view.rerender(<VsdxEditor file={foundation} fonts={[second]} onReady={onReady} />);
    await waitFor(() => expect(registered).toHaveLength(1));
    expect(registered[0].bytes).toEqual(second.bytes);
    view.rerender(<VsdxEditor file={foundation} fonts={[first]} onReady={onReady} />);
    await waitFor(() => expect(registered).toHaveLength(2));
    expect(registered[1].bytes).toEqual(first.bytes);
    expect(ready).toBe(original);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});


test('a superseded font load cannot replace the browser font after the newer face loads', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  const originalFontFace = globalThis.FontFace;
  const originalFonts = document.fonts;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const finishes: Array<() => void> = [];
  const installed: FontFace[] = [];
  const removed: FontFace[] = [];
  class DeferredFontFace {
    load() { return new Promise<FontFace>((resolve) => { finishes.push(() => resolve(this as unknown as FontFace)); }); }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: DeferredFontFace });
  Object.defineProperty(document, 'fonts', { configurable: true, value: { add: (face: FontFace) => { installed.push(face); }, delete: (face: FontFace) => { removed.push(face); } } });
  const first = { family: 'Arial', bytes: await readFile(resolve(root, 'packages/fonts/assets/LiberationSans-Regular.ttf')) };
  const second = { family: 'Arial', bytes: await readFile(resolve(root, 'packages/fonts/assets/LiberationSerif-Regular.ttf')) };
  let ready: DiagramHandle | undefined;
  const onReady = (api: { handle: DiagramHandle }) => { ready = api.handle; };
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={onReady} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    view.rerender(<VsdxEditor file={foundation} fonts={[first]} onReady={onReady} />);
    await waitFor(() => expect(finishes).toHaveLength(1));
    view.rerender(<VsdxEditor file={foundation} fonts={[second]} onReady={onReady} />);
    await waitFor(() => expect(finishes).toHaveLength(2));
    await act(async () => { finishes[1](); });
    expect(installed).toHaveLength(1);
    await act(async () => { finishes[0](); });
    expect(installed).toHaveLength(1);
    cleanup();
    expect(removed).toEqual(installed);
  } finally {
    cleanup();
    Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: originalFontFace });
    Object.defineProperty(document, 'fonts', { configurable: true, value: originalFonts });
    canvasPrototype.getContext = getContext;
  }
});


test('a peer reordering pages preserves the active page by its identity', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: DiagramHandle | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} clientId={950} onReady={(api) => { ready = api.handle; }} />);
  let peer: DiagramHandle | undefined;
  try {
    await waitFor(() => expect(ready).toBeDefined());
    peer = originalOpenDiagram(fixture, { clientId: 951, initialUpdate: ready!.encodeStateAsUpdate() });
    peer.reorderPage('page:1', 1);
    await act(async () => { ready!.applyUpdate(peer!.encodeStateAsUpdate(ready!.encodeStateVector())); });
    expect(view.getByRole('tab', { name: 'Product map' }).getAttribute('aria-selected')).toBe('true');
    expect(view.getByRole('tab', { name: 'Release flow' }).getAttribute('aria-selected')).toBe('false');
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toBe('Page 2 of 2');
  } finally { peer?.dispose(); cleanup(); canvasPrototype.getContext = getContext; }
});


test('an update received in onReady preserves the initial active page before React commits', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let view!: ReturnType<typeof render>;
  await act(async () => { view = render(<VsdxEditor file={fixture} fonts={[]} clientId={952} onReady={({ handle }) => {
    const peer = originalOpenDiagram(fixture, { clientId: 953, initialUpdate: handle.encodeStateAsUpdate() });
    try { peer.reorderPage('page:1', 1); handle.applyUpdate(peer.encodeStateAsUpdate(handle.encodeStateVector())); }
    finally { peer.dispose(); }
  }} />); });
  try {
    await waitFor(() => expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toBe('Page 2 of 2'));
    expect(view.getByRole('tab', { name: 'Product map' }).getAttribute('aria-selected')).toBe('true');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('drag paints a live preview on the overlay and commits the release geometry', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    const resizes: string[][] = [];
    const originalResize = handle.resizeShape.bind(handle);
    handle.resizeShape = ((...args: [string, string, string, string]) => { resizes.push([...args]); return originalResize(...args); }) as DiagramHandle['resizeShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    const shapeBefore = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    const initialPinX = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinX')?.value);
    const initialPinY = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinY')?.value);
    turnSnapOff(view);
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 101, clientY: 101 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 30)); });
    expect(calls.some((entry) => entry === 'setLineDash:4,4')).toBe(false);
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    expect(calls.some((entry) => entry === 'setLineDash:4,4')).toBe(true);
    expect(calls.some((entry) => entry.startsWith('stroke:'))).toBe(true);
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(Number(moves[0][2])).toBeCloseTo(initialPinX + (120 - 100) / 96, 4);
    expect(Number(moves[0][3])).toBeCloseTo(initialPinY - (130 - 100) / 96, 4);
    calls.length = 0;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => {});
    expect(view.container.querySelector('output')).toBeNull();
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    fireEvent.pointerMove(main, { pointerId: 2, clientX: 201, clientY: 201 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 30)); });
    expect(calls.some((entry) => entry === 'setLineDash:4,4')).toBe(false);
    fireEvent.pointerUp(main, { pointerId: 2, clientX: 201, clientY: 201 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(resizes).toHaveLength(0);
    expect(view.container.querySelector('output')).toBeNull();
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('Ctrl+S saves the diagram from anywhere in the editor and never reaches the browser', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const createObjectURL = URL.createObjectURL;
  const revokeObjectURL = URL.revokeObjectURL;
  URL.createObjectURL = (() => 'blob:diagram') as typeof URL.createObjectURL;
  URL.revokeObjectURL = (() => {}) as typeof URL.revokeObjectURL;
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    let saves = 0;
    const originalSave = handle.save.bind(handle);
    handle.save = (() => { saves += 1; return originalSave(); }) as DiagramHandle['save'];
    const header = view.container.querySelector('header') as HTMLElement;
    expect(fireEvent.keyDown(header, { key: 's', ctrlKey: true })).toBe(false);
    await act(async () => {});
    expect(saves).toBe(1);
    expect(fireEvent.keyDown(header, { key: 'r', ctrlKey: true })).toBe(true);
    expect(fireEvent.keyDown(header, { key: 'l', ctrlKey: true })).toBe(true);
    expect(saves).toBe(1);
  } finally { cleanup(); canvasPrototype.getContext = getContext; URL.createObjectURL = createObjectURL; URL.revokeObjectURL = revokeObjectURL; }
});

test('a snapped drag commits the point the preview painted, and Alt keeps the raw release', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const [main, overlay] = drawingCanvases(view.container);
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({ canvas: {} }, { get: (target, key) => (key in target ? Reflect.get(target, key) : () => {}), set: () => true })) as unknown as typeof overlay.getContext);
    const pinBefore = Number(handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20')?.cells.find((cell) => cell.name === 'PinX')?.value);
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 120, clientY: 130, altKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    const snapped = Number(moves[0][2]);
    expect(snapped).not.toBeCloseTo(pinBefore + 20 / 96, 4);
    expect(Math.abs(snapped - (pinBefore + 20 / 96))).toBeLessThanOrEqual(6 / 96);
    fireEvent.pointerDown(main, { pointerId: 2, clientX: 200, clientY: 200, altKey: true });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: 220, clientY: 230, altKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: 220, clientY: 230, altKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(2);
    expect(Number(moves[1][2])).toBeCloseTo(snapped + 20 / 96, 4);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a drag returning near its start keeps the preview and commit in agreement', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    const shapeBefore = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    const initialPinX = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinX')?.value);
    const initialPinY = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinY')?.value);
    turnSnapOff(view);
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    expect(calls.some((entry) => entry.startsWith('stroke:'))).toBe(true);
    const previewPath = (entries: string[]): string[] => {
      const dash = entries.lastIndexOf('setLineDash:4,4');
      if (dash < 0) return [];
      const path: string[] = [];
      for (let index = dash + 1; index < entries.length && path.length < 4; index += 1) {
        if (entries[index].startsWith('moveTo:') || entries[index].startsWith('lineTo:')) path.push(entries[index]);
      }
      return path;
    };
    const far = previewPath(calls);
    expect(far.length).toBe(4);
    calls.length = 0;
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 101, clientY: 101 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    expect(calls.some((entry) => entry.startsWith('stroke:'))).toBe(true);
    const near = previewPath(calls);
    expect(near.length).toBe(4);
    expect(near).not.toEqual(far);
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 101, clientY: 101 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(Number(moves[0][2])).toBeCloseTo(initialPinX + (101 - 100) / 96, 4);
    expect(Number(moves[0][3])).toBeCloseTo(initialPinY - (101 - 100) / 96, 4);
    const corners = near.map((entry) => { const coords = entry.split(':')[1].split(',').map(Number); return { x: coords[0], y: coords[1] }; });
    const centre = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
    expect(centre.x).toBeCloseTo(Number(moves[0][2]) * 96, 3);
    expect(centre.y).toBeCloseTo(-Number(moves[0][3]) * 96 + 720, 3);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('concurrent pointers cannot commit or cancel each other', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    const shapeBefore = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    const initialPinX = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinX')?.value);
    const initialPinY = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinY')?.value);
    turnSnapOff(view);
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerDown(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => {});
    fireEvent.pointerCancel(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => {});
    fireEvent.lostPointerCapture(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 9, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(0);
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(Number(moves[0][2])).toBeCloseTo(initialPinX + (120 - 100) / 96, 4);
    expect(Number(moves[0][3])).toBeCloseTo(initialPinY - (130 - 100) / 96, 4);
    fireEvent.pointerUp(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    fireEvent.pointerDown(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: 220, clientY: 230 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: 220, clientY: 230 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(2);
    expect(Number(moves[1][2])).toBeCloseTo(Number(moves[0][2]) + (220 - 200) / 96, 4);
    expect(Number(moves[1][3])).toBeCloseTo(Number(moves[0][3]) - (230 - 200) / 96, 4);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

for (const formulaPins of [false, true]) test(`a handle resize with ${formulaPins ? 'formula' : 'literal'} LocPins commits one update matching its preview and undoes the whole gesture`, async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  let fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  if (formulaPins) {
    const seed = originalOpenDiagram(fixture, { clientId: 7801 });
    seed.setCellFormula('page:1', 'page:1:shape:20', { cellName: 'LocPinX' }, 'Width*0.5+0.25');
    seed.setCellFormula('page:1', 'page:1:shape:20', { cellName: 'LocPinY' }, 'Height*0.5');
    fixture = Buffer.from(seed.save());
    seed.dispose();
  }
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    const resizes: string[][] = [];
    const originalResize = handle.resizeShape.bind(handle);
    handle.resizeShape = ((...args: [string, string, string, string]) => { resizes.push([...args]); return originalResize(...args); }) as DiagramHandle['resizeShape'];
    const updates: ReturnType<typeof handle.snapshot>[] = [];
    handle.onUpdate(() => updates.push(handle.snapshot()));
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(view.container.querySelector('output')).toBeNull();
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    const page = handle.snapshot().pages[0];
    const shapeBefore = page.shapes.find((shape) => shape.id === 'page:1:shape:20');
    const width = Number(shapeBefore?.cells.find((cell) => cell.name === 'Width')?.value);
    const corners = selectionCorners(page, fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:20', hit: { kind: 'shape', shapeId: 'page:1:shape:20' } });
    expect(corners).not.toBeNull();
    const se = selectionHandlePositions(corners!).handles.se;
    calls.length = 0;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: se.x, clientY: se.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: se.x + 48, clientY: se.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    expect(calls.some((entry) => entry === 'setLineDash:4,4')).toBe(true);
    expect(calls.some((entry) => entry.startsWith('arc:'))).toBe(true);
    expect(calls.some((entry) => entry.startsWith('fillRect:'))).toBe(false);
    const previewStart = calls.lastIndexOf('setLineDash:4,4');
    const preview = calls.slice(previewStart).filter((entry) => entry.startsWith('moveTo:') || entry.startsWith('lineTo:')).slice(0, 4).map((entry) => entry.split(':')[1].split(',').map(Number));
    fireEvent.pointerUp(main, { pointerId: 2, clientX: se.x + 48, clientY: se.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(0);
    expect(resizes).toHaveLength(0);
    expect(updates).toHaveLength(1);
    const afterPage = handle.snapshot().pages[0];
    const after = afterPage.shapes.find((shape) => shape.id === 'page:1:shape:20')!;
    expect(Number(after.cells.find((cell) => cell.name === 'Width')?.value)).toBeCloseTo(width + 0.5);
    const committed = selectionCorners(afterPage, fakeFrame as never, { pageId: page.id, shapeId: after.id, hit: { kind: 'shape', shapeId: after.id } })!;
    committed.forEach((point, index) => { expect(point.x).toBeCloseTo(preview[index][0], 3); expect(point.y).toBeCloseTo(preview[index][1], 3); });
    expect(committed[3].x).toBeCloseTo(corners![3].x, 3);
    expect(committed[3].y).toBeCloseTo(corners![3].y, 3);
    await act(async () => { handle.undo(); });
    expect(handle.snapshot().pages[0].shapes.find((shape) => shape.id === after.id)).toEqual(shapeBefore);
    expect(handle.canUndo()).toBe(false);
    expect(view.container.querySelector('output')).toBeNull();
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a rotate grip drag commits the expected angle', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const formulas: Array<{ cellName: string; formula: string }> = [];
    const originalSet = handle.setCellFormula.bind(handle);
    handle.setCellFormula = ((pageId: string, shapeId: string, locator: { cellName: string }, formula: string) => {
      formulas.push({ cellName: locator.cellName, formula });
      return originalSet(pageId, shapeId, locator, formula);
    }) as DiagramHandle['setCellFormula'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const page = handle.snapshot().pages[0];
    const shape = page.shapes.find((item) => item.id === 'page:1:shape:20');
    const pinX = Number(shape?.cells.find((cell) => cell.name === 'PinX')?.value);
    const pinY = Number(shape?.cells.find((cell) => cell.name === 'PinY')?.value);
    const startAngle = Number(shape?.cells.find((cell) => cell.name === 'Angle')?.value ?? 0);
    const corners = selectionCorners(page, fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:20', hit: { kind: 'shape', shapeId: 'page:1:shape:20' } });
    const grip = rotationGripPosition(corners!, 1);
    const toModel = (canvasX: number, canvasY: number) => ({ x: canvasX / 96, y: (720 - canvasY) / 96 });
    const startModel = toModel(grip.x, grip.y);
    const startPointerAngle = Math.atan2(startModel.y - pinY, startModel.x - pinX);
    const endPointerAngle = startPointerAngle + Math.PI / 2;
    const endModelX = pinX + Math.cos(endPointerAngle) * Math.hypot(startModel.x - pinX, startModel.y - pinY);
    const endModelY = pinY + Math.sin(endPointerAngle) * Math.hypot(startModel.x - pinX, startModel.y - pinY);
    const endCanvasX = endModelX * 96;
    const endCanvasY = 720 - endModelY * 96;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: endCanvasX, clientY: endCanvasY });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: endCanvasX, clientY: endCanvasY });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const angle = formulas.find((entry) => entry.cellName === 'Angle');
    expect(angle).toBeDefined();
    expect(Number(angle!.formula)).toBeCloseTo(startAngle + Math.PI / 2, 2);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a queued rotation preview follows the latest Shift state', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const page = handle.snapshot().pages[0];
    const shape = page.shapes.find((item) => item.id === 'page:1:shape:20');
    const pinX = Number(shape?.cells.find((cell) => cell.name === 'PinX')?.value);
    const pinY = Number(shape?.cells.find((cell) => cell.name === 'PinY')?.value);
    const startAngle = Number(shape?.cells.find((cell) => cell.name === 'Angle')?.value ?? 0);
    const corners = selectionCorners(page, fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:20', hit: { kind: 'shape', shapeId: 'page:1:shape:20' } });
    const grip = rotationGripPosition(corners!, 1);
    const startModel = { x: grip.x / 96, y: (720 - grip.y) / 96 };
    const radius = Math.hypot(startModel.x - pinX, startModel.y - pinY);
    const turn = 0.35;
    const endAngle = Math.atan2(startModel.y - pinY, startModel.x - pinX) + turn;
    const endX = (pinX + Math.cos(endAngle) * radius) * 96;
    const endY = 720 - (pinY + Math.sin(endAngle) * radius) * 96;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    calls.length = 0;
    fireEvent.pointerMove(main, { pointerId: 2, clientX: endX, clientY: endY, shiftKey: true });
    fireEvent.pointerMove(main, { pointerId: 2, clientX: endX, clientY: endY, shiftKey: false });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    const previewStart = calls.lastIndexOf('setLineDash:4,4');
    expect(previewStart).toBeGreaterThanOrEqual(0);
    const preview = calls.slice(previewStart).filter((entry) => entry.startsWith('moveTo:') || entry.startsWith('lineTo:')).slice(0, 2).map((entry) => entry.split(':')[1].split(',').map(Number));
    expect(Math.atan2(preview[0][1] - preview[1][1], preview[1][0] - preview[0][0])).toBeCloseTo(startAngle + turn, 2);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a rotation-locked grip stays disabled and keeps the angle', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const errors: unknown[] = [];
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} onError={(error) => { errors.push(error); }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    const formulas: Array<{ cellName: string; formula: string }> = [];
    const originalSet = handle.setCellFormula.bind(handle);
    handle.setCellFormula = ((pageId: string, shapeId: string, locator: { cellName: string }, formula: string) => {
      formulas.push({ cellName: locator.cellName, formula });
      return originalSet(pageId, shapeId, locator, formula);
    }) as DiagramHandle['setCellFormula'];
    const { fireEvent } = await import('@testing-library/react');
    const main = drawingCanvas(view.container)!;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const pageId = handle.snapshot().pages[0].id;
    const added = await act(async () => handle.addShape(pageId, { name: 'rotation-locked', cells: [
      { locator: { cellName: 'PinX' }, formula: '4' },
      { locator: { cellName: 'PinY' }, formula: '4' },
      { locator: { cellName: 'Width' }, formula: '2' },
      { locator: { cellName: 'Height' }, formula: '1' },
      { locator: { cellName: 'Angle' }, formula: '0' },
      { locator: { cellName: 'LockRotate' }, formula: '1' },
    ] }));
    const lockedId = (added as unknown as { shapeId: string }).shapeId;
    handle.hitTest = (() => ({ kind: 'shape', shapeId: lockedId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const { selectionCorners } = await import('./VsdxEditor');
    const corners = selectionCorners(handle.snapshot().pages[0], fakeFrame as never, { pageId, shapeId: lockedId, hit: { kind: 'shape', shapeId: lockedId } });
    const grip = rotationGripPosition(corners!, 1);
    formulas.length = 0;
    errors.length = 0;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: grip.x + 40, clientY: grip.y - 40 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: grip.x + 40, clientY: grip.y - 40 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(formulas.some((entry) => entry.cellName === 'Angle')).toBe(false);
    expect(Number(handle.snapshot().pages[0].shapes.find((item) => item.id === lockedId)?.cells.find((cell) => cell.name === 'Angle')?.value ?? 0)).toBe(0);
    expect(errors.map(String)).toEqual(['Error: Shape rotation is locked and cannot be changed with handles.']);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('hovering handles sets resize and rotation cursors', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const page = handle.snapshot().pages[0];
    const corners = selectionCorners(page, fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:20', hit: { kind: 'shape', shapeId: 'page:1:shape:20' } });
    const se = selectionHandlePositions(corners!).handles.se;
    fireEvent.pointerMove(main, { pointerId: 3, clientX: se.x, clientY: se.y });
    await act(async () => {});
    expect(main.style.cursor).toContain('resize');
    const grip = rotationGripPosition(corners!, 1);
    fireEvent.pointerMove(main, { pointerId: 3, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    expect(main.style.cursor).toBe('grab');
    fireEvent.pointerMove(main, { pointerId: 3, clientX: 5, clientY: 5 });
    await act(async () => {});
    expect(main.style.cursor).toBe('');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('the overlay paints the selection frame at a zoom other than 1', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    calls.length = 0;
    fireEvent.click(view.getByRole('button', { name: 'Zoom in' }));
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 30)); });
    expect(calls.some((entry) => entry.startsWith('setTransform:1.5,0,0,1.5,0,0'))).toBe(true);
    expect(calls.some((entry) => entry.startsWith('fillRect:'))).toBe(false);
    expect(calls.some((entry) => entry.startsWith('strokeRect:'))).toBe(false);
    expect(calls.some((entry) => entry.startsWith('arc:'))).toBe(true);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a refused handle resize preserves the pin and size', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  const errors: Error[] = [];
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} onError={(error) => { errors.push(error); }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    const page = handle.snapshot().pages[0];
    const shapeBefore = page.shapes.find((shape) => shape.id === 'page:1:shape:20');
    const pinX = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinX')?.value);
    const pinY = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinY')?.value);
    await act(async () => { handle.setCellFormula(page.id, 'page:1:shape:20', { cellName: 'Width' }, 'GUARD(1)'); });
    const corners = selectionCorners(handle.snapshot().pages[0], fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:20', hit: { kind: 'shape', shapeId: 'page:1:shape:20' } });
    expect(corners).not.toBeNull();
    const se = selectionHandlePositions(corners!).handles.se;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: se.x, clientY: se.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: se.x + 48, clientY: se.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: se.x + 48, clientY: se.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const shapeAfter = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    expect(Number(shapeAfter?.cells.find((cell) => cell.name === 'PinX')?.value)).toBeCloseTo(pinX, 6);
    expect(Number(shapeAfter?.cells.find((cell) => cell.name === 'PinY')?.value)).toBeCloseTo(pinY, 6);
    expect(errors.length).toBeGreaterThan(0);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a handle resize on a move-locked shape commits neither size nor pin', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  const errors: Error[] = [];
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} onError={(error) => { errors.push(error); }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    const resizes: string[][] = [];
    const originalResize = handle.resizeShape.bind(handle);
    handle.resizeShape = ((...args: [string, string, string, string]) => { resizes.push([...args]); return originalResize(...args); }) as DiagramHandle['resizeShape'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    const pageId = handle.snapshot().pages[0].id;
    const added = await act(async () => handle.addShape(pageId, { name: 'locked', cells: [
      { locator: { cellName: 'PinX' }, formula: '1' },
      { locator: { cellName: 'PinY' }, formula: '1' },
      { locator: { cellName: 'Width' }, formula: '2' },
      { locator: { cellName: 'Height' }, formula: '1' },
      { locator: { cellName: 'LockMoveX' }, formula: '1' },
    ] }));
    const lockedId = (added as unknown as { shapeId: string }).shapeId;
    handle.hitTest = (() => ({ kind: 'shape', shapeId: lockedId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain(`selected shape ${lockedId}`);
    const shapeBefore = handle.snapshot().pages[0].shapes.find((shape) => shape.id === lockedId);
    const pinX = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinX')?.value);
    const pinY = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinY')?.value);
    const width = Number(shapeBefore?.cells.find((cell) => cell.name === 'Width')?.value);
    const height = Number(shapeBefore?.cells.find((cell) => cell.name === 'Height')?.value);
    const corners = selectionCorners(handle.snapshot().pages[0], fakeFrame as never, { pageId, shapeId: lockedId, hit: { kind: 'shape', shapeId: lockedId } });
    expect(corners).not.toBeNull();
    const se = selectionHandlePositions(corners!).handles.se;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: se.x, clientY: se.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: se.x + 48, clientY: se.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: se.x + 48, clientY: se.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(0);
    expect(resizes).toHaveLength(0);
    const shapeAfter = handle.snapshot().pages[0].shapes.find((shape) => shape.id === lockedId);
    expect(Number(shapeAfter?.cells.find((cell) => cell.name === 'Width')?.value)).toBeCloseTo(width, 6);
    expect(Number(shapeAfter?.cells.find((cell) => cell.name === 'Height')?.value)).toBeCloseTo(height, 6);
    expect(Number(shapeAfter?.cells.find((cell) => cell.name === 'PinX')?.value)).toBeCloseTo(pinX, 6);
    expect(Number(shapeAfter?.cells.find((cell) => cell.name === 'PinY')?.value)).toBeCloseTo(pinY, 6);
    expect(errors.length).toBeGreaterThan(0);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('the canvas is focusable and ArrowUp nudges PinY by the grid-off step', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: Array<ReadonlyArray<{ pageId: string; shapeId: string; xFormula: string; yFormula: string }>> = [];
    const originalMoves = handle.moveShapes.bind(handle);
    handle.moveShapes = ((arg: ReadonlyArray<{ pageId: string; shapeId: string; xFormula: string; yFormula: string }>) => { moves.push(arg); return originalMoves(arg); }) as DiagramHandle['moveShapes'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    expect(main.tabIndex).toBe(0);
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    const shapeBefore = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    const pinY = Number(shapeBefore?.cells.find((cell) => cell.name === 'PinY')?.value);
    main.focus();
    expect(document.activeElement).toBe(main);
    fireEvent.keyDown(main, { key: 'ArrowUp' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(moves[0]).toHaveLength(1);
    expect(Number(moves[0][0].yFormula)).toBeGreaterThan(pinY);
    expect(Number(moves[0][0].yFormula)).toBeCloseTo(pinY + 1 / 16, 6);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('Delete removes the selected shape and Escape cancels a drag without a commit', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    const released: number[] = [];
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = (id: number) => { released.push(id); };
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => true;
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    fireEvent.pointerDown(main, { pointerId: 2, clientX: 200, clientY: 200 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: 230, clientY: 240 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    expect(calls.some((entry) => entry === 'setLineDash:4,4')).toBe(true);
    main.focus();
    fireEvent.keyDown(main, { key: 'Escape' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(0);
    expect(released).toEqual([2]);
    expect(main.getAttribute('aria-label')).not.toContain('selected shape');
    fireEvent.pointerUp(main, { pointerId: 2, clientX: 230, clientY: 240 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(0);
    fireEvent.pointerDown(main, { pointerId: 3, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 3, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    main.focus();
    fireEvent.keyDown(main, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(false);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('typing Delete in the shapes search box keeps the selected shape', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    const search = view.getByLabelText('Search shapes') as HTMLInputElement;
    search.focus();
    fireEvent.keyDown(search, { key: 'Delete' });
    fireEvent.change(search, { target: { value: 'rect' } });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(true);
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a right-click opens the shape menu on a shape and the canvas menu on empty canvas', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => null) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvases(view.container)[0];
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const { fireEvent } = await import('@testing-library/react');
    const { en } = await import('@betteroffice/vsdx-i18n');
    expect(fireEvent.contextMenu(main, { clientX: 900, clientY: 700, button: 2 }) === false).toBe(true);
    const emptyMenu = document.querySelector('[role="menu"]');
    expect(emptyMenu === null).toBe(false);
    expect(emptyMenu?.getAttribute('aria-label')).toBe(en.contextMenu.canvasLabel);
    expect(main.getAttribute('aria-label')).not.toContain('selected shape');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(document.activeElement).toBe(main);
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    expect(fireEvent.contextMenu(main, { clientX: 100, clientY: 100, button: 2 }) === false).toBe(true);
    const menu = document.querySelector('[role="menu"]');
    expect(menu === null).toBe(false);
    expect(menu?.getAttribute('aria-label')).toBe(en.contextMenu.label);
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(document.activeElement).toBe(main);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('an invalidated selection drops the shape menu instead of reopening it on the next shape', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    const page = handle.snapshot().pages[0];
    const [first, second] = page.shapes.map((shape) => shape.id);
    expect(second).toBeDefined();
    handle.hitTest = (() => ({ kind: 'shape', shapeId: first })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvases(view.container)[0];
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.contextMenu(main, { clientX: 100, clientY: 100, button: 2 });
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    await act(async () => { handle.deleteShape(page.id, first); ready!.refresh(); });
    expect(main.getAttribute('aria-label')).not.toContain('selected shape');
    expect(document.querySelector('[role="menu"]')).toBeNull();
    handle.hitTest = (() => ({ kind: 'shape', shapeId: second })) as unknown as DiagramHandle['hitTest'];
    fireEvent.pointerDown(main, { pointerId: 3, clientX: 300, clientY: 300 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 3, clientX: 300, clientY: 300 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain(second);
    expect(document.querySelector('[role="menu"]')).toBeNull();
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('replacing the document closes the open canvas menu', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => null) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvases(view.container)[0];
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    expect(fireEvent.contextMenu(main, { clientX: 900, clientY: 700, button: 2 }) === false).toBe(true);
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    await act(async () => { view.rerender(<VsdxEditor file={fixture.slice()} fonts={[]} onReady={(api) => { ready = api; }} />); });
    await waitFor(() => expect(document.querySelector('[role="menu"]')).toBeNull());
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('switching the active page closes the open canvas menu', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => null) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvases(view.container)[0];
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    expect(fireEvent.contextMenu(main, { clientX: 900, clientY: 700, button: 2 }) === false).toBe(true);
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    const pageTablist = view.container.querySelector('[aria-label="Page tabs"]');
    expect(pageTablist).not.toBeNull();
    const pageTabs = pageTablist!.querySelectorAll('[role="tab"]');
    expect(pageTabs.length).toBeGreaterThan(1);
    await act(async () => { fireEvent.click(pageTabs[1] as HTMLElement); });
    expect(document.querySelector('[role="menu"]')).toBeNull();
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a right-click during a drag opens no menu and adds no commit', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvases(view.container)[0];
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const { fireEvent } = await import('@testing-library/react');
    const before = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    const initialPinX = Number(before?.cells.find((cell) => cell.name === 'PinX')?.value);
    const initialPinY = Number(before?.cells.find((cell) => cell.name === 'PinY')?.value);
    turnSnapOff(view);
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    expect(fireEvent.contextMenu(main, { clientX: 120, clientY: 130, button: 2 }) === false).toBe(true);
    expect(document.querySelector('[role="menu"]')).toBeNull();
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 120, clientY: 130 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(Number(moves[0][2])).toBeCloseTo(initialPinX + (120 - 100) / 96, 4);
    expect(Number(moves[0][3])).toBeCloseTo(initialPinY - (130 - 100) / 96, 4);
    fireEvent.pointerDown(main, { pointerId: 2, button: 2, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 2, button: 2, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a guarded Angle leaves the rotation grip inert without an Angle write', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  const errors: Error[] = [];
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} onError={(error) => { errors.push(error); }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const { rotationGripPosition } = await import('./interactions');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    const page = handle.snapshot().pages[0];
    await act(async () => { handle.setCellFormula(page.id, 'page:1:shape:20', { cellName: 'Angle' }, 'GUARD(0)'); });
    await act(async () => { ready!.refresh(); });
    const guarded = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    expect(guarded?.cells.find((cell) => cell.name === 'Angle')?.formula).toContain('GUARD');
    const angleWrites: string[] = [];
    const originalSet = handle.setCellFormula.bind(handle);
    handle.setCellFormula = ((pageId: string, shapeId: string, locator: { cellName: string }, formula: string) => {
      if (locator.cellName === 'Angle') angleWrites.push(formula);
      return originalSet(pageId, shapeId, locator, formula);
    }) as DiagramHandle['setCellFormula'];
    const corners = selectionCorners(handle.snapshot().pages[0], fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:20', hit: { kind: 'shape', shapeId: 'page:1:shape:20' } });
    expect(corners).not.toBeNull();
    const grip = rotationGripPosition(corners!, 1);
    fireEvent.pointerMove(main, { clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    expect(main.style.cursor).toBe('');
    const errorsBefore = errors.length;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: grip.x + 48, clientY: grip.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: grip.x + 48, clientY: grip.y + 48 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(angleWrites).toHaveLength(0);
    expect(errors.length).toBeGreaterThan(errorsBefore);
    expect(handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20')?.cells.find((cell) => cell.name === 'Angle')?.formula).toContain('GUARD');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a selected control handle paints yellow and drags through the edit session', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  let readyControl: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { readyControl = api; }} />);
  try {
    await waitFor(() => expect(readyControl).toBeDefined());
    const handle = readyControl!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    const pageId = handle.snapshot().pages[0].id;
    const added = await act(async () => handle.addShape(pageId, { name: 'adjustable', cells: [
      { locator: { cellName: 'PinX' }, formula: '1' },
      { locator: { cellName: 'PinY' }, formula: '1' },
      { locator: { cellName: 'Width' }, formula: '2' },
      { locator: { cellName: 'Height' }, formula: '1' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'X' }, formula: 'Width*0.25' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'Y' }, formula: 'Height*0.5' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'XCon' }, formula: '0' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'YCon' }, formula: '0' },
    ] }));
    const shapeId = (added as unknown as { shapeId: string }).shapeId;
    handle.hitTest = (() => ({ kind: 'shape', shapeId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { readyControl!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const calls: string[] = [];
    const overlayContext = new Proxy({ canvas: {} }, {
      get(target, key) { if (key in target) return Reflect.get(target, key); return (...args: unknown[]) => { calls.push(`${String(key)}:${args.join(',')}`); }; },
      set(target, key, value) { calls.push(`${String(key)}=${String(value)}`); Reflect.set(target, key, value); return true; },
    }) as unknown as CanvasRenderingContext2D;
    overlay.getContext = ((() => overlayContext) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(drawingCanvas(view.container)?.getAttribute('aria-label')).toContain(`selected shape ${shapeId}`);
    await act(async () => { readyControl!.refresh(); });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(calls).toContain('fillStyle=#ffeb00');
    const page = handle.snapshot().pages[0];
    const shape = page.shapes.find((entry) => entry.id === shapeId)!;
    const controlCells = shape.cells.filter((cell) => cell.locator.section === 'Control');
    expect(controlCells).toHaveLength(4);
    for (const [cellName, value] of [['X', '0.5'], ['Y', '0.5'], ['XCon', '0'], ['YCon', '0']] as const) {
      const cell = controlCells.find((entry) => entry.locator.cellName === cellName);
      expect(cell?.value).toBe(value);
    }
    const numeric = (name: string, fallback: number): number => {
      const cell = shape.cells.find((entry) => entry.name === name);
      if (!cell) return fallback;
      const parsed = Number(cell.value ?? cell.formula ?? '');
      return Number.isFinite(parsed) ? parsed : fallback;
    };
    const width = numeric('Width', 2);
    const height = numeric('Height', 1);
    const positions = controlHandleCanvasPositions(shape, {
      pin: { x: numeric('PinX', 1), y: numeric('PinY', 1) },
      locPin: { x: numeric('LocPinX', width / 2), y: numeric('LocPinY', height / 2) },
      size: { width, height },
    }, fakeFrame.paintTransform);
    expect(positions.map((position) => position.row)).toEqual(['Row_1']);
    const writes: Array<Parameters<DiagramHandle['setControlHandle']>> = [];
    const originalSetControl = handle.setControlHandle.bind(handle);
    handle.setControlHandle = ((...args: Parameters<DiagramHandle['setControlHandle']>) => {
      writes.push(args);
      return originalSetControl(...args);
    }) as DiagramHandle['setControlHandle'];
    const moved: string[][] = [];
    const originalMoveForControl = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moved.push([...args]); return originalMoveForControl(...args); }) as DiagramHandle['moveShape'];
    const anchor = positions[0].canvas;
    expect(anchor).toEqual({ x: 48, y: 624 });
    calls.length = 0;
    fireEvent.pointerMove(main, { pointerId: 99, clientX: anchor.x, clientY: anchor.y });
    await act(async () => {});
    expect(main.style.cursor).toBe('move');
    fireEvent.pointerDown(main, { pointerId: 2, clientX: anchor.x, clientY: anchor.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: anchor.x + 48, clientY: anchor.y });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: anchor.x + 48, clientY: anchor.y });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect({ writes: writes.length, moved: moved.length }).toEqual({ writes: 1, moved: 0 });
    expect(writes[0].slice(2)).toEqual(['Row_1', '1', '0.5']);
    const after = handle.snapshot().pages[0].shapes.find((entry) => entry.id === shapeId)!;
    const nextX = after.cells.find((cell) => cell.locator.section === 'Control' && cell.name === 'X');
    expect(nextX?.formula).toBe('1');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('a control drag skips guarded cells instead of partially committing', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  let readyGuarded: { handle: DiagramHandle; refresh: () => void } | undefined;
  const errors: string[] = [];
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { readyGuarded = api; }} onError={(error) => { errors.push(error.message); }} />);
  try {
    await waitFor(() => expect(readyGuarded).toBeDefined());
    const handle = readyGuarded!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    const pageId = handle.snapshot().pages[0].id;
    const added = await act(async () => handle.addShape(pageId, { name: 'adjustable', cells: [
      { locator: { cellName: 'PinX' }, formula: '1' },
      { locator: { cellName: 'PinY' }, formula: '1' },
      { locator: { cellName: 'Width' }, formula: '2' },
      { locator: { cellName: 'Height' }, formula: '1' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'X' }, formula: 'GUARD(Width*0.25)' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'Y' }, formula: 'Height*0.5' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'XCon' }, formula: '0' },
      { locator: { section: 'Control', rowName: 'Row_1', cellName: 'YCon' }, formula: '0' },
    ] }));
    const shapeId = (added as unknown as { shapeId: string }).shapeId;
    handle.hitTest = (() => ({ kind: 'shape', shapeId })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { readyGuarded!.refresh(); });
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const writes: Array<Parameters<DiagramHandle['setControlHandle']>> = [];
    const originalSetControl = handle.setControlHandle.bind(handle);
    handle.setControlHandle = ((...args: Parameters<DiagramHandle['setControlHandle']>) => {
      writes.push(args);
      return originalSetControl(...args);
    }) as DiagramHandle['setControlHandle'];
    const shape = handle.snapshot().pages[0].shapes.find((entry) => entry.id === shapeId)!;
    const positions = controlHandleCanvasPositions(shape, {
      pin: { x: 1, y: 1 },
      locPin: { x: 1, y: 0.5 },
      size: { width: 2, height: 1 },
    }, fakeFrame.paintTransform);
    expect(positions.map((position) => position.row)).toEqual(['Row_1']);
    const anchor = positions[0].canvas;
    fireEvent.pointerDown(main, { pointerId: 2, clientX: anchor.x, clientY: anchor.y });
    await act(async () => {});
    fireEvent.pointerMove(main, { pointerId: 2, clientX: anchor.x + 48, clientY: anchor.y });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
    fireEvent.pointerUp(main, { pointerId: 2, clientX: anchor.x + 48, clientY: anchor.y });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(writes.map((write) => write.slice(2))).toEqual([['Row_1', null, '0.5']]);
    expect(errors).toEqual([]);
    const after = handle.snapshot().pages[0].shapes.find((entry) => entry.id === shapeId)!;
    expect(after.cells.find((cell) => cell.locator.section === 'Control' && cell.name === 'X')?.formula).toBe('GUARD(Width*0.25)');
    expect(after.cells.find((cell) => cell.locator.section === 'Control' && cell.name === 'Y')?.formula).toBe('0.5');
    await act(async () => { handle.setCellFormula(pageId, shapeId, { section: 'Control', rowName: 'Row_1', cellName: 'Y' }, 'GUARD(Height*0.5)'); });
    await act(async () => { readyGuarded!.refresh(); });
    writes.length = 0;
    fireEvent.pointerDown(main, { pointerId: 3, clientX: anchor.x, clientY: anchor.y });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 3, clientX: anchor.x, clientY: anchor.y });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(writes).toEqual([]);
    expect(errors).toEqual([]);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('the view toggle draws page-break guides over the page', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 240, printHeight: 240, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    ready!.handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    await act(async () => { ready!.refresh(); });
    expect(view.queryByTestId('vsdx-page-breaks')).toBeNull();
    const { fireEvent } = await import('@testing-library/react');
    await act(async () => { fireEvent.click(view.container.querySelector('[data-command-id="pageBreaks"]')!); });
    expect(view.getByTestId('vsdx-page-breaks').querySelectorAll('div')).toHaveLength(5);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('the error banner dismisses and clears on the next successful edit', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/guard-format.vsdx'));
  const errors: Error[] = [];
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} onError={(error) => { errors.push(error); }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 5, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:1' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = drawingCanvases(view.container);
    const main = canvases[0] as HTMLCanvasElement;
    const overlay = canvases[1] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    overlay.getContext = ((() => new Proxy({}, { get: () => () => {}, set: () => true })) as unknown as typeof overlay.getContext);
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 1, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    const page = handle.snapshot().pages[0];
    const corners = selectionCorners(page, fakeFrame as never, { pageId: page.id, shapeId: 'page:1:shape:1', hit: { kind: 'shape', shapeId: 'page:1:shape:1' } });
    const grip = rotationGripPosition(corners!, 1);
    fireEvent.pointerDown(main, { pointerId: 2, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    const banner = view.container.querySelector('[role="alert"]');
    expect(banner?.textContent).toContain('rotation');
    const dismiss = view.getByRole('button', { name: 'Dismiss error' });
    fireEvent.click(dismiss);
    await act(async () => {});
    expect(view.container.querySelector('[role="alert"]')).toBeNull();
    fireEvent.pointerDown(main, { pointerId: 3, clientX: grip.x, clientY: grip.y });
    await act(async () => {});
    expect(view.container.querySelector('[role="alert"]')?.textContent).toContain('rotation');
    const plainId = handle.snapshot().pages[0].shapes[2].id;
    await act(async () => { handle.setCellFormula(page.id, plainId, { cellName: 'FillForegnd' }, 'RGB(9,9,9)'); ready!.refresh(); });
    await act(async () => {});
    expect(view.container.querySelector('[role="alert"]')).toBeNull();
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('Delete, undo and Escape work while editor chrome holds focus', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvas(view.container)!;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    const selectShape = async () => {
      fireEvent.pointerDown(main, { pointerId: 7, clientX: 100, clientY: 100 });
      await act(async () => {});
      fireEvent.pointerUp(main, { pointerId: 7, clientX: 100, clientY: 100 });
      await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
      expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    };
    await selectShape();
    fireEvent.click(view.getByRole('tab', { name: 'Home' }));
    const ribbonDelete = view.container.querySelector('[data-command-id="delete"]') as HTMLButtonElement;
    ribbonDelete.focus();
    expect(document.activeElement).toBe(ribbonDelete);
    fireEvent.keyDown(ribbonDelete, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(false);
    fireEvent.keyDown(ribbonDelete, { key: 'z', ctrlKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(true);
    await selectShape();
    const pageTab = view.container.querySelector('footer [role="tab"]') as HTMLButtonElement;
    pageTab.focus();
    expect(document.activeElement).toBe(pageTab);
    fireEvent.keyDown(pageTab, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(false);
    const undoAfterPageDelete = view.container.querySelector('[data-command-id="undo"]') as HTMLButtonElement;
    undoAfterPageDelete.focus();
    fireEvent.keyDown(undoAfterPageDelete, { key: 'z', ctrlKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(true);
    await selectShape();
    const escapeTarget = view.container.querySelector('footer [role="tab"]') as HTMLButtonElement;
    escapeTarget.focus();
    expect(document.activeElement).toBe(escapeTarget);
    fireEvent.keyDown(escapeTarget, { key: 'Escape' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).not.toContain('selected shape');
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('Ctrl+A selects every shape on the page and Delete removes them', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const shapeCount = handle.snapshot().pages[0].shapes.length;
    expect(shapeCount).toBeGreaterThan(1);
    const main = drawingCanvas(view.container)!;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    /** The Shape tab replaces the Home panel on selection, so the button is re-read every time. */
    const homeDelete = (): HTMLButtonElement => {
      fireEvent.click(view.getByRole('tab', { name: 'Home' }));
      return view.container.querySelector('[data-command-id="delete"]') as HTMLButtonElement;
    };
    const first = homeDelete();
    first.focus();
    expect(fireEvent.keyDown(first, { key: 'a', ctrlKey: true }) === false).toBe(true);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain(`${shapeCount} shapes selected`);
    const selected = homeDelete();
    fireEvent.keyDown(selected, { key: 'Escape' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).not.toContain('shapes selected');
    expect(handle.snapshot().pages[0].shapes).toHaveLength(shapeCount);
    main.focus();
    expect(fireEvent.keyDown(main, { key: 'a', ctrlKey: true }) === false).toBe(true);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain(`${shapeCount} shapes selected`);
    const deleteAll = homeDelete();
    deleteAll.focus();
    fireEvent.keyDown(deleteAll, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes).toHaveLength(0);
    expect(main.getAttribute('aria-label')).not.toContain('shapes selected');
    fireEvent.keyDown(homeDelete(), { key: 'z', ctrlKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.length).toBeGreaterThan(0);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('text fields keep their own undo while the ribbon uses the document history', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvas(view.container)!;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 7, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 7, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    fireEvent.keyDown(main, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(false);
    const search = view.container.querySelector('input[type="search"]') as HTMLInputElement;
    search.focus();
    expect(document.activeElement).toBe(search);
    expect(fireEvent.keyDown(search, { key: 'z', ctrlKey: true }) === false).toBe(false);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(false);
    fireEvent.click(view.getByRole('tab', { name: 'Home' }));
    const ribbonDelete = view.container.querySelector('[data-command-id="delete"]') as HTMLButtonElement;
    ribbonDelete.focus();
    fireEvent.keyDown(ribbonDelete, { key: 'z', ctrlKey: true });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(true);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});

test('Delete waits out an open context menu', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  const fixture = await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx'));
  let ready: { handle: DiagramHandle; refresh: () => void } | undefined;
  const view = render(<VsdxEditor file={fixture} fonts={[]} onReady={(api) => { ready = api; }} />);
  try {
    await waitFor(() => expect(ready).toBeDefined());
    const handle = ready!.handle;
    const fakeFrame = { contractVersion: 7, width: 960, height: 720, printWidth: 960, printHeight: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const main = drawingCanvas(view.container)!;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    const { fireEvent } = await import('@testing-library/react');
    fireEvent.pointerDown(main, { pointerId: 7, clientX: 100, clientY: 100 });
    await act(async () => {});
    fireEvent.pointerUp(main, { pointerId: 7, clientX: 100, clientY: 100 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(main.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    expect(fireEvent.contextMenu(main, { clientX: 100, clientY: 100, button: 2 }) === false).toBe(true);
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(true);
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
    expect(document.querySelector('[role="menu"]')).toBeNull();
    fireEvent.click(view.getByRole('tab', { name: 'Home' }));
    const ribbonDelete = view.container.querySelector('[data-command-id="delete"]') as HTMLButtonElement;
    ribbonDelete.focus();
    fireEvent.keyDown(ribbonDelete, { key: 'Delete' });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(handle.snapshot().pages[0].shapes.some((shape) => shape.id === 'page:1:shape:20')).toBe(false);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});


test('pans from the workspace margin and cancels only the active pointer', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  try {
    let ready: { handle: DiagramHandle } | undefined;
    const view = render(<VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; }} />);
    await waitFor(() => expect(ready).toBeDefined());
    const canvas = drawingCanvas(view.container)!;
    const workspace = view.container.querySelector('main')!;
    const surface = canvas.parentElement!;
    const captured = new Set<number>();
    workspace.setPointerCapture = (id) => { captured.add(id); };
    workspace.hasPointerCapture = (id) => captured.has(id);
    workspace.releasePointerCapture = (id) => { captured.delete(id); };
    const before = ready!.handle.snapshot();
    workspace.scrollLeft = 500; workspace.scrollTop = 600;
    fireEvent.pointerDown(surface, { button: 1, pointerId: 7, clientX: 100, clientY: 100 });
    expect(captured.has(7)).toBe(true);
    fireEvent.pointerMove(workspace, { pointerId: 8, clientX: 200, clientY: 200 });
    expect(workspace.scrollLeft).toBe(500);
    fireEvent.pointerMove(workspace, { pointerId: 7, clientX: 130, clientY: 150 });
    expect([workspace.scrollLeft, workspace.scrollTop]).toEqual([470, 550]);
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(captured.size).toBe(0);
    fireEvent.pointerMove(workspace, { pointerId: 7, clientX: 200, clientY: 200 });
    expect([workspace.scrollLeft, workspace.scrollTop]).toEqual([470, 550]);
    fireEvent.keyDown(canvas, { key: ' ' });
    fireEvent.pointerDown(surface, { button: 0, pointerId: 9, clientX: 100, clientY: 100 });
    fireEvent.pointerMove(workspace, { pointerId: 9, clientX: 110, clientY: 120 });
    expect([workspace.scrollLeft, workspace.scrollTop]).toEqual([460, 530]);
    fireEvent.blur(window);
    expect(captured.size).toBe(0);
    expect(workspace.style.cursor).toBe('');
    expect(ready!.handle.snapshot()).toEqual(before);
  } finally { cleanup(); canvasPrototype.getContext = getContext; }
});
