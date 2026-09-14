import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { mock } from 'bun:test';
import { rotationGripPosition, selectionHandlePositions } from './interactions';

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

const { act, cleanup, render, waitFor } = await import('@testing-library/react');
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
  view.container.querySelector('canvas')!.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
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
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toBe('Page 2 of 2');
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
    await waitFor(() => expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toBe('Page 2 of 2'));
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    const resizes: string[][] = [];
    const originalResize = handle.resizeShape.bind(handle);
    handle.resizeShape = ((...args: [string, string, string, string]) => { resizes.push([...args]); return originalResize(...args); }) as DiagramHandle['resizeShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
    fireEvent.pointerMove(main, { pointerId: 2, clientX: 201, clientY: 201 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 30)); });
    expect(calls.some((entry) => entry === 'setLineDash:4,4')).toBe(false);
    fireEvent.pointerUp(main, { pointerId: 2, clientX: 201, clientY: 201 });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(moves).toHaveLength(1);
    expect(resizes).toHaveLength(0);
    expect(view.container.querySelector('output')).toBeNull();
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
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
    const canvases = view.container.querySelectorAll('canvas');
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
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
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
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
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
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const { selectionCorners } = await import('./VsdxEditor');
    const canvases = view.container.querySelectorAll('canvas');
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
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toContain('selected shape page:1:shape:20');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
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
    const canvases = view.container.querySelectorAll('canvas');
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
    expect(view.container.querySelector('canvas')?.getAttribute('aria-label')).toContain(`selected shape ${lockedId}`);
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

test('the canvas is focusable and ArrowUp nudges PinY by one screen pixel', async () => {
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
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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
    expect(Number(moves[0][3])).toBeGreaterThan(pinY);
    expect(Number(moves[0][3])).toBeCloseTo(pinY + 1 / 96, 6);
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    await act(async () => { ready!.refresh(); });
    const canvases = view.container.querySelectorAll('canvas');
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

test('a right-click opens the shape menu with a selection and empty canvas opens nothing', async () => {
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
    const main = view.container.querySelectorAll('canvas')[0] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const { fireEvent } = await import('@testing-library/react');
    const { en } = await import('@betteroffice/vsdx-i18n');
    expect(fireEvent.contextMenu(main, { clientX: 900, clientY: 700, button: 2 }) === false).toBe(true);
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(main.getAttribute('aria-label')).not.toContain('selected shape');
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
    const fakeFrame = { contractVersion: 4, width: 960, height: 720, paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 720 }, primitives: [] };
    handle.layoutPage = (() => fakeFrame) as unknown as DiagramHandle['layoutPage'];
    handle.hitTest = (() => ({ kind: 'shape', shapeId: 'page:1:shape:20' })) as unknown as DiagramHandle['hitTest'];
    const moves: string[][] = [];
    const originalMove = handle.moveShape.bind(handle);
    handle.moveShape = ((...args: [string, string, string, string]) => { moves.push([...args]); return originalMove(...args); }) as DiagramHandle['moveShape'];
    await act(async () => { ready!.refresh(); });
    const main = view.container.querySelectorAll('canvas')[0] as HTMLCanvasElement;
    main.getBoundingClientRect = (() => ({ left: 0, top: 0, width: 960, height: 720, right: 960, bottom: 720, x: 0, y: 0, toJSON: () => ({}) })) as unknown as typeof main.getBoundingClientRect;
    (main as unknown as { setPointerCapture: (id: number) => void }).setPointerCapture = () => {};
    (main as unknown as { releasePointerCapture: (id: number) => void }).releasePointerCapture = () => {};
    (main as unknown as { hasPointerCapture: (id: number) => boolean }).hasPointerCapture = () => false;
    const { fireEvent } = await import('@testing-library/react');
    const before = handle.snapshot().pages[0].shapes.find((shape) => shape.id === 'page:1:shape:20');
    const initialPinX = Number(before?.cells.find((cell) => cell.name === 'PinX')?.value);
    const initialPinY = Number(before?.cells.find((cell) => cell.name === 'PinY')?.value);
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
