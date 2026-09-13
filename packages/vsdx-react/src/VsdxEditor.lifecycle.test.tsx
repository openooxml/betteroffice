import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { mock } from 'bun:test';

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
