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

test('attaches collaboration that arrives after the file opens', async () => {
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  let connected = false;
  let replicas = 0;
  const view = render(<VsdxEditor file={foundation} fonts={[]} />);
  await waitFor(() => expect(opens).toBeGreaterThan(0));
  view.rerender(<VsdxEditor file={foundation} fonts={[]} collaboration={{ clientId: 1, onReplica: (replica) => { if (replica) { replicas++; connected = true; } } }} />);
  await waitFor(() => expect(replicas).toBe(1));
  expect(connected).toBe(true);
  cleanup();
  canvasPrototype.getContext = getContext;
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
  cleanup();
  paintPageOverride = null;
  canvasPrototype.getContext = getContext;
});

