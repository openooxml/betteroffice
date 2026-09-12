import { beforeAll, expect, test } from 'bun:test';
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

const { cleanup, render, waitFor } = await import('@testing-library/react');
const originalOpenDiagram = vsdx.openDiagram;
let opens = 0;
let disposals = 0;
mock.module('@betteroffice/vsdx', () => ({
  ...vsdx,
  openDiagram: (...args: Parameters<typeof originalOpenDiagram>) => {
    opens++;
    const handle = originalOpenDiagram(...args);
    const dispose = handle.dispose.bind(handle);
    handle.dispose = () => { disposals++; dispose(); };
    return handle;
  },
}));
const { VsdxEditor } = await import('./VsdxEditor');
const { useState } = await import('react');

test('does not reopen for inline fonts and a state-setting onReady callback', async () => {
  let ready: { handle: DiagramHandle } | undefined;
  let paints = 0;
  opens = 0;
  disposals = 0;
  const getContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => { paints++; }, set: () => true }) as never;
  function Host() {
    const [, setApi] = useState<unknown>();
    return <VsdxEditor file={foundation} fonts={[]} onReady={(api) => { ready = api; setApi(api); }} />;
  }
  render(<Host />);
  await waitFor(() => expect(ready).toBeDefined());
  await waitFor(() => expect(paints).toBeGreaterThan(0));
  expect(opens).toBe(1);
  cleanup();
  await waitFor(() => expect(disposals).toBe(1));
  HTMLCanvasElement.prototype.getContext = getContext;
});

test('attaches collaboration that arrives after the file opens', async () => {
  const getContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  let connected = false;
  let replicas = 0;
  const view = render(<VsdxEditor file={foundation} fonts={[]} />);
  await waitFor(() => expect(opens).toBeGreaterThan(0));
  view.rerender(<VsdxEditor file={foundation} fonts={[]} collaboration={{ clientId: 1, onReplica: (replica) => { if (replica) { replicas++; connected = true; } } }} />);
  await waitFor(() => expect(replicas).toBe(1));
  expect(connected).toBe(true);
  cleanup();
  HTMLCanvasElement.prototype.getContext = getContext;
});

test('attaches collaboration that arrives while initialization is pending', async () => {
  const getContext = HTMLCanvasElement.prototype.getContext;
  const originalFontFace = globalThis.FontFace;
  const originalFonts = document.fonts;
  let finishFontLoad: (() => void) | undefined;
  HTMLCanvasElement.prototype.getContext = () => new Proxy({}, { get: () => () => {}, set: () => true }) as never;
  class DeferredFontFace {
    constructor(_family: string, _source: ArrayBuffer, _descriptors: FontFaceDescriptors) {}
    load() { return new Promise<FontFace>((resolve) => { finishFontLoad = () => resolve(this as unknown as FontFace); }); }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: DeferredFontFace });
  Object.defineProperty(document, 'fonts', { configurable: true, value: { add: () => {} } });
  const fonts = [{ family: 'Deferred', bytes: await readFile('C:/Windows/Fonts/arial.ttf') }];
  let replicas = 0;
  const view = render(<VsdxEditor file={foundation} fonts={fonts} />);
  await waitFor(() => expect(finishFontLoad).toBeDefined());
  view.rerender(<VsdxEditor file={foundation} fonts={fonts} collaboration={{ clientId: 2, onReplica: (replica) => { if (replica) replicas++; } }} />);
  finishFontLoad?.();
  await waitFor(() => expect(replicas).toBe(1));
  cleanup();
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: originalFontFace });
  Object.defineProperty(document, 'fonts', { configurable: true, value: originalFonts });
  HTMLCanvasElement.prototype.getContext = getContext;
});

