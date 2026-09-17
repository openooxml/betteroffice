import { afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as vsdx from '@betteroffice/vsdx';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const root = resolve(import.meta.dir, '../../..');
let stencil: Uint8Array;

beforeAll(async () => {
  const [wasm, fixture] = await Promise.all([
    readFile(resolve(import.meta.dir, '../../vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/document-stencil.vsdx')),
  ]);
  await vsdx.initWasm(wasm);
  stencil = fixture;
});

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const originalOpenDiagram = vsdx.openDiagram;
let masterQueries = 0;
mock.module('@betteroffice/vsdx', () => ({
  ...vsdx,
  openDiagram: (...args: Parameters<typeof originalOpenDiagram>) => {
    const handle = originalOpenDiagram(...args);
    const masters = handle.masters.bind(handle);
    handle.masters = () => { masterQueries += 1; return masters(); };
    return handle;
  },
}));
const { VsdxEditor } = await import('./VsdxEditor');

afterEach(() => cleanup());

test('master previews wait for the Document Stencil instead of blocking the open', async () => {
  masterQueries = 0;
  const canvasPrototype = Object.getPrototypeOf(document.createElement('canvas')) as HTMLCanvasElement;
  const getContext = canvasPrototype.getContext;
  canvasPrototype.getContext = () => new Proxy({}, { get: (_target, key) => key === 'measureText' ? () => ({ width: 0 }) : () => {}, set: () => true }) as never;
  try {
    const view = render(<VsdxEditor file={stencil} fonts={[]} />);
    const rail = (label: string) => view.container.querySelector(`button[aria-label="${label}"]`) as HTMLButtonElement;
    await waitFor(() => expect(rail('Document Stencil')).not.toBeNull());
    expect(masterQueries).toBe(0);
    await act(async () => { fireEvent.click(rail('Document Stencil')); });
    expect(masterQueries).toBe(1);
    const tile = view.container.querySelector('button[aria-label="Stencil-Rect"]') as HTMLButtonElement;
    expect(tile).not.toBeNull();
    expect(tile.querySelector('path')?.getAttribute('d')).not.toBe('');
  } finally {
    canvasPrototype.getContext = getContext;
  }
});
