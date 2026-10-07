import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type { ResidentEngineWorkerRequest } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { pagedDocx } from './__fixtures__/pagedDocx';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');

// With a preview, a load's second session is its full one.
const real = await import('@betteroffice/docx/yrs');
const { createYrsSession } = real;
let created = 0;
let holdFullOpen: Promise<void> | null = null;
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: async (options: Parameters<typeof createYrsSession>[0]) => {
    created += 1;
    if (created === 2 && holdFullOpen) await holdFullOpen;
    return createYrsSession(options);
  },
}));
const canvasReplay = await import('./canvasReplay');
const { presentCanvasReplay } = canvasReplay;
let holdReplays: Promise<void> | null = null;
mock.module('./canvasReplay', () => ({
  ...canvasReplay,
  presentCanvasReplay: async (...args: Parameters<typeof presentCanvasReplay>) => {
    if (holdReplays) await holdReplays;
    return presentCanvasReplay(...args);
  },
}));
const { DocxEditor } = await import('../../index');
type Editor = import('../../index').DocxEditorRef;

const ROOT = resolve(import.meta.dir, '../../../../..');
const WASM = resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const fixture = (path: string) => {
  const bytes = readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures', path));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};
const quiet = { error: console.error, warn: console.warn };
const originalWorker = globalThis.Worker;
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>> | null = null;

async function installWorker(holdFullOpen?: Promise<void>) {
  const startWorker = await residentWorkerFactory();
  const workers: InProcessResidentWorker[] = [];
  const fullOpens: ResidentEngineWorkerRequest[] = [];
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  setMissingWorkerCapabilitiesForTests([]);
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      const post = worker.postMessage.bind(worker);
      worker.postMessage = (request, transfer) => {
        if (request.type === 'open' && request.previewBlocks === undefined) {
          fullOpens.push(request);
          if (holdFullOpen) {
            void holdFullOpen.then(() => post(request, transfer));
            return;
          }
        }
        post(request, transfer);
      };
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  return { workers, fullOpens };
}

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
beforeEach(() => {
  globalThis.Worker = undefined as unknown as typeof Worker;
});
afterEach(async () => {
  cleanup();
  takePreloadedResidentEngineWorker()?.destroy();
  await act(async () => {});
  compileModule?.mockRestore();
  compileModule = null;
  resetEngineChoiceForTests();
  globalThis.Worker = originalWorker;
  created = 0;
  holdFullOpen = null;
  holdReplays = null;
});
afterAll(async () => {
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  globalThis.Worker = originalWorker;
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a preview reports its first page painted before the full document opens', async () => {
  let releaseFullOpen = () => {};
  holdFullOpen = new Promise((done) => (releaseFullOpen = done));
  const ref = createRef<Editor>();
  let painted = 0;
  render(
    <DocxEditor
      ref={ref}
      previewFirstPage
      experimentalWorkerOpen={false}
      documentBuffer={fixture('page-fragments/pages.docx')}
      onFirstPagePainted={() => (painted += 1)}
    />
  );
  await waitFor(() => expect(created).toBe(2), { timeout: 20_000 });
  expect(painted).toBe(1);

  await act(async () => releaseFullOpen());
  await waitFor(() => expect(ref.current!.getDocument()).not.toBeNull(), {
    timeout: 20_000,
  });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(painted).toBe(1);
}, 40_000);

test('a document reports its first page painted once its canvas presents, not at layout', async () => {
  let releaseReplays = () => {};
  holdReplays = new Promise((done) => (releaseReplays = done));
  let painted = 0;
  const view = render(
    <DocxEditor
      experimentalWorkerOpen={false}
      documentBuffer={fixture('page-fragments/pages.docx')}
      onFirstPagePainted={() => (painted += 1)}
    />
  );
  await waitFor(() => expect(view.container.querySelector('.canvas-page')).not.toBeNull(), {
    timeout: 20_000,
  });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(painted).toBe(0);

  await act(async () => releaseReplays());
  await waitFor(() => expect(painted).toBe(1), { timeout: 20_000 });
}, 40_000);

test('each document reports its first page painted once, however often its pages repaint', async () => {
  const ref = createRef<Editor>();
  const painted: string[] = [];
  const editor = (name: string, buffer: ArrayBuffer) => (
    <DocxEditor ref={ref} experimentalWorkerOpen={false} documentBuffer={buffer} onFirstPagePainted={() => painted.push(name)} />
  );
  const view = render(editor('pages', fixture('page-fragments/pages.docx')));
  await waitFor(() => expect(painted).toEqual(['pages']), { timeout: 20_000 });

  await act(async () => ref.current!.setZoom(1.5));
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages']);

  view.rerender(editor('principal', fixture('structured-export/principal.docx')));
  await waitFor(() => expect(painted).toEqual(['pages', 'principal']), { timeout: 20_000 });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages', 'principal']);
}, 60_000);

test('a worker preview reports its first page painted before the full document opens', async () => {
  let releaseFullOpen = () => {};
  const gate = new Promise<void>((done) => (releaseFullOpen = done));
  const { workers, fullOpens } = await installWorker(gate);
  const ref = createRef<Editor>();
  let painted = 0;
  try {
    render(
      <DocxEditor
        ref={ref}
        previewFirstPage
        documentBuffer={await pagedDocx(2)}
        onFirstPagePainted={() => (painted += 1)}
      />
    );
    await waitFor(() => expect(fullOpens).toHaveLength(1), { timeout: 20_000 });
    await waitFor(() => expect(painted).toBe(1), { timeout: 20_000 });
    expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);

    await act(async () => releaseFullOpen());
    await waitFor(() => expect(ref.current!.getDocument()).not.toBeNull(), { timeout: 20_000 });
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(painted).toBe(1);
  } finally {
    releaseFullOpen();
  }
}, 40_000);

test('a worker document reports its first page painted once its canvas presents, not at layout', async () => {
  const { workers } = await installWorker();
  let releaseReplays = () => {};
  holdReplays = new Promise((done) => (releaseReplays = done));
  let painted = 0;
  try {
    const view = render(
      <DocxEditor
        documentBuffer={await pagedDocx(2)}
        onFirstPagePainted={() => (painted += 1)}
      />
    );
    await waitFor(() => expect(view.container.querySelector('.canvas-page')).not.toBeNull(), {
      timeout: 20_000,
    });
    expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(painted).toBe(0);

    await act(async () => releaseReplays());
    await waitFor(() => expect(painted).toBe(1), { timeout: 20_000 });
  } finally {
    releaseReplays();
  }
}, 40_000);

test('each worker document reports its first page painted once, however often its pages repaint', async () => {
  const { workers } = await installWorker();
  const ref = createRef<Editor>();
  const painted: string[] = [];
  const editor = (name: string, buffer: ArrayBuffer) => (
    <DocxEditor ref={ref} documentBuffer={buffer} onFirstPagePainted={() => painted.push(name)} />
  );
  const view = render(editor('pages', await pagedDocx(2)));
  await waitFor(() => expect(painted).toEqual(['pages']), { timeout: 20_000 });
  expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);

  await act(async () => ref.current!.setZoom(1.5));
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages']);

  view.rerender(editor('principal', await pagedDocx(1)));
  await waitFor(() => expect(painted).toEqual(['pages', 'principal']), { timeout: 20_000 });
  expect(workers.reduce((count, worker) => count + worker.requests.filter((request) => request === 'open').length, 0)).toBe(2);
  await act(async () => {
    await new Promise((done) => setTimeout(done, 300));
  });
  expect(painted).toEqual(['pages', 'principal']);
}, 60_000);
