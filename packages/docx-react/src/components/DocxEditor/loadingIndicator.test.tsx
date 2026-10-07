import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useLayoutEffect, useRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { pagedDocx } from './__fixtures__/pagedDocx';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');
const { DocxEditor } = await import('../../index');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(import.meta.dir, 'hooks/__fixtures__/probe-linked-header.docx');
const quiet = { error: console.error, warn: console.warn };
const originalWorker = globalThis.Worker;
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>> | null = null;

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
afterEach(async () => {
  cleanup();
  takePreloadedResidentEngineWorker()?.destroy();
  await act(async () => {});
  compileModule?.mockRestore();
  compileModule = null;
  resetEngineChoiceForTests();
  globalThis.Worker = originalWorker;
});
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

interface Placement {
  phase: 'parse' | 'renderer';
  overViewport: boolean;
}

function PlacementProbe({ placements }: { placements: Placement[] }) {
  const ref = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    const layer = ref.current!.closest('[data-testid="canvas-renderer-loading"]');
    placements.push({
      phase: layer ? 'renderer' : 'parse',
      overViewport: !!layer?.parentElement?.contains(
        document.querySelector('.docx-editor__scroll-container')
      ),
    });
  }, [placements]);
  return <span ref={ref} />;
}

test('the indicator stays over the whole editor while the chrome mounts, until the first page', async () => {
  const bytes = readFileSync(FIXTURE);
  const placements: Placement[] = [];
  const view = render(
    <DocxEditor
      experimentalWorkerOpen={false}
      documentBuffer={
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
      }
      loadingIndicator={<PlacementProbe placements={placements} />}
    />
  );
  for (let attempt = 0; attempt < 200; attempt += 1) {
    if (view.container.querySelector('.canvas-page')) break;
    await act(async () => {
      await new Promise((done) => setTimeout(done, 10));
    });
  }
  expect(placements).toEqual([
    { phase: 'parse', overViewport: false },
    { phase: 'renderer', overViewport: true },
  ]);
  expect(view.container.querySelector('.canvas-page')).not.toBeNull();
  expect(view.queryByTestId('canvas-renderer-loading')).toBeNull();
}, 30_000);

test('the worker indicator stays over the whole editor while the chrome mounts, until the first page', async () => {
  const startWorker = await residentWorkerFactory();
  const workers: InProcessResidentWorker[] = [];
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  setMissingWorkerCapabilitiesForTests([]);
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  const placements: Placement[] = [];
  const view = render(
    <DocxEditor
      documentBuffer={await pagedDocx(2)}
      loadingIndicator={<PlacementProbe placements={placements} />}
    />
  );
  await waitFor(() => expect(view.container.querySelector('.canvas-page')).not.toBeNull(), {
    timeout: 20_000,
  });
  expect(workers.some((worker) => worker.requests.includes('open') && worker.sessions.length > 0)).toBe(true);
  expect(placements).toEqual([
    { phase: 'parse', overViewport: false },
    { phase: 'renderer', overViewport: true },
  ]);
  expect(view.container.querySelector('.canvas-page')).not.toBeNull();
  expect(view.queryByTestId('canvas-renderer-loading')).toBeNull();
}, 30_000);
