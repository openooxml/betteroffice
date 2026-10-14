import { afterEach, beforeEach, spyOn } from 'bun:test';
import { takePreloadedResidentEngineWorker } from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';

export function setupWorkerEngine(configureWorker?: (worker: InProcessResidentWorker) => void) {
  const workers: InProcessResidentWorker[] = [];
  let originalWorker: typeof Worker;
  let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>> | null = null;

  beforeEach(async () => {
    originalWorker = globalThis.Worker;
    const startWorker = await residentWorkerFactory();
    compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
      new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
    ));
    globalThis.Worker = class {
      constructor() {
        const worker = startWorker();
        configureWorker?.(worker);
        workers.push(worker);
        return worker;
      }
    } as unknown as typeof Worker;
  });

  afterEach(async () => {
    const { act, cleanup } = await import('@testing-library/react');
    try {
      cleanup();
      takePreloadedResidentEngineWorker()?.destroy();
      await act(async () => {});
    } finally {
      for (const worker of workers.splice(0)) worker.terminate();
      compileModule?.mockRestore();
      compileModule = null;
      globalThis.Worker = originalWorker;
    }
  });

  return workers;
}
