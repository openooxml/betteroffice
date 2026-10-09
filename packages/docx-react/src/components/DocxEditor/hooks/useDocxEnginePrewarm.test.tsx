import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { StrictMode } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as yrs from '@betteroffice/docx/yrs';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import type {
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { DocxEditor } from '../../DocxEditor';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from '../internals/engineChoice';
import { useDocxEnginePrewarm, useDocxEnginePrewarmOnBytes } from './useDocxEnginePrewarm';
import { useYrsCoreSession } from './useYrsCoreSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
const editModule = new WebAssembly.Module(
  new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
);
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>>;

beforeEach(() => {
  setMissingWorkerCapabilitiesForTests([]);
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(editModule);
});

class FakeWorker {
  static instances: FakeWorker[] = [];
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: (ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule)[] = [];
  terminated = false;
  constructor() {
    FakeWorker.instances.push(this);
  }
  postMessage(request: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule): void {
    this.posted.push(request);
  }
  requestAt(index: number): ResidentEngineWorkerRequest {
    const message = this.posted.at(index)!;
    if (!('id' in message)) throw new Error('Expected a worker request');
    return message;
  }
  reply(response: ResidentEngineWorkerResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {
    this.terminated = true;
  }
}

beforeAll(async () => {
  if (!document.fonts) {
    Object.defineProperty(document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
});

afterEach(() => {
  cleanup();
  resetEngineChoiceForTests();
  yrs.takePreloadedResidentEngineWorker()?.destroy();
  compileModule.mockRestore();
  globalThis.Worker = originalWorker;
  FakeWorker.instances = [];
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function installWorker(): void {
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
}

test('DocxEditor leaves prewarming off by default', async () => {
  installWorker();
  const editor = render(<DocxEditor showToolbar={false} />);
  await act(async () => {});
  expect(FakeWorker.instances).toHaveLength(0);
  expect(yrs.takePreloadedResidentEngineWorker()).toBeNull();
  editor.unmount();
});

test('an opted-in editor warms on mount once in StrictMode and releases an unused spare', async () => {
  installWorker();
  const editor = render(
    <StrictMode>
      <DocxEditor experimentalPrewarm showToolbar={false} />
    </StrictMode>
  );
  expect(FakeWorker.instances).toHaveLength(1);
  const worker = FakeWorker.instances[0];
  expect(worker.posted.map((request) => request.type)).toEqual(['warm']);
  await act(async () => {
    worker.reply({ id: worker.requestAt(0).id, ok: true });
  });
  expect(worker.posted).toEqual([
    { id: 1, type: 'warm', hostModule: true },
    { type: 'editModule', module: editModule },
  ]);
  expect(worker.terminated).toBe(false);
  editor.unmount();
  await waitFor(() => expect(worker.terminated).toBe(true));
  expect(yrs.takePreloadedResidentEngineWorker()).toBeNull();
});

test('byte arrival retries warm in parallel with the main-thread open', async () => {
  installWorker();
  const preparedWorkers: number[] = [];
  const prepareDocxBytes = yrs.prepareDocxBytes;
  const prepare = spyOn(yrs, 'prepareDocxBytes').mockImplementation((bytes) => {
    preparedWorkers.push(FakeWorker.instances.length);
    return prepareDocxBytes(bytes);
  });
  const errors: Error[] = [];
  const hook = renderHook(
    ({ bytes }: { bytes: Uint8Array | null }) => {
      useDocxEnginePrewarm(true);
      useDocxEnginePrewarmOnBytes(true, bytes);
      return useYrsCoreSession(true, null, null, bytes, 1, undefined, {
        onError: (error) => errors.push(error),
      });
    },
    { initialProps: { bytes: null as Uint8Array | null } }
  );
  try {
    const worker = FakeWorker.instances[0];
    await act(async () => {
      worker.reply({ id: worker.requestAt(0).id, ok: false, error: 'init failed' });
    });
    expect(worker.terminated).toBe(false);
    hook.rerender({
      bytes: new Uint8Array(
        readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx'))
      ),
    });
    expect(FakeWorker.instances).toHaveLength(1);
    expect(
      worker.posted.filter((message) => 'id' in message).map((request) => request.type)
    ).toEqual(['warm', 'warm']);
    await waitFor(() => {
      expect(errors).toEqual([]);
      expect(hook.result.current.session).not.toBeNull();
    });
    expect(preparedWorkers).toEqual([1]);
    expect(worker.posted).toEqual([
      { id: 1, type: 'warm', hostModule: true },
      { type: 'editModule', module: editModule },
      { id: 2, type: 'warm', hostModule: true },
      { type: 'editModule', module: editModule },
    ]);
    await act(async () => {
      worker.reply({ id: worker.requestAt(2).id, ok: true });
    });
    expect(worker.terminated).toBe(false);
  } finally {
    prepare.mockRestore();
    hook.unmount();
  }
});

test('disabled prewarming creates no spare when bytes arrive', async () => {
  installWorker();
  const hook = renderHook(
    ({ bytes }: { bytes: Uint8Array | null }) => {
      useDocxEnginePrewarm(false);
      useDocxEnginePrewarmOnBytes(false, bytes);
    },
    { initialProps: { bytes: null as Uint8Array | null } }
  );
  hook.rerender({ bytes: Uint8Array.of(1, 2, 3) });
  await act(async () => {});
  expect(FakeWorker.instances).toHaveLength(0);
  expect(yrs.takePreloadedResidentEngineWorker()).toBeNull();
  hook.unmount();
});
