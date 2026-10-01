import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { StrictMode } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as yrs from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { DocxEditor } from '../../DocxEditor';
import { useDocxEnginePrewarm, useDocxEnginePrewarmOnBytes } from './useDocxEnginePrewarm';
import { useYrsCoreSession } from './useYrsCoreSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;

class FakeWorker {
  static instances: FakeWorker[] = [];
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  terminated = false;
  constructor() {
    FakeWorker.instances.push(this);
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
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
  yrs.takePreloadedResidentEngineWorker()?.destroy();
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
    worker.reply({ id: worker.posted[0].id, ok: true });
  });
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
      worker.reply({ id: worker.posted[0].id, ok: false, error: 'init failed' });
    });
    expect(worker.terminated).toBe(false);
    hook.rerender({
      bytes: new Uint8Array(
        readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx'))
      ),
    });
    expect(FakeWorker.instances).toHaveLength(1);
    expect(worker.posted.map((request) => request.type)).toEqual(['warm', 'warm']);
    await waitFor(() => {
      expect(errors).toEqual([]);
      expect(hook.result.current.session).not.toBeNull();
    });
    expect(preparedWorkers).toEqual([1]);
    expect(worker.posted.map((request) => request.type)).toEqual(['warm', 'warm']);
    await act(async () => {
      worker.reply({ id: worker.posted[1].id, ok: true });
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
