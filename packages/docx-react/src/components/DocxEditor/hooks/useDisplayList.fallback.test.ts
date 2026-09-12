import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { ResidentEngineWorkerRequest, ResidentEngineWorkerResponse } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useRustDisplayList } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
)))));

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('starts a fresh frame and query history after a worker with a higher epoch fails', async () => {
  const native = createEditSession(9101);
  native.create_story('body', 'Fallback text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  new DataView(frame.buffer, frame.byteOffset, frame.byteLength).setBigUint64(32, 100n, true);
  let worker: FakeWorker;
  class FakeWorker {
    onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
    onerror: ((event: ErrorEvent) => void) | null = null;
    onmessageerror = null;
    constructor() {
      worker = this;
    }
    bootstrapId = 0;
    postMessage(request: ResidentEngineWorkerRequest): void {
      if (request.type === 'bootstrap') this.bootstrapId = request.id;
    }
    reply(): void {
      this.onmessage?.({ data: {
        id: this.bootstrapId, ok: true, frame: frame.slice().buffer,
        caret: { frameEpoch: 100, caretRect: null }, selection: null, layoutRevision: 1,
      } } as MessageEvent<ResidentEngineWorkerResponse>);
    }
    terminate(): void {}
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const expectedEpochs: number[] = [];
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    buildDisplayListFrame: (input: string, epoch: number) => {
      expectedEpochs.push(epoch);
      return native.build_display_list_frame(input, epoch);
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const overrides = { getInputs: () => inputs };
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await act(async () => {
      worker!.reply();
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(100);
    });
    await act(async () => {
      worker!.onerror?.({ message: 'worker crashed' } as ErrorEvent);
      rerender({ layout: { ...inputs.layout } });
    });
    await waitFor(() => expect(expectedEpochs.length).toBeGreaterThan(0));
    await waitFor(() => expect(result.current.error).toBeNull());
    expect(expectedEpochs[0]).toBe(0);
    expect(result.current.frame?.frameEpoch).toBeLessThan(100);
    expect(result.current.loading).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(
      errors.mock.calls.some(([message]) => String(message).includes('Rust display-list build failed'))
    ).toBe(false);
    const fallbackEpoch = result.current.frame!.frameEpoch;
    await act(async () => {
      rerender({ layout: { ...inputs.layout } });
    });
    await waitFor(() => expect(result.current.frame!.frameEpoch).toBeGreaterThan(fallbackEpoch));
    expect(expectedEpochs[1]).toBe(fallbackEpoch);
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});
