import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsSelection, YrsSession } from '@betteroffice/docx/yrs';
import type { ResidentEngineWorkerRequest, ResidentEngineWorkerResponse } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useRustDisplayList, type ResidentFrameApplyResult } from './useDisplayList';

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

class InputFakeWorker {
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  terminated = false;
  constructor(
    private readonly bootstrapFrame: Uint8Array,
    public onPost?: (request: ResidentEngineWorkerRequest) => void
  ) {}
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
    this.onPost?.(request);
  }
  replyBootstrap(): void {
    const bootstrap = this.posted.find((request) => request.type === 'bootstrap');
    if (!bootstrap) throw new Error('worker never received a bootstrap request');
    this.onmessage?.({ data: {
      id: bootstrap.id, ok: true, frame: this.bootstrapFrame.slice().buffer,
      caret: { frameEpoch: 100, caretRect: null }, selection: null, layoutRevision: 1,
    } } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  replyInputError(message: string): void {
    const input = [...this.posted].reverse().find((request) => request.type === 'applyInput');
    if (!input) throw new Error('worker never received an applyInput request');
    this.onmessage?.({ data: {
      id: input.id, ok: false, error: message,
    } } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  replyInputCorruptFrame(payload: Uint8Array): void {
    const input = [...this.posted]
      .reverse()
      .find((request) => request.type === 'applyInput' || request.type === 'applyDelete');
    if (!input) throw new Error('worker never received an input request');
    this.onmessage?.({ data: {
      id: input.id, ok: true, frame: payload.slice().buffer,
      caret: { frameEpoch: 100, caretRect: null }, selection: null, layoutRevision: 1,
    } } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  crash(): void {
    this.onerror?.({ message: 'worker crashed' } as ErrorEvent);
  }
  terminate(): void {
    this.terminated = true;
  }
}

async function flushInputRequest(worker: InputFakeWorker): Promise<void> {
  for (let i = 0; i < 25 && !worker.posted.some((request) => request.type === 'applyInput'); i += 1) {
    await Promise.resolve();
  }
  expect(worker.posted.some((request) => request.type === 'applyInput')).toBe(true);
}

test('falls back to the main thread and keeps the keystroke when the worker crashes mid-input', async () => {
  const native = createEditSession(9202);
  native.create_story('body', 'Fallback text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  new DataView(frame.buffer, frame.byteOffset, frame.byteLength).setBigUint64(32, 100n, true);
  const paragraphs = JSON.parse(native.paragraphs('body')) as Array<{ paraId: string; text: string }>;
  const para = paragraphs[0]!;
  native.set_selection('body', para.paraId, para.text.length, para.paraId, para.text.length);
  let worker: InputFakeWorker | null = null;
  class FakeWorker extends InputFakeWorker {
    constructor() {
      super(frame);
      worker = this;
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => native.apply_input(text, epoch),
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    onUpdate: () => () => {},
    selection: () => JSON.parse(native.selection()) as YrsSelection,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const overrides = { getInputs: () => inputs };
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await act(async () => {
      worker!.replyBootstrap();
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(100);
    });
    let outcome: ResidentFrameApplyResult | null | undefined;
    await act(async () => {
      const pending = result.current.applyInput('QUACK');
      await flushInputRequest(worker!);
      worker!.crash();
      outcome = await pending;
    });
    await waitFor(() => expect(result.current.error).toBeNull());
    expect(outcome?.frameEpoch).not.toBeNull();
    expect(result.current.displayList).not.toBeNull();
    expect(JSON.stringify(result.current.displayList)).toContain('QUACK');
    expect(result.current.loading).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(
      errors.mock.calls.some(([message]) => String(message).includes('falling back to the main-thread engine'))
    ).toBe(true);
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('surfaces an engine-level input rejection instead of falling back', async () => {
  const native = createEditSession(9203);
  native.create_story('body', 'Fallback text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  new DataView(frame.buffer, frame.byteOffset, frame.byteLength).setBigUint64(32, 100n, true);
  const paragraphs = JSON.parse(native.paragraphs('body')) as Array<{ paraId: string; text: string }>;
  const para = paragraphs[0]!;
  native.set_selection('body', para.paraId, para.text.length, para.paraId, para.text.length);
  let worker: InputFakeWorker | null = null;
  class FakeWorker extends InputFakeWorker {
    constructor() {
      super(frame);
      worker = this;
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => native.apply_input(text, epoch),
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    onUpdate: () => () => {},
    selection: () => JSON.parse(native.selection()) as YrsSelection,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const overrides = { getInputs: () => inputs };
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await act(async () => {
      worker!.replyBootstrap();
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(100);
    });
    let outcome: ResidentFrameApplyResult | null | undefined;
    await act(async () => {
      const pending = result.current.applyInput('QUACK');
      await flushInputRequest(worker!);
      worker!.replyInputError('apply_input requires a collapsed selection');
      outcome = await pending;
    });
    expect(outcome).toEqual({ frameEpoch: null, caretSynchronized: false });
    expect(result.current.error?.message).toContain('apply_input requires a collapsed selection');
    expect(JSON.stringify(result.current.displayList ?? '')).not.toContain('QUACK');
    expect(worker!.terminated).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(true);
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('falls back to the main thread and keeps the keystroke when the worker returns a corrupt frame', async () => {
  const native = createEditSession(9204);
  native.create_story('body', 'Fallback text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  new DataView(frame.buffer, frame.byteOffset, frame.byteLength).setBigUint64(32, 100n, true);
  const paragraphs = JSON.parse(native.paragraphs('body')) as Array<{ paraId: string; text: string }>;
  const para = paragraphs[0]!;
  native.set_selection('body', para.paraId, para.text.length, para.paraId, para.text.length);
  let worker: InputFakeWorker | null = null;
  class FakeWorker extends InputFakeWorker {
    constructor() {
      super(frame);
      worker = this;
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => native.apply_input(text, epoch),
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    onUpdate: () => () => {},
    selection: () => JSON.parse(native.selection()) as YrsSelection,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const overrides = { getInputs: () => inputs };
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await act(async () => {
      worker!.replyBootstrap();
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(100);
    });
    let outcome: ResidentFrameApplyResult | null | undefined;
    await act(async () => {
      const pending = result.current.applyInput('QUACK');
      await flushInputRequest(worker!);
      worker!.replyInputCorruptFrame(new Uint8Array([1, 2, 3, 4]));
      outcome = await pending;
    });
    await waitFor(() => expect(result.current.error).toBeNull());
    expect(outcome?.frameEpoch).not.toBeNull();
    expect(result.current.displayList).not.toBeNull();
    expect(JSON.stringify(result.current.displayList)).toContain('QUACK');
    expect(result.current.loading).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(worker!.terminated).toBe(true);
    expect(
      errors.mock.calls.some(([message]) => String(message).includes('falling back to the main-thread engine'))
    ).toBe(true);
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});
