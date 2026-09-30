import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { ResidentEngineWorkerClient, type YrsSelection, type YrsSession } from '@betteroffice/docx/yrs';
import type { ResidentEngineWorkerRequest, ResidentEngineWorkerResponse } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useCanvasRenderer, useRustDisplayList, type ResidentFrameApplyResult } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, configure, renderHook, waitFor } = await import('@testing-library/react');
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

test('continues the frame epochs after a worker with a higher epoch fails', async () => {
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
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) => {
      expectedEpochs.push(epoch);
      return native.build_display_list_frame(input, epoch);
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
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
    expect(expectedEpochs[0]).toBe(100);
    expect(result.current.frame?.frameEpoch).toBe(101);
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

test('recovers from a worker whose frame number the host engine already used', async () => {
  const native = createEditSession(9101);
  native.create_story('body', 'Fallback text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  new DataView(frame.buffer, frame.byteOffset, frame.byteLength).setBigUint64(32, 1n, true);
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
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) => {
      expectedEpochs.push(epoch);
      return native.build_display_list_frame(input, epoch);
    },
    encodeStateVector: () => new Uint8Array(),
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
      expect(result.current.frame?.frameEpoch).toBe(1);
    });
    await act(async () => {
      worker!.onerror?.({ message: 'worker crashed' } as ErrorEvent);
      rerender({ layout: { ...inputs.layout } });
    });
    await waitFor(() => expect(expectedEpochs.length).toBeGreaterThan(0));
    await waitFor(() => expect(result.current.error).toBeNull());
    expect(expectedEpochs[0]).toBe(1);
    expect(result.current.frame?.frameEpoch).toBe(2);
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
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => native.apply_input(text, epoch),
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
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

test('a layout after a worker crash the host cannot absorb as input still renders', async () => {
  const native = createEditSession(9205);
  native.create_story('body', 'Fallback text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  const header = new DataView(frame.buffer, frame.byteOffset, frame.byteLength);
  header.setBigUint64(16, 50n, true);
  header.setBigUint64(24, 50n, true);
  header.setBigUint64(32, 100n, true);
  let worker: InputFakeWorker | null = null;
  class FakeWorker extends InputFakeWorker {
    constructor() {
      super(frame);
      worker = this;
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const selection: YrsSelection = {
    anchor: { story: 'body', paraId: 'p', offset: 0 },
    head: { story: 'body', paraId: 'p', offset: 0 },
  };
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: () => {
      throw new Error('resident input state is not ready for this paragraph');
    },
    encodeStateVector: () => new Uint8Array(),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    onUpdate: () => () => {},
    selection: () => selection,
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
      worker!.replyBootstrap();
    });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(100));
    let outcome: ResidentFrameApplyResult | null | undefined;
    await act(async () => {
      const pending = result.current.applyInput('QUACK');
      await flushInputRequest(worker!);
      worker!.crash();
      outcome = await pending;
    });
    expect(outcome).toBeNull();
    await act(async () => {
      rerender({ layout: { ...inputs.layout } });
    });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBeGreaterThan(100));
    expect(result.current.error).toBeNull();
    expect(
      errors.mock.calls.some(([message]) => String(message).includes('Rust display-list build failed'))
    ).toBe(false);
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
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => native.apply_input(text, epoch),
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
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
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => native.apply_input(text, epoch),
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
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

test('a StrictMode remount keeps the worker its second mount started', async () => {
  const native = createEditSession(9102);
  native.create_story('body', 'Remounted text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  const frame = native.build_display_list_frame(JSON.stringify(inputs), 0);
  const workers: FakeWorker[] = [];
  class FakeWorker {
    onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
    onerror: ((event: ErrorEvent) => void) | null = null;
    onmessageerror = null;
    bootstrapId = 0;
    terminated = false;
    constructor() {
      workers.push(this);
    }
    postMessage(request: ResidentEngineWorkerRequest): void {
      if (request.type === 'bootstrap') this.bootstrapId = request.id;
    }
    reply(): void {
      this.onmessage?.({ data: {
        id: this.bootstrapId, ok: true, frame: frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null }, selection: null, layoutRevision: 1,
      } } as MessageEvent<ResidentEngineWorkerResponse>);
    }
    terminate(): void {
      this.terminated = true;
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const overrides = { getInputs: () => inputs };
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  configure({ reactStrictMode: true });
  try {
    const { result, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    configure({ reactStrictMode: false });
    expect(workers).toHaveLength(2);
    expect(workers[0].terminated).toBe(true);
    await act(async () => {
      workers[1].reply();
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
    });
    expect(workers[1].terminated).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(true);
    unmount();
    expect(workers[1].terminated).toBe(true);
  } finally {
    configure({ reactStrictMode: false });
    errors.mockRestore();
    native.free();
  }
});

test('unmounting while the worker builds never builds on the host engine', async () => {
  const native = createEditSession(9103);
  native.create_story('body', 'Unmounted text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  let posted = 0;
  class FakeWorker {
    onmessage = null;
    onerror = null;
    onmessageerror = null;
    postMessage(): void {
      posted += 1;
    }
    terminate(): void {}
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const builds: string[] = [];
  const engine = {
    buildDisplayListJson: (input: string) => {
      builds.push('json');
      return native.build_display_list_json(input);
    },
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) => {
      builds.push('frame');
      return native.build_display_list_frame(input, epoch);
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, { getInputs: () => inputs }, undefined, undefined, engine)
    );
    expect(posted).toBe(1);
    unmount();
    await act(async () => {
      await new Promise((done) => setTimeout(done, 0));
    });
    expect(builds).toEqual([]);
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('a load that fails while the worker builds drops the worker, never the released engine', async () => {
  const native = createEditSession(9104);
  native.create_story('body', 'Released text', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  })));
  let worker: InputFakeWorker | null = null;
  class FakeWorker extends InputFakeWorker {
    constructor() {
      super(new Uint8Array());
      worker = this;
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const hostCalls: string[] = [];
  const engine = {
    buildDisplayListJson: (input: string) => {
      hostCalls.push('json');
      return native.build_display_list_json(input);
    },
    resetFrameBase: () => {
      hostCalls.push('resetFrameBase');
      return native.reset_frame_base();
    },
    buildDisplayListFrame: (input: string, epoch: number) => {
      hostCalls.push('frame');
      return native.build_display_list_frame(input, epoch);
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { rerender } = renderHook(
      ({ layout, session }) =>
        useRustDisplayList(layout, { getInputs: () => inputs }, undefined, undefined, session),
      {
        initialProps: {
          layout: inputs.layout as Layout | null,
          session: engine as YrsSession | null,
        },
      }
    );
    expect(worker!.posted.map((request) => request.type)).toEqual(['bootstrap']);
    rerender({ layout: null, session: null });
    expect(worker!.terminated).toBe(true);
    await act(async () => {
      worker!.crash();
      await new Promise((done) => setTimeout(done, 0));
    });
    expect(hostCalls).toEqual([]);
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('a replaced worker preserves the current document when its delayed provisional layout fails', async () => {
  const request = JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  });
  const documents = ['Document A', 'Document B'].map((text, index) => {
    const native = createEditSession(9302 + index);
    native.create_story('body', text, 'Normal', 'left');
    const [paragraph] = JSON.parse(native.paragraphs('body')) as Array<{ paraId: string }>;
    native.set_selection('body', paragraph!.paraId, 0, paragraph!.paraId, 0);
    const selection = JSON.parse(native.selection()) as YrsSelection;
    const layoutJson = native.layout_document_with_regions_retained_json(request);
    const frame = native.build_display_list_frame(JSON.stringify({}), 0);
    const frameEpoch = 100 + index;
    new DataView(frame.buffer, frame.byteOffset, frame.byteLength)
      .setBigUint64(32, BigInt(frameEpoch), true);
    const caret = { ...JSON.parse(native.resident_caret_snapshot_json()), frameEpoch };
    const adopted: string[] = [];
    const resets: string[] = [];
    const engine = {
      adoptResidentWorkerLayout: (input: string) => {
        adopted.push(input);
        return adopted.length;
      },
      residentLayoutInWorker: () => true,
      resetFrameBase: () => {
        resets.push('resetFrameBase');
        native.reset_frame_base();
      },
      residentWorkerProbe: () => ({ layoutRevision: adopted.length }),
      residentWorkerSnapshot: () => ({
        state: new Uint8Array(), fonts: [], fontsRevision: 0, layoutRevision: adopted.length,
      }),
      encodeStateVector: () => new Uint8Array(),
      onUpdate: () => () => {},
      selection: () => selection,
      applyUpdate: () => null,
    } as unknown as YrsSession;
    return { native, layoutJson, frame, frameEpoch, caret, selection, engine, adopted, resets };
  });
  const documentA = documents[0]!;
  const documentB = documents[1]!;
  const workers: FakeWorker[] = [];
  class FakeWorker extends InputFakeWorker {
    constructor() {
      super(new Uint8Array());
      workers.push(this);
    }
    reply(response: ResidentEngineWorkerResponse): void {
      this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, undefined, undefined, undefined, source),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const pendingA = result.current.layoutInWorker(documentA.engine, request);
    expect(pendingA).not.toBeNull();
    const workerA = workers[0]!;
    expect(workerA.posted[0]).toMatchObject({ type: 'bootstrap', provisionalPages: 3 });
    workerA.reply({
      id: workerA.posted[0]!.id,
      ok: true,
      frame: documentA.frame.slice().buffer,
      caret: documentA.caret,
      selection: documentA.selection,
      layoutRevision: 1,
      layoutJson: documentA.layoutJson,
      layoutProvisional: true,
    });
    const provisional = await pendingA!;
    expect(provisional?.complete).toBeDefined();
    let completed = false;
    const completeA = provisional!.complete!.then((outcome) => {
      completed = true;
      return outcome;
    });

    const pendingB = result.current.layoutInWorker(documentB.engine, request);
    expect(pendingB).not.toBeNull();
    expect(workers).toHaveLength(2);
    expect(workerA.terminated).toBe(true);
    expect(workerA.posted.map((message) => message.type)).toEqual(['bootstrap', 'destroy']);
    const workerB = workers[1]!;
    workerB.reply({
      id: workerB.posted[0]!.id,
      ok: true,
      frame: documentB.frame.slice().buffer,
      caret: documentB.caret,
      selection: documentB.selection,
      layoutRevision: 1,
      layoutJson: documentB.layoutJson,
    });
    const computationB = await pendingB!;
    await act(async () => {
      rerender({ layout: computationB!.layout, source: documentB.engine });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(documentB.frameEpoch);
    });
    act(() => result.current.setWorkerPresentationActive(true));
    const { frame, queries, caret, displayList } = result.current;
    expect(queries).not.toBeNull();
    expect(caret).not.toBeNull();
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(result.current.workerPresentationActive).toBe(true);
    expect(completed).toBe(false);

    await act(async () => {
      expect(await completeA).toBeNull();
    });

    expect(completed).toBe(true);
    expect(result.current.frame).toBe(frame);
    expect(result.current.frame?.frameEpoch).toBe(documentB.frameEpoch);
    expect(result.current.queries).toBe(queries);
    expect(result.current.caret).toBe(caret);
    expect(result.current.displayList).toBe(displayList);
    expect(result.current.error).toBeNull();
    expect(result.current.loading).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(result.current.workerPresentationActive).toBe(true);
    expect(workerA.posted.map((message) => message.type)).toEqual(['bootstrap', 'destroy']);
    expect(workerB.posted.map((message) => message.type)).toEqual(['bootstrap']);
    expect(workerB.terminated).toBe(false);
    expect(documentA.adopted).toEqual([request]);
    expect(documentB.adopted).toEqual([request]);
    expect(documentA.resets).toEqual([]);
    expect(documentB.resets).toEqual([]);
    expect(
      errors.mock.calls.some(([message]) => String(message).includes('Resident engine worker unavailable'))
    ).toBe(false);
    unmount();
  } finally {
    errors.mockRestore();
    for (const { native } of documents) native.free();
  }
});

test('a failed successor worker constructor ignores the released engine when provisional completion rejects', async () => {
  const request = JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
    renderEnv: {},
  });
  const documents = ['Document A', 'Document B'].map((text, index) => {
    const native = createEditSession(9304 + index);
    native.create_story('body', text, 'Normal', 'left');
    const inputs = JSON.parse(native.layout_document_with_regions_json(request));
    const layoutJson = native.layout_document_with_regions_retained_json(request);
    const frame = native.build_display_list_frame(JSON.stringify({}), 0);
    const frameEpoch = 100 + index;
    new DataView(frame.buffer, frame.byteOffset, frame.byteLength)
      .setBigUint64(32, BigInt(frameEpoch), true);
    const caret = { ...JSON.parse(native.resident_caret_snapshot_json()), frameEpoch };
    const adopted: string[] = [];
    let released = false;
    const afterRelease: string[] = [];
    const host = {
      adoptResidentWorkerLayout: (input: string) => {
        adopted.push(input);
        return adopted.length;
      },
      residentLayoutInWorker: () => adopted.length > 0,
      buildDisplayListJson: (input: string) => native.build_display_list_json(input),
      buildDisplayListFrame: (input: string, epoch: number) =>
        native.build_display_list_frame(input, epoch),
      resetFrameBase: () => native.reset_frame_base(),
      residentWorkerProbe: () => adopted.length > 0 ? { layoutRevision: adopted.length } : null,
      residentWorkerSnapshot: () => ({
        state: new Uint8Array(), fonts: [], fontsRevision: 0, layoutRevision: adopted.length,
      }),
      encodeStateVector: () => new Uint8Array(),
      onUpdate: () => () => {},
      selection: () => null,
      applyUpdate: () => null,
      displayHitTestRegionsJson: (pageIndex: number, x: number, y: number) =>
        native.display_hit_test_regions_json(pageIndex, x, y),
      displayVerticalMoveJson: (position: number, direction: 'up' | 'down', goalX: number) =>
        native.display_vertical_move_json(position, direction, goalX),
      displayRangeRectsJson: (from: number, to: number) => native.display_range_rects_json(from, to),
      displayRangeRectsRegionJson: (...args: Parameters<YrsSession['displayRangeRectsRegionJson']>) =>
        native.display_range_rects_region_json(...args),
    } as unknown as YrsSession;
    const engine = new Proxy(host, {
      get(target, key, receiver) {
        const value = Reflect.get(target, key, receiver);
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => {
          if (released) {
            afterRelease.push(String(key));
            throw new Error(`Released engine method: ${String(key)}`);
          }
          return value.apply(target, args);
        };
      },
    });
    return {
      native, inputs, layoutJson, frame, frameEpoch, caret, engine, adopted, afterRelease,
      release: () => { released = true; },
    };
  });
  const documentA = documents[0]!;
  const documentB = documents[1]!;
  const constructorFailure = new Error('successor worker construction failed');
  let constructionAttempts = 0;
  const workers: FakeWorker[] = [];
  class FakeWorker extends InputFakeWorker {
    constructor() {
      constructionAttempts += 1;
      if (constructionAttempts === 2) throw constructorFailure;
      super(new Uint8Array());
      workers.push(this);
    }
    reply(response: ResidentEngineWorkerResponse): void {
      this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
    }
  }
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  let failCompletion!: () => void;
  const blocked = new Promise<void>((resolve) => {
    failCompletion = resolve;
  });
  const completionFailure = new Error('delayed provisional completion failed');
  const completeLayout = spyOn(ResidentEngineWorkerClient.prototype, 'completeLayout')
    .mockImplementation(async () => {
      await blocked;
      throw completionFailure;
    });
  const overrides = { getInputs: () => documentB.inputs };
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const pendingA = result.current.layoutInWorker(documentA.engine, request);
    expect(pendingA).not.toBeNull();
    const workerA = workers[0]!;
    workerA.reply({
      id: workerA.posted[0]!.id,
      ok: true,
      frame: documentA.frame.slice().buffer,
      caret: documentA.caret,
      selection: null,
      layoutRevision: 1,
      layoutJson: documentA.layoutJson,
      layoutProvisional: true,
    });
    const provisional = await pendingA!;
    expect(provisional?.complete).toBeDefined();
    const completeA = provisional!.complete!.catch((cause) => cause);
    await waitFor(() => expect(completeLayout).toHaveBeenCalledTimes(1));

    expect(() => result.current.layoutInWorker(documentB.engine, request)).toThrow(constructorFailure);
    expect(constructionAttempts).toBe(2);
    expect(workers).toHaveLength(1);
    expect(workerA.terminated).toBe(true);
    expect(workerA.posted.map((message) => message.type)).toEqual(['bootstrap', 'destroy']);
    documentA.release();
    documentB.engine.resetFrameBase();
    await act(async () => {
      rerender({ layout: documentB.inputs.layout as Layout, source: documentB.engine });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame).not.toBeNull();
      expect(result.current.loading).toBe(false);
    });
    const { frame, queries, caret, displayList } = result.current;
    expect(queries).not.toBeNull();
    expect(queries!.isReady()).toBe(true);
    const rects = queries!.rangeRects(1, 2);
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(await result.current.resolveQueries()).toEqual({ queries: queries!, frameEpoch: frame!.frameEpoch });

    let outcome: unknown;
    await act(async () => {
      failCompletion();
      outcome = await completeA;
    });

    expect(documentA.afterRelease).toEqual([]);
    expect(outcome).toBeNull();
    expect(result.current.frame).toBe(frame);
    expect(result.current.queries).toBe(queries);
    expect(result.current.caret).toBe(caret);
    expect(result.current.displayList).toBe(displayList);
    expect(queries!.rangeRects(1, 2)).toEqual(rects);
    expect(await result.current.resolveQueries()).toEqual({ queries: queries!, frameEpoch: frame!.frameEpoch });
    expect(result.current.error).toBeNull();
    expect(result.current.loading).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(result.current.workerPresentationActive).toBe(false);
    expect(constructionAttempts).toBe(2);
    expect(documentA.adopted).toEqual([request]);
    expect(documentB.adopted).toEqual([]);
    expect(documentB.afterRelease).toEqual([]);
    expect(errors).not.toHaveBeenCalled();
    unmount();
  } finally {
    completeLayout.mockRestore();
    errors.mockRestore();
    for (const { native } of documents) native.free();
  }
});

test('unmounting while the worker applies input never replays it on the host engine', async () => {
  const native = createEditSession(9204);
  native.create_story('body', 'Unmounted input', 'Normal', 'left');
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
  const hostCalls: string[] = [];
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    resetFrameBase: () => {
      hostCalls.push('resetFrameBase');
      return native.reset_frame_base();
    },
    buildDisplayListFrame: (input: string, epoch: number) =>
      native.build_display_list_frame(input, epoch),
    applyInput: (text: string, epoch: number) => {
      hostCalls.push('applyInput');
      return native.apply_input(text, epoch);
    },
    residentCaretSnapshot: () => JSON.parse(native.resident_caret_snapshot_json()),
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => JSON.parse(native.selection()) as YrsSelection,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, { getInputs: () => inputs }, undefined, undefined, engine)
    );
    await act(async () => {
      worker!.replyBootstrap();
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(100);
    });
    let pending!: Promise<ResidentFrameApplyResult | null>;
    await act(async () => {
      pending = result.current.applyInput('QUACK');
      await flushInputRequest(worker!);
    });
    unmount();
    expect(worker!.terminated).toBe(true);
    const outcome = await pending;
    expect(outcome?.frameEpoch).toBeNull();
    expect(hostCalls).toEqual([]);
    expect(native.paragraphs('body')).not.toContain('QUACK');
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('each session decodes its images into a cache of its own', () => {
  const { result } = renderHook(() => useCanvasRenderer());
  const first = { name: 'first' } as unknown as YrsSession;
  act(() => result.current.onLayoutComputed(null, first));
  const firstImages = result.current.resolveImage;
  act(() => result.current.onLayoutComputed(null, first));
  expect(result.current.resolveImage).toBe(firstImages);
  act(() => result.current.onLayoutComputed(null, { name: 'next' } as unknown as YrsSession));
  expect(result.current.resolveImage).not.toBe(firstImages);
});
