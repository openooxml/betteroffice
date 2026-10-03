import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import {
  DocxAsyncOnlyError,
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginContext,
  type DocxPluginSelection,
  type SelectionState,
} from '../../index';
import { pagedDocx } from './__fixtures__/pagedDocx';
import * as scrollApi from './hooks/usePagedScrollApi';
import * as viewerReads from './internals/viewerRefReads';
import { markPresented, stampWorkerFrameVersion } from './internals/layoutProvenance';
import type { YrsInputRef } from './YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
let startWorker: () => InProcessResidentWorker;
const workers: InProcessResidentWorker[] = [];
const font = readFileSync(resolve(
  import.meta.dir, '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
));

beforeAll(async () => {
  if (!document.fonts) {
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});
afterEach(() => {
  cleanup();
  mock.restore();
  globalThis.Worker = originalWorker;
  workers.length = 0;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function navigationViewer() {
  spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  const ref = createRef<DocxEditorRef>();
  let inputRef: { readonly current: YrsInputRef | null } | null = null;
  const input = { get current() { return inputRef?.current ?? null; } };
  const scrolls: Array<ReturnType<typeof mock<(position: number, forParaIdScroll?: boolean) => void>>> = [];
  const useScroll = scrollApi.usePagedScrollApi;
  spyOn(scrollApi, 'usePagedScrollApi').mockImplementation((options) => {
    const api = useScroll(options);
    inputRef = options.yrsInputRef;
    const scroll = mock((_position: number, _forParaIdScroll?: boolean) => {});
    scrolls.push(scroll);
    return { ...api, scrollToPositionImpl: scroll };
  });
  const provider = { resolve: () => () => Promise.resolve(font.buffer.slice(
    font.byteOffset, font.byteOffset + font.byteLength
  ) as ArrayBuffer) };
  const context: { current: DocxPluginContext<null> | null } = { current: null };
  const plugin = defineDocxPlugin({
    id: 'test.viewer-navigation',
    createState: () => null,
    initialize: (next) => { context.current = next; },
    onEvent: (next) => { context.current = next; },
  });
  const view = render(<DocxEditor
    ref={ref} documentBuffer={await pagedDocx(1, 2)} readOnly experimentalWorkerOpen
    previewFirstPage={false} measurementFontProvider={provider} plugins={[plugin]}
  />);
  await waitFor(() => expect(() => ref.current!.getDocument()).toThrow(DocxAsyncOnlyError), { timeout: 20_000 });
  await ref.current!.whenLayoutComplete({ timeoutMs: 20_000 });
  await waitFor(() => expect(context.current?.snapshot.layout).toBeTruthy(), { timeout: 20_000 });
  await waitFor(() => expect(input.current).not.toBeNull());
  return { ref, input, scrolls, view, worker: workers[0]! };
}

for (const gesture of ['keyboard', 'pointer']) {
  test(`pending viewer navigation yields to a newer ${gesture} selection`, async () => {
    const { ref, input, view, worker, scrolls } = await navigationViewer();
    worker.hold();
    const reads = worker.requests.filter((kind) => kind === 'documentRead').length;
    const navigation = ref.current!.scrollToParagraph('00000002');
    await waitFor(() => expect(worker.requests.filter((kind) => kind === 'documentRead').length).toBeGreaterThan(reads));
    if (gesture === 'keyboard') {
      const textarea = view.getByTestId('yrs-input');
      expect(view.container.contains(textarea)).toBe(false);
      fireEvent.keyDown(textarea, { key: 'a', ctrlKey: true });
    } else {
      act(() => {
        const next = input.current!.beginGesture!();
        input.current!.setSelectionFromDisplay(2, 5, undefined, next);
      });
    }
    const selection = input.current!.displaySelection();
    const currentGesture = input.current!.currentGesture!();
    const calls = scrolls.reduce((count, scroll) => count + scroll.mock.calls.length, 0);
    await act(async () => worker.release());
    expect(await navigation).toBe(false);
    expect(input.current!.isGestureCurrent!(currentGesture)).toBe(true);
    const after = input.current!.displaySelection()!;
    if (gesture === 'keyboard') {
      expect(Math.min(after.anchor, after.head)).toBe(0);
      expect(Math.max(after.anchor, after.head)).toBeGreaterThanOrEqual(Math.max(selection!.anchor, selection!.head));
    } else {
      expect(after).toEqual({ anchor: 2, head: 5 });
    }
    expect(scrolls.reduce((count, scroll) => count + scroll.mock.calls.length, 0)).toBe(calls);
    expect(worker.requests).not.toContain('encodeState');
  }, 40_000);
}

test('viewer navigation re-reads a newer frame and applies with the latest scroll implementation', async () => {
  const { ref, input, scrolls } = await navigationViewer();
  let finish!: () => void;
  const pending = new Promise<void>((resolve) => { finish = resolve; });
  const host = document.createElement('div');
  const frame = (version: string) => {
    const queries = { displayList: {} } as DisplayListQueries;
    stampWorkerFrameVersion(queries, version);
    markPresented(host, queries.displayList);
    return queries;
  };
  let queries = frame('v');
  const versions: string[] = [];
  const navigate = viewerReads.navigateViewer;
  spyOn(viewerReads, 'navigateViewer').mockImplementation((access, target, apply, options) =>
    navigate({
      ...access, host: () => host, queries: () => queries,
      awaitFrame: async () => { queries = frame('v2'); return queries; },
      read: (async (request: ResidentDocumentRead) => {
        if (!('expectVersion' in request)) throw new Error('Expected versioned read');
        versions.push(request.expectVersion);
        if (request.expectVersion === 'v') await pending;
        return { version: 'v2', value: { anchor: 2, head: 5 } };
      }) as ResidentEngineWorkerClient['documentRead'],
    }, target, apply, options)
  );
  const selection = spyOn(input.current!, 'setSelectionFromDisplay');
  const navigation = ref.current!.scrollToParagraph('00000002');
  const initial = scrolls.at(-1)!;
  await act(async () => ref.current!.setZoom(1.25));
  expect(scrolls.at(-1)).not.toBe(initial);
  const latest = scrolls.at(-1)!;
  await act(async () => finish());
  expect(await navigation).toBe(true);
  expect(versions).toEqual(['v', 'v2']);
  expect(initial).not.toHaveBeenCalled();
  expect(latest).toHaveBeenCalledWith(2, true);
  expect(selection).toHaveBeenCalledWith(2, 2);
}, 40_000);

test('viewer selections reach the prop, ref subscribers and plugin snapshot, including clear', async () => {
  const compile = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  try {
    const ref = createRef<DocxEditorRef>();
    const propSelections: Array<SelectionState | null> = [];
    const subscribedSelections: Array<SelectionState | null> = [];
    const pluginSelections: DocxPluginSelection[] = [];
    const context: { current: DocxPluginContext<null> | null } = { current: null };
    const plugin = defineDocxPlugin({
      id: 'test.viewer-selection',
      createState: () => null,
      initialize: (next) => { context.current = next; },
      onEvent(next, event) {
        context.current = next;
        if (event.type === 'selection-change') pluginSelections.push(event.selection);
      },
    });
    const provider = { resolve: () => () => Promise.resolve(font.buffer.slice(
      font.byteOffset, font.byteOffset + font.byteLength
    ) as ArrayBuffer) };
    const element = (buffer: ArrayBuffer) => (
      <DocxEditor
        ref={ref}
        documentBuffer={buffer}
        readOnly
        experimentalWorkerOpen
        previewFirstPage={false}
        plugins={[plugin]}
        measurementFontProvider={provider}
        onSelectionChange={(selection) => propSelections.push(selection)}
      />
    );
    const view = render(element(await pagedDocx(1, 2)));
    await waitFor(() => expect(context.current?.snapshot.layout).toBeTruthy(), { timeout: 20_000 });
    const unsubscribe = ref.current!.onSelectionChange((selection) => subscribedSelections.push(selection));
    propSelections.length = 0;
    pluginSelections.length = 0;
    fireEvent.keyDown(view.getByTestId('yrs-input'), { key: 'a', ctrlKey: true });
    await waitFor(() => expect(propSelections.at(-1)?.isMultiParagraph).toBe(true), { timeout: 10_000 });
    expect(subscribedSelections).toEqual(propSelections);
    expect(propSelections).toContainEqual({
      hasSelection: true, isMultiParagraph: false,
      textFormatting: {}, paragraphFormatting: {}, styleId: null,
      startParagraphIndex: -1, endParagraphIndex: -1,
    });
    expect(propSelections.at(-1)).toEqual({
      hasSelection: true, isMultiParagraph: true,
      textFormatting: {}, paragraphFormatting: {}, styleId: null,
      startParagraphIndex: -1, endParagraphIndex: -1,
    });
    const selected = context.current!.snapshot.selection;
    expect(selected).toEqual({ formatting: null, displayRange: {
      story: 'body', from: 0, to: expect.any(Number),
      layoutId: context.current!.snapshot.layout!.id,
    } });
    expect(selected.displayRange!.to).toBeGreaterThan(0);
    expect(pluginSelections.at(-1)).toEqual(selected);
    expect(workers[0]!.requests).not.toContain('encodeState');

    const before = propSelections.length;
    await act(async () => ref.current!.highlightRange(0, selected.displayRange!.to));
    await waitFor(() => expect(propSelections.at(-1)?.isMultiParagraph).toBe(true));
    expect(context.current!.snapshot.selection).toEqual(selected);
    expect(propSelections.slice(before)).toContainEqual({
      hasSelection: false, isMultiParagraph: false,
      textFormatting: {}, paragraphFormatting: {}, styleId: null,
      startParagraphIndex: -1, endParagraphIndex: -1,
    });
    await act(async () => ref.current!.highlightRange(selected.displayRange!.to, 0));
    expect(propSelections.at(-1)?.hasSelection).toBe(false);
    expect(context.current!.snapshot.selection).toEqual({
      formatting: null, displayRange: null,
    });
    expect(pluginSelections.at(-1)).toEqual({ formatting: null, displayRange: null });
    expect(subscribedSelections).toEqual(propSelections);
    unsubscribe();
  } finally {
    compile.mockRestore();
  }
}, 60_000);
