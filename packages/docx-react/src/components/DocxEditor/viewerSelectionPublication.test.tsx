import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginContext,
  type DocxPluginSelection,
  type SelectionState,
} from '../../index';
import { pagedDocx } from './__fixtures__/pagedDocx';

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
  globalThis.Worker = originalWorker;
  workers.length = 0;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

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
