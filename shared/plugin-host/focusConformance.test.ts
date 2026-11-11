import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import { act, createElement, type ReactNode } from 'react';
import type { PresentationHandle } from '@betteroffice/pptx';
import type { WorkbookHandle } from '@betteroffice/xlsx';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
const { createRoot } = await import('react-dom/client');
const { trackFocus } = await import('./focusTracker');
const { createDocxPluginHost } = await import('../../packages/docx-react/src/plugins/createDocxPluginHost');
const { defineDocxPlugin } = await import('../../packages/docx-react/src/plugins/defineDocxPlugin');
const { PluginOverlays: DocxPluginOverlays } = await import('../../packages/docx-react/src/plugins/PluginOverlays');
const { createXlsxPluginHost } = await import('../../packages/xlsx-react/src/plugins/createXlsxPluginHost');
const { defineXlsxPlugin } = await import('../../packages/xlsx-react/src/plugins/defineXlsxPlugin');
const { PluginOverlays: XlsxPluginOverlays } = await import('../../packages/xlsx-react/src/plugins/PluginOverlays');
const { createPptxPluginHost } = await import('../../packages/pptx-react/src/plugins/createPptxPluginHost');
const { definePptxPlugin } = await import('../../packages/pptx-react/src/plugins/definePptxPlugin');
const { PluginOverlays: PptxPluginOverlays } = await import('../../packages/pptx-react/src/plugins/PluginOverlays');

const settle = () => new Promise((done) => setTimeout(done, 0));
const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
  document.body.replaceChildren();
});
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

type PluginEvent = { type: string; focused?: boolean };

interface EngineHost {
  open(): void;
  close(): void;
  focusChanged(focused: boolean): void;
  setFocusCheck(check: (() => void) | null): void;
  focused(): boolean | undefined;
  focusEvents: boolean[];
  overlayMount: 'root' | 'editor';
  overlays(editor: HTMLElement): ReactNode;
}

const record = (events: boolean[]) => (_context: unknown, event: PluginEvent) => {
  if (event.type === 'focus-change') events.push(event.focused!);
};

function docxHost(): EngineHost {
  const focusEvents: boolean[] = [];
  const host = createDocxPluginHost({
    pagedEditorRef: { current: null },
    writeMode: () => 'viewing',
    commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }),
    subscribeLayout: () => () => {},
    geometry: () => null,
    translate: (key) => key,
  });
  host.setPlugins([defineDocxPlugin({ id: 'focus', createState: () => null, onEvent: record(focusEvents) })]);
  return {
    open: () => host.openWorker('v1'),
    close: () => host.close('unmounted'),
    focusChanged: (focused) => host.focusChanged(focused),
    setFocusCheck: (check) => host.setFocusCheck(check),
    focused: () => host.activations()[0]?.context.snapshot.focused,
    focusEvents,
    overlayMount: 'editor',
    overlays: (editor) => createElement(DocxPluginOverlays, { host, activations: [], target: editor, layerRef: () => {} }),
  };
}

function xlsxHost(): EngineHost {
  const focusEvents: boolean[] = [];
  const workbook = { version: () => 'v1', onUpdate: () => () => {} } as unknown as WorkbookHandle;
  const host = createXlsxPluginHost({
    handle: () => workbook,
    admit: async <T>(_handle: WorkbookHandle, operation: () => T) => ({ ok: true as const, value: operation() }),
    readOnlyRefusal: () => null,
    applyEdits: () => { throw new Error('Unavailable'); },
    commands: () => null,
    navigator: { selectCells: () => null, scrollToCell: () => null },
    geometry: () => null,
    translate: (key) => key,
  });
  host.setPlugins([defineXlsxPlugin({ id: 'focus', createState: () => null, onEvent: record(focusEvents) })]);
  return {
    open: () => host.open(workbook),
    close: () => host.close('unmounted'),
    focusChanged: (focused) => host.focusChanged(focused),
    setFocusCheck: (check) => host.setFocusCheck(check),
    focused: () => host.activations()[0]?.context.snapshot.focused,
    focusEvents,
    overlayMount: 'root',
    overlays: () => createElement(XlsxPluginOverlays, { host, activations: [], layerRef: () => {}, width: 300, height: 200 }),
  };
}

function pptxHost(): EngineHost {
  const focusEvents: boolean[] = [];
  const presentation = { version: () => 'v1', onUpdate: () => () => {} } as unknown as PresentationHandle;
  const host = createPptxPluginHost({
    handle: () => presentation,
    admit: async (_handle, operation) => ({ ok: true, value: operation() }),
    readOnlyRefusal: () => null,
    applyEdits: () => { throw new Error('Unavailable'); },
    commands: () => null,
    navigator: { goToSlide: () => null, selectShape: () => null, selectText: () => null },
    geometry: () => null,
    translate: (key) => key,
  });
  host.setPlugins([definePptxPlugin({ id: 'focus', createState: () => null, onEvent: record(focusEvents) })]);
  return {
    open: () => host.open(presentation),
    close: () => host.close('unmounted'),
    focusChanged: (focused) => host.focusChanged(focused),
    setFocusCheck: (check) => host.setFocusCheck(check),
    focused: () => host.activations()[0]?.context.snapshot.focused,
    focusEvents,
    overlayMount: 'root',
    overlays: () => createElement(PptxPluginOverlays, { host, activations: [], layerRef: () => {} }),
  };
}

function mountSurface(engine: EngineHost) {
  const editor = document.createElement('div');
  const root = document.createElement('div');
  root.className = 'paged-editor';
  const input = document.createElement('input');
  root.append(input);
  editor.append(root);
  document.body.append(editor);
  const surface = { editor, root, input };
  const track = () => {
    const tracker = trackFocus(() => [surface.root], (focused) => engine.focusChanged(focused));
    engine.setFocusCheck(tracker.recheck);
    cleanups.push(() => {
      engine.setFocusCheck(null);
      tracker();
    });
  };
  const renderOverlays = () => {
    const mount = document.createElement('div');
    (engine.overlayMount === 'root' ? root : document.body).append(mount);
    const react = createRoot(mount);
    act(() => react.render(engine.overlays(editor)));
    cleanups.push(() => act(() => react.unmount()));
    const layer = editor.querySelector('[data-testid="plugin-overlays"]')!;
    const field = document.createElement('input');
    layer.append(field);
    return field;
  };
  const swap = () => {
    const next = document.createElement('div');
    next.className = 'paged-editor';
    surface.root.replaceWith(next);
    surface.root = next;
  };
  return { surface, track, renderOverlays, swap };
}

const engines: [string, () => EngineHost][] = [['docx', docxHost], ['xlsx', xlsxHost], ['pptx', pptxHost]];

const cases: [string, (engine: EngineHost) => Promise<void>][] = [
  ['focus in a plugin overlay input is outside the document surface', async (engine) => {
    const { surface, track, renderOverlays } = mountSurface(engine);
    const field = renderOverlays();
    track();
    engine.open();
    surface.input.focus();
    await settle();
    expect(engine.focused()).toBe(true);
    field.focus();
    await settle();
    expect(engine.focusEvents).toEqual([true, false]);
    expect(engine.focused()).toBe(false);
  }],
  ['swapping the surface with focus inside reports one blur when the new surface opens', async (engine) => {
    const { surface, track, swap } = mountSurface(engine);
    track();
    engine.open();
    surface.input.focus();
    await settle();
    expect(engine.focused()).toBe(true);
    swap();
    engine.open();
    await settle();
    await settle();
    expect(engine.focusEvents).toEqual([true, false]);
    expect(engine.focused()).toBe(false);
  }],
  ['opening while the document lacks focus starts unfocused', async (engine) => {
    const { surface, track } = mountSurface(engine);
    surface.input.focus();
    const hasFocus = spyOn(document, 'hasFocus').mockReturnValue(false);
    try {
      track();
      engine.open();
      await settle();
      expect(engine.focusEvents).toEqual([]);
      expect(engine.focused()).toBe(false);
    } finally {
      hasFocus.mockRestore();
    }
  }],
];

const table = engines.flatMap(([name, make]) => cases.map(([label, run]) => [name, label, make, run] as const));

test.each(table)('%s: %s', async (_engine, _label, make, run) => {
  const engine = make();
  cleanups.push(() => engine.close());
  await run(engine);
});
