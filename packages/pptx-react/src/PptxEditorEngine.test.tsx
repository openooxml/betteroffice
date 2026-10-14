import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { StrictMode, useEffect } from 'react';
import * as pptx from '@betteroffice/pptx/editor';
import type { PptxWorkerViewerProps } from './PptxEditor';
import type { PptxWorkerEditorProps } from './worker/PptxWorkerEditor';
import { definePptxPlugin } from './plugins/definePptxPlugin';
import { resetPptxEngineWarningsForTests } from './engineWarnings';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, within } = await import('@testing-library/react');
const { PptxWorkerEditor: originalEditor } = await import('./worker/PptxWorkerEditor');
const { PptxSessionViewer: originalViewer } = await import('./viewer/PptxSessionViewer');
let editorMounts = 0;

mock.module('./worker/PptxWorkerEditor', () => ({
  PptxWorkerEditor: (props: Omit<PptxWorkerEditorProps, 'readOnly'> & { readOnly?: boolean }) => {
    useEffect(() => { editorMounts += 1; }, []);
    return <div data-testid="worker-editor" data-read-only={props.readOnly ?? false}
      data-plugins={props.plugins?.length ?? 0} data-on-change={props.onChange !== undefined} />;
  },
}));
mock.module('./viewer/PptxSessionViewer', () => ({
  PptxSessionViewer: (props: PptxWorkerViewerProps) =>
    <div data-testid="worker-viewer" data-read-only={props.readOnly}
      data-plugins={props.plugins?.length ?? 0} />,
}));
const { PptxEditor } = await import('./PptxEditor');

const file = new Uint8Array([1]);
const plugin = definePptxPlugin({ id: 'reader', createState: () => null });
const warnings = {
  legacy: '[PptxEditor] `legacySyncApi` selects the deprecated hybrid runtime and sync API, removed in 0.6.0. Omit it to use the async editor.',
  deprecated: '[PptxEditor] `experimentalWorkerOpen={false}` selects the deprecated in-thread engine. Omit `experimentalWorkerOpen` to use the worker-owned editor.',
  collaboration: '[PptxEditor] the legacy hybrid runtime does not support `collaboration`; using the deprecated in-thread engine.',
  worker: '[PptxEditor] `Worker` is not available in this environment; using the deprecated in-thread engine.',
};
const restorers: (() => void)[] = [];
let warn: ReturnType<typeof spyOn<typeof console, 'warn'>>;

function availableWorker(): void {
  Object.defineProperty(globalThis, 'Worker', { configurable: true, writable: true, value: class {} });
}

function expectInThread(view: ReturnType<typeof render>): void {
  const content = within(view.container);
  expect(content.queryByTestId('worker-editor')).toBeNull();
  expect(content.queryByTestId('worker-viewer')).toBeNull();
  expect(content.getByRole('application')).toBeDefined();
}

beforeEach(() => {
  editorMounts = 0;
  resetPptxEngineWarningsForTests();
  const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  availableWorker();
  warn = spyOn(console, 'warn').mockImplementation(() => {});
  const init = spyOn(pptx, 'initWasm').mockResolvedValue(undefined);
  const open = spyOn(pptx, 'openPresentation').mockImplementation(() => { throw new Error('In-thread open'); });
  restorers.push(() => warn.mockRestore(), () => init.mockRestore(), () => open.mockRestore(),
    () => {
      if (worker) Object.defineProperty(globalThis, 'Worker', worker);
      else Reflect.deleteProperty(globalThis, 'Worker');
    });
});
afterEach(async () => {
  await act(async () => {});
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
  resetPptxEngineWarningsForTests();
});
afterAll(async () => {
  mock.module('./worker/PptxWorkerEditor', () => ({ PptxWorkerEditor: originalEditor }));
  mock.module('./viewer/PptxSessionViewer', () => ({ PptxSessionViewer: originalViewer }));
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

for (const experimentalWorkerOpen of [undefined, true] as const) {
  test(`selects the worker editor and viewer with flag ${experimentalWorkerOpen}`, () => {
    const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} experimentalWorkerOpen={experimentalWorkerOpen} />);
    expect(view.getByTestId('worker-editor').getAttribute('data-read-only')).toBe('false');
    view.rerender(<PptxEditor legacySyncApi file={new Uint8Array([2])} fonts={[]} readOnly
      experimentalWorkerOpen={experimentalWorkerOpen} />);
    expect(view.getByTestId('worker-viewer').getAttribute('data-read-only')).toBe('true');
    expect(warn.mock.calls).toEqual([[warnings.legacy]]);
  });
}

test('warns once for the deprecated engine across mounts and rerenders', () => {
  const editor = <PptxEditor legacySyncApi file={file} fonts={[]} experimentalWorkerOpen={false} />;
  const first = render(editor);
  expectInThread(first);
  first.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} experimentalWorkerOpen={false} readOnly />);
  expectInThread(first);
  first.unmount();
  expectInThread(render(editor));
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.deprecated]]);
});

test('warns once in StrictMode', () => {
  const view = render(<StrictMode><PptxEditor legacySyncApi file={file} fonts={[]} experimentalWorkerOpen={false} /></StrictMode>);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.deprecated]]);
});

test('falls back for collaboration before checking Worker', () => {
  Reflect.deleteProperty(globalThis, 'Worker');
  const collaboration = { clientId: 1 };
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} collaboration={collaboration} readOnly />);
  expectInThread(view);
  const second = render(<PptxEditor legacySyncApi file={file} fonts={[]} collaboration={collaboration} experimentalWorkerOpen />);
  expectInThread(second);
  second.unmount();
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} collaboration={collaboration} />);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.collaboration]]);
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} collaboration={collaboration} experimentalWorkerOpen={false} />);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.collaboration], [warnings.deprecated]]);
});

test('pins the missing-Worker fallback until the file changes', () => {
  Reflect.deleteProperty(globalThis, 'Worker');
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} />);
  expectInThread(view);
  const second = render(<PptxEditor legacySyncApi file={file} fonts={[]} experimentalWorkerOpen />);
  expectInThread(second);
  second.unmount();
  availableWorker();
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} />);
  expectInThread(view);
  view.rerender(<PptxEditor legacySyncApi file={new Uint8Array([2])} fonts={[]} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.worker]]);
});

test('pins the worker engine until the file changes', () => {
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} />);
  Reflect.deleteProperty(globalThis, 'Worker');
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(1);
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
  view.rerender(<PptxEditor legacySyncApi file={new Uint8Array([2])} fonts={[]} />);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.worker]]);
});

test('keeps the missing-Worker fallback across clientId and font changes', () => {
  Reflect.deleteProperty(globalThis, 'Worker');
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} clientId={1} />);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.worker]]);
  availableWorker();
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} clientId={2} />);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.worker]]);
  const fonts = [{ family: 'Test', bytes: new Uint8Array([3]) }];
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={fonts} clientId={2} />);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.worker]]);
  view.rerender(<PptxEditor legacySyncApi file={new Uint8Array([2])} fonts={fonts} clientId={2} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.worker]]);
});

test('hosts plugins in a read-only worker editor and promotes a viewer when plugins arrive', () => {
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} readOnly plugins={[plugin]} />);
  const marker = view.getByTestId('worker-editor');
  expect(marker.getAttribute('data-read-only')).toBe('true');
  expect(marker.getAttribute('data-plugins')).toBe('1');
  const next = new Uint8Array([2]);
  view.rerender(<PptxEditor legacySyncApi file={next} fonts={[]} readOnly plugins={[]} />);
  expect(view.getByTestId('worker-viewer')).toBeDefined();
  view.rerender(<PptxEditor legacySyncApi file={next} fonts={[]} readOnly plugins={[plugin]} />);
  expect(view.getByTestId('worker-editor').getAttribute('data-plugins')).toBe('1');
  view.rerender(<PptxEditor legacySyncApi file={next} fonts={[]} readOnly />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(2);
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
});

test('keeps read-only onChange in the worker editor and promotes a viewer when it arrives', () => {
  const onChange = mock(() => {});
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} readOnly onChange={onChange} />);
  expect(view.getByTestId('worker-editor').getAttribute('data-read-only')).toBe('true');
  expect(view.getByTestId('worker-editor').getAttribute('data-on-change')).toBe('true');
  const next = new Uint8Array([2]);
  view.rerender(<PptxEditor legacySyncApi file={next} fonts={[]} readOnly />);
  expect(view.getByTestId('worker-viewer')).toBeDefined();
  view.rerender(<PptxEditor legacySyncApi file={next} fonts={[]} readOnly onChange={onChange} />);
  expect(view.getByTestId('worker-editor').getAttribute('data-on-change')).toBe('true');
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
});

test('treats an absent flag and true as the same session', () => {
  const view = render(<PptxEditor legacySyncApi file={file} fonts={[]} />);
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} experimentalWorkerOpen />);
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={[]} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(1);
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
});

test('re-picks the engine when the false opt-out toggles on the same file', () => {
  const fonts: PptxWorkerEditorProps['fonts'] = [];
  const view = render(<PptxEditor legacySyncApi file={file} fonts={fonts} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(1);
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={fonts} experimentalWorkerOpen={false} />);
  expectInThread(view);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.deprecated]]);
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={fonts} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(2);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.deprecated]]);
});

test('re-picks the engine when collaboration is added, replaced and removed on the same file', async () => {
  const fonts: PptxWorkerEditorProps['fonts'] = [];
  const view = render(<PptxEditor legacySyncApi file={file} fonts={fonts} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(1);
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
  await act(async () => {
    view.rerender(<PptxEditor legacySyncApi file={file} fonts={fonts} collaboration={{ clientId: 1 }} />);
  });
  expectInThread(view);
  expect(pptx.openPresentation).toHaveBeenCalledTimes(1);
  expect(pptx.openPresentation).toHaveBeenLastCalledWith(file, {
    clientId: 1, fonts, initialUpdate: undefined,
  });
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.collaboration]]);
  await act(async () => {
    view.rerender(<PptxEditor legacySyncApi file={file} fonts={fonts} collaboration={{ clientId: 2 }} />);
  });
  expectInThread(view);
  expect(pptx.openPresentation).toHaveBeenCalledTimes(2);
  expect(pptx.openPresentation).toHaveBeenLastCalledWith(file, {
    clientId: 2, fonts, initialUpdate: undefined,
  });
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.collaboration]]);
  view.rerender(<PptxEditor legacySyncApi file={file} fonts={fonts} />);
  expect(view.getByTestId('worker-editor')).toBeDefined();
  expect(editorMounts).toBe(2);
  expect(warn.mock.calls).toEqual([[warnings.legacy], [warnings.collaboration]]);
});

test('legacy sync API warns once across hybrid mounts and StrictMode', () => {
  const editor = <StrictMode><PptxEditor legacySyncApi file={file} fonts={[]} /></StrictMode>;
  const first = render(editor);
  expect(first.getByTestId('worker-editor')).toBeDefined();
  first.rerender(editor);
  first.unmount();
  expect(render(editor).getByTestId('worker-editor')).toBeDefined();
  expect(warn.mock.calls).toEqual([[warnings.legacy]]);
});
