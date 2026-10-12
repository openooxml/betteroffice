import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import { en } from '@betteroffice/xlsx-i18n';
import { StrictMode, useEffect, useState } from 'react';
import type { XlsxAsyncEditorProps, XlsxEditorProps } from './XlsxEditor';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import { XlsxAsyncApiUnavailableError, XlsxWorkerEditorCollaborationError } from './worker/createWorkerEditorApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const file = new Uint8Array([1, 2, 3]);
const restorers: (() => void)[] = [];
const warn = mock((..._messages: unknown[]) => {});
let editorInstances = 0;
const mounts = mock(() => {});
const unmounts = mock(() => {});

type MarkerProps = Pick<XlsxEditorProps, 'readOnly' | 'plugins' | 'pluginGrants' | 'onPluginError' | 'onReady'> & { backend?: 'legacy' | 'peerless' };
const workerEditor = mock((props: MarkerProps) => {
  const [instance] = useState(() => ++editorInstances);
  const [backend] = useState(() => props.backend ?? 'legacy');
  useEffect(() => {
    mounts();
    const cleanup = props.onReady?.({} as Parameters<NonNullable<MarkerProps['onReady']>>[0]);
    return () => {
      unmounts();
      if (typeof cleanup === 'function') cleanup();
    };
  }, []);
  return <div data-testid="worker-editor" data-instance={instance}
    data-read-only={props.readOnly ?? false} data-plugins={props.plugins?.length ?? 0} data-backend={backend} />;
});
mock.module('./worker/XlsxWorkerEditor', () => ({ XlsxWorkerEditor: workerEditor }));
mock.module('./viewer/XlsxSessionViewer', () => ({
  XlsxSessionViewer: (props: MarkerProps) => <div data-testid="worker-viewer"
    data-read-only={props.readOnly} data-plugins={props.plugins?.length ?? 0} />,
}));
const { PEERLESS_FLAG, XlsxEditor } = await import('./XlsxEditor');
const asyncProps = { [PEERLESS_FLAG]: true } as const;

function availableWorker() {
  Object.defineProperty(globalThis, 'Worker', { configurable: true, writable: true, value: class {} });
}

function expectInThread(view: ReturnType<typeof render>) {
  expect(view.queryByTestId('worker-editor')).toBeNull();
  expect(view.queryByTestId('worker-viewer')).toBeNull();
  expect(view.getByTestId('xlsx-scroll')).toBeDefined();
}

function expectAsyncRefusal(view: ReturnType<typeof render>) {
  expect(view.container.children).toHaveLength(1);
  expect(view.container.firstElementChild).toBe(view.getByRole('alert'));
  expect(view.container.querySelector('canvas, input, textarea, [contenteditable="true"], [role="grid"], [role="toolbar"]')).toBeNull();
  expect(view.queryByTestId('xlsx-scroll')).toBeNull();
  expect(view.queryByTestId('worker-editor')).toBeNull();
  expect(view.queryByTestId('worker-viewer')).toBeNull();
}

function inThreadWorkbook() {
  const info = { sheetIds: ['sheet:0'], sheetNames: ['Sheet'], activeSheet: 0,
    contentWidth: 4000, contentHeight: 6000, frozenRows: 0, frozenCols: 0, initialScrollX: 0, initialScrollY: 0 };
  const displayList = mock((): xlsx.DisplayList => ({ width: 800, height: 600, commands: [] }));
  const handle = {
    sheetInfo: () => info, sheetCount: () => 1,
    cell: () => ({ a1: 'A1', input: '', isFormula: false }),
    onUpdate: () => () => {}, listProposals: () => [], dispose: mock(() => {}),
    historyState: () => ({ canUndo: true, canRedo: false }),
    selectionFormatting: () => null, mergedRanges: () => [],
    version: () => '0', displayList,
  } as unknown as xlsx.WorkbookHandle;
  const init = spyOn(xlsx, 'initWasm').mockResolvedValue(undefined);
  const open = spyOn(xlsx, 'openWorkbook').mockReturnValue(handle);
  restorers.push(() => open.mockRestore(), () => init.mockRestore());
  return { open, init, displayList };
}

function inThreadPainting() {
  const frames = new Map<number, FrameRequestCallback>();
  let nextFrame = 0;
  const originalRaf = globalThis.requestAnimationFrame;
  const originalCancel = globalThis.cancelAnimationFrame;
  globalThis.requestAnimationFrame = (callback) => { const id = ++nextFrame; frames.set(id, callback); return id; };
  globalThis.cancelAnimationFrame = (id) => { frames.delete(id); };
  restorers.push(() => {
    globalThis.requestAnimationFrame = originalRaf;
    globalThis.cancelAnimationFrame = originalCancel;
  });
  for (const [name, value] of [['clientWidth', 800], ['clientHeight', 600]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, name);
    Object.defineProperty(HTMLElement.prototype, name, { configurable: true, get: () => value });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, name, original);
      else Reflect.deleteProperty(HTMLElement.prototype, name);
    });
  }
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({} as CanvasRenderingContext2D);
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation(() => {});
  restorers.push(() => context.mockRestore(), () => paint.mockRestore());
  return { tick: async () => act(async () => {
    const callbacks = [...frames.values()];
    frames.clear();
    for (const callback of callbacks) callback(0);
  }) };
}

beforeEach(() => {
  const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
  Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, writable: true, value: class {} });
  restorers.push(() => {
    if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
    else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
  });
  const originalWorker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  availableWorker();
  const warning = spyOn(console, 'warn').mockImplementation(warn);
  const init = spyOn(xlsx, 'initWasm').mockImplementation(() => new Promise(() => {}));
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
  restorers.push(() => {
    if (originalWorker) Object.defineProperty(globalThis, 'Worker', originalWorker);
    else Reflect.deleteProperty(globalThis, 'Worker');
  }, () => warning.mockRestore(), () => init.mockRestore(), () => context.mockRestore());
  warn.mockClear();
  workerEditor.mockClear();
  mounts.mockClear();
  unmounts.mockClear();
  editorInstances = 0;
});
afterEach(async () => {
  await act(async () => { cleanup(); });
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('XlsxEditor engine dispatch', () => {
  it('keeps the in-thread editor as the default', () => {
    const view = render(<XlsxEditor file={file} />);
    expectInThread(view);
    expect(workerEditor).not.toHaveBeenCalled();
  });

  it('keeps the worker editor behind the existing hybrid opt-in', () => {
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('legacy');
    expect(warn).not.toHaveBeenCalled();
  });

  it('mounts the async editor for an initially read-only workbook with empty plugins', () => {
    const view = render(<XlsxEditor {...asyncProps} file={file} readOnly plugins={[]} />);
    const marker = view.getByTestId('worker-editor');
    expect(marker.getAttribute('data-backend')).toBe('peerless');
    expect(view.queryByTestId('worker-viewer')).toBeNull();
    view.rerender(<XlsxEditor {...asyncProps} file={file} plugins={[]} />);
    expect(view.getByTestId('worker-editor')).toBe(marker);
    expect(editorInstances).toBe(1);
    expect(warn).not.toHaveBeenCalled();
  });

  it('remounts the editor and cleans up readiness when the async flag changes', () => {
    const cleanupReady = mock(() => {});
    const ready = mock(() => cleanupReady);
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={ready} />);
    const first = view.getByTestId('worker-editor');
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('legacy');
    expect(mounts).toHaveBeenCalledTimes(1);
    view.rerender(<XlsxEditor {...asyncProps} file={file} onReady={ready} />);
    expect(view.getByTestId('worker-editor')).not.toBe(first);
    expect(mounts).toHaveBeenCalledTimes(2);
    expect(unmounts).toHaveBeenCalledTimes(1);
    expect(cleanupReady).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    Reflect.deleteProperty(globalThis, 'Worker');
    view.rerender(<XlsxEditor {...asyncProps} file={file} onReady={ready} />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    expect(mounts).toHaveBeenCalledTimes(2);
    expect(cleanupReady).toHaveBeenCalledTimes(1);
    view.rerender(<XlsxEditor file={file} />);
    expectInThread(view);
    expect(unmounts).toHaveBeenCalledTimes(2);
    expect(cleanupReady).toHaveBeenCalledTimes(2);
    expect(warn).not.toHaveBeenCalled();
    availableWorker();
    view.rerender(<XlsxEditor {...asyncProps} file={file} />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
  });

  it('refuses the async editor without OffscreenCanvas until a new file', () => {
    const original = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
    Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
    restorers.push(() => {
      if (original) Object.defineProperty(globalThis, 'OffscreenCanvas', original);
      else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
    });
    const host = inThreadWorkbook();
    const onError = mock((_error: Error) => {});
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(<XlsxEditor {...asyncProps} file={file} onReady={ready} onError={onError} />);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0]![0]).toMatchObject({
      name: 'XlsxAsyncApiUnavailableError', code: 'async-api-unavailable', message: 'asyncApi requires OffscreenCanvas',
    });
    expectAsyncRefusal(view);
    expect(ready).not.toHaveBeenCalled();
    expect(host.open).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
  });

  it('refuses the async editor without Worker until a new file', () => {
    Reflect.deleteProperty(globalThis, 'Worker');
    const host = inThreadWorkbook();
    const onError = mock((_error: Error) => {});
    const nextError = mock((_error: Error) => {});
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(<XlsxEditor {...asyncProps} file={file} onReady={ready} onError={onError} />);
    expect(onError).toHaveBeenCalledTimes(1);
    const failure = onError.mock.calls[0]![0];
    expect(failure).toBeInstanceOf(XlsxAsyncApiUnavailableError);
    expect(failure).toMatchObject({
      name: 'XlsxAsyncApiUnavailableError', code: 'async-api-unavailable', message: 'asyncApi requires Worker',
    });
    const alert = view.getByRole('alert');
    expect(alert.textContent).toBe(`${en.editor.openError}: ${failure.message}`);
    expectAsyncRefusal(view);
    view.rerender(<XlsxEditor {...asyncProps} file={file} readOnly onReady={ready} onError={nextError} />);
    availableWorker();
    view.rerender(<XlsxEditor {...asyncProps} file={file} onReady={ready} onError={nextError} />);
    expect(view.getByRole('alert')).toBe(alert);
    expectAsyncRefusal(view);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(nextError).not.toHaveBeenCalled();
    expect(ready).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
    expect(host.open).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
    view.rerender(<XlsxEditor {...asyncProps} file={file.slice()} onReady={ready} onError={nextError} />);
    expect(view.queryByRole('alert')).toBeNull();
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    expect(ready).toHaveBeenCalledTimes(1);
    expect(nextError).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
    expect(host.open).not.toHaveBeenCalled();
  });

  for (const withCollaboration of [false, true]) {
    it(`refuses the async editor with an in-thread opt-out (collaboration: ${withCollaboration})`, () => {
      const host = inThreadWorkbook();
      const onError = mock((_error: Error) => {});
      const ready = mock((_api: XlsxAsyncEditorApi) => {});
      const props = { ...asyncProps, file, experimentalWorkerOpen: false,
        collaboration: withCollaboration ? {} : undefined, onReady: ready, onError } as unknown as XlsxAsyncEditorProps;
      const view = render(<XlsxEditor {...props} />);
      expect(onError).toHaveBeenCalledTimes(1);
      expect(onError.mock.calls[0]![0]).toBeInstanceOf(XlsxAsyncApiUnavailableError);
      expect(onError.mock.calls[0]![0]).toMatchObject({
        name: 'XlsxAsyncApiUnavailableError', code: 'async-api-unavailable', message: 'asyncApi requires Worker',
      });
      expect(view.getByRole('alert').textContent).toBe(`${en.editor.openError}: asyncApi requires Worker`);
      expectAsyncRefusal(view);
      view.rerender(<XlsxEditor {...props} readOnly />);
      expectAsyncRefusal(view);
      expect(onError).toHaveBeenCalledTimes(1);
      expect(ready).not.toHaveBeenCalled();
      expect(host.init).not.toHaveBeenCalled();
      expect(host.open).not.toHaveBeenCalled();
      expect(warn).not.toHaveBeenCalled();
    });
  }

  it('refuses unavailable async collaboration with the Worker requirement', () => {
    Reflect.deleteProperty(globalThis, 'Worker');
    const host = inThreadWorkbook();
    const onError = mock((_error: Error) => {});
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(<XlsxEditor {...asyncProps} file={file} collaboration={{}} onReady={ready} onError={onError} />);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0]![0]).toMatchObject({
      name: 'XlsxAsyncApiUnavailableError', code: 'async-api-unavailable', message: 'asyncApi requires Worker',
    });
    expect(onError.mock.calls[0]![0]).not.toBeInstanceOf(XlsxWorkerEditorCollaborationError);
    expect(onError.mock.calls[0]![0]).toBeInstanceOf(XlsxAsyncApiUnavailableError);
    expect(view.getByRole('alert').textContent).toBe(`${en.editor.openError}: asyncApi requires Worker`);
    expectAsyncRefusal(view);
    expect(ready).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
    expect(host.open).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
  });

  for (const change of ['file', 'flag', 'collaboration'] as const) {
    it(`reports async unavailability once per session after a ${change} change`, () => {
      Reflect.deleteProperty(globalThis, 'Worker');
      const host = inThreadWorkbook();
      const onError = mock((_error: Error) => {});
      const ready = mock((_api: XlsxAsyncEditorApi) => {});
      const props = { ...asyncProps, file, onReady: ready, onError };
      const view = render(<XlsxEditor {...props} />);
      expect(onError).toHaveBeenCalledTimes(1);
      const next = { ...props, ...(change === 'file' ? { file: file.slice() }
        : change === 'flag' ? { experimentalWorkerOpen: false } : { collaboration: {} }) } as unknown as XlsxAsyncEditorProps;
      view.rerender(<XlsxEditor {...next} />);
      expect(onError).toHaveBeenCalledTimes(2);
      expect(onError.mock.calls[1]![0]).not.toBe(onError.mock.calls[0]![0]);
      for (const [failure] of onError.mock.calls) {
        expect(failure).toBeInstanceOf(XlsxAsyncApiUnavailableError);
        expect(failure).toMatchObject({ name: 'XlsxAsyncApiUnavailableError', code: 'async-api-unavailable' });
      }
      view.rerender(<XlsxEditor {...next} readOnly />);
      expect(onError).toHaveBeenCalledTimes(2);
      expectAsyncRefusal(view);
      expect(ready).not.toHaveBeenCalled();
      expect(host.init).not.toHaveBeenCalled();
      expect(host.open).not.toHaveBeenCalled();
      expect(warn).not.toHaveBeenCalled();
    });
  }

  it('reports async unavailability once in StrictMode when the callback throws', () => {
    Reflect.deleteProperty(globalThis, 'Worker');
    const host = inThreadWorkbook();
    const onError = mock((_error: Error) => { throw new Error('Callback failed'); });
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(<StrictMode><XlsxEditor {...asyncProps} file={file} onReady={ready} onError={onError} /></StrictMode>);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0]![0]).toMatchObject({ name: 'XlsxAsyncApiUnavailableError', code: 'async-api-unavailable' });
    expectAsyncRefusal(view);
    expect(ready).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
    expect(host.open).not.toHaveBeenCalled();
  });

  it('shows in-thread display failures without reopening', async () => {
    const host = inThreadWorkbook();
    const painting = inThreadPainting();
    const view = render(<XlsxEditor file={file} />);
    await waitFor(() => expect(host.displayList).toHaveBeenCalled());
    const failure = new Error('Rendering failed');
    host.displayList.mockImplementation(() => { throw failure; });
    act(() => fireEvent.scroll(view.getByTestId('xlsx-scroll')));
    await painting.tick();
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('xlsx-render-error').textContent).toContain(failure.message);
  });

  it('opens async collaboration in the peer-less editor when Worker is available', () => {
    const host = inThreadWorkbook();
    const onError = mock((_error: Error) => {});
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(<XlsxEditor {...asyncProps} file={file} collaboration={{}} onReady={ready} onError={onError} />);
    expect(view.queryByRole('alert')).toBeNull();
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    expect(onError).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
    expect(host.open).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
  });

});
