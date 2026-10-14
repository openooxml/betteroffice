import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import { useEffect, useState } from 'react';
import type { XlsxInThreadEditorProps } from './XlsxEditor';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import { xlsxCommandController } from './commands/createXlsxCommandStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const file = new Uint8Array([1, 2, 3]);
const restorers: (() => void)[] = [];
const warn = mock((..._messages: unknown[]) => {});
let editorInstances = 0;
const mounts = mock(() => {});
const unmounts = mock(() => {});

type MarkerProps = Pick<XlsxInThreadEditorProps, 'readOnly' | 'plugins' | 'pluginGrants' | 'onPluginError' | 'onReady'> & { backend?: 'legacy' | 'peerless' };
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
const { resetXlsxEngineWarningsForTests } = await import('./engineWarnings');
const { XlsxEditor } = await import('./XlsxEditor');

function availableWorker() {
  Object.defineProperty(globalThis, 'Worker', { configurable: true, writable: true, value: class {} });
}

function expectInThread(view: ReturnType<typeof render>) {
  expect(view.queryByTestId('worker-editor')).toBeNull();
  expect(view.queryByTestId('worker-viewer')).toBeNull();
  expect(view.getByTestId('xlsx-scroll')).toBeDefined();
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
  return { open, init, displayList, handle };
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
  resetXlsxEngineWarningsForTests();
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
  it('mounts the default worker-only editor without a main-thread engine', () => {
    const host = inThreadWorkbook();
    const view = render(<XlsxEditor file={file} />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    expect(host.open).not.toHaveBeenCalled();
    expect(host.init).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
  });

  it('mounts the hybrid only with legacySyncApi and warns once across remounts', () => {
    const view = render(<XlsxEditor legacySyncApi file={file} />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('legacy');
    view.unmount();
    render(<XlsxEditor legacySyncApi file={file} />);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]?.[0]).toContain('removed in 0.6.0');
  });

  it('keeps experimentalWorkerOpen=true on the async default', () => {
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    expect(warn).not.toHaveBeenCalled();
  });

  for (const reason of ['Worker', 'OffscreenCanvas', 'experimentalWorkerOpen'] as const) {
    it(`falls back to the in-thread async API for ${reason} and warns once`, async () => {
      if (reason !== 'experimentalWorkerOpen') Reflect.deleteProperty(globalThis, reason);
      const host = inThreadWorkbook();
      const ready = mock((_api: XlsxAsyncEditorApi) => {});
      const props = { file, onReady: ready, experimentalWorkerOpen: reason === 'experimentalWorkerOpen' ? false : undefined };
      const view = render(<XlsxEditor {...props} />);
      expectInThread(view);
      await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
      const api = ready.mock.calls[0]![0];
      expect(api.hydrated).toBe(true);
      expect('handle' in api).toBe(false);
      expect('save' in api).toBe(false);
      await api.whenHydrated();
      expect(host.open).toHaveBeenCalledTimes(1);
      view.unmount();
      render(<XlsxEditor {...props} />);
      await waitFor(() => expect(ready).toHaveBeenCalledTimes(2));
      expect(warn).toHaveBeenCalledTimes(1);
      expect(warn.mock.calls[0]?.[0]).toContain('deprecated in-thread engine');
    });
  }

  it('retains the worker-only session when readOnly and plugins change', () => {
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(<XlsxEditor file={file} readOnly onReady={ready} />);
    const editor = view.getByTestId('worker-editor');
    view.rerender(<XlsxEditor file={file} plugins={[]} onReady={ready} />);
    expect(view.getByTestId('worker-editor')).toBe(editor);
    expect(mounts).toHaveBeenCalledTimes(1);
    expect(unmounts).not.toHaveBeenCalled();
  });

  for (const reason of ['Worker', 'OffscreenCanvas', 'experimentalWorkerOpen'] as const) {
    it(`rejects retired fallback commands after workbook replacement (${reason})`, async () => {
      if (reason !== 'experimentalWorkerOpen') Reflect.deleteProperty(globalThis, reason);
      const host = inThreadWorkbook();
      const undo = mock<xlsx.WorkbookHandle['undo']>(() => ({ applied: true, sheetInfo: host.handle.sheetInfo() }));
      const replacement = { ...host.handle, undo } as xlsx.WorkbookHandle;
      const ready = mock((_api: XlsxAsyncEditorApi) => {});
      const props = { onReady: ready, experimentalWorkerOpen: reason === 'experimentalWorkerOpen' ? false : undefined };
      const view = render(<XlsxEditor {...props} file={file} />);
      await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
      const retired = ready.mock.calls[0]![0];
      const prepared = xlsxCommandController(retired.commands)!.prepare('undo');
      host.open.mockReturnValue(replacement);
      view.rerender(<XlsxEditor {...props} file={new Uint8Array([4, 5, 6])} />);
      await waitFor(() => expect(ready).toHaveBeenCalledTimes(2));
      await expect(retired.whenHydrated()).rejects.toMatchObject({ code: 'document-replaced' });
      await expect(retired.commands.execute('undo', null)).rejects.toMatchObject({ code: 'document-replaced' });
      await expect(retired.commands.execute('bold', null)).rejects.toMatchObject({ code: 'document-replaced' });
      await expect(prepared.execute(null)).rejects.toMatchObject({ code: 'document-replaced' });
      expect(() => xlsxCommandController(retired.commands)!.prepare('undo')).toThrow('The workbook was replaced');
      expect(undo).not.toHaveBeenCalled();
      const current = ready.mock.calls[1]![0];
      await act(async () => {
        await expect(current.commands.execute('undo', null)).resolves.toMatchObject({ ok: true });
      });
      expect(undo).toHaveBeenCalledTimes(1);
      view.unmount();
      await expect(current.commands.execute('undo', null)).rejects.toMatchObject({ code: 'document-replaced' });
      expect(undo).toHaveBeenCalledTimes(1);
    });
  }

  it('opens default collaboration in the worker-only editor', () => {
    const host = inThreadWorkbook();
    const view = render(<XlsxEditor file={file} collaboration={{}} />);
    expect(view.getByTestId('worker-editor').getAttribute('data-backend')).toBe('peerless');
    expect(host.open).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
  });

  it('keeps the explicit legacy in-thread API synchronous', async () => {
    inThreadWorkbook();
    const ready = mock((_api: xlsx.WorkbookHandle) => {});
    render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file} onReady={(api) => ready(api.handle)} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(ready.mock.calls[0]![0].version()).toBe('0');
    expect(warn).toHaveBeenCalledTimes(2);
  });

  it('shows in-thread display failures without reopening', async () => {
    const host = inThreadWorkbook();
    const painting = inThreadPainting();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen={false} />);
    await waitFor(() => expect(host.displayList).toHaveBeenCalled());
    const failure = new Error('Rendering failed');
    host.displayList.mockImplementation(() => { throw failure; });
    act(() => fireEvent.scroll(view.getByTestId('xlsx-scroll')));
    await painting.tick();
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('xlsx-render-error').textContent).toContain(failure.message);
  });
});
