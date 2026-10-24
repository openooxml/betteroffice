import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { Profiler, StrictMode } from 'react';
import * as React from 'react';
import { XlsxWorkerEditor, XlsxWorkerEditorChunkLoadError, xlsxPeerlessEditorLoader } from './XlsxWorkerEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const realSetTimeout = globalThis.setTimeout;
let holdRevealTimers = false;
const revealDelays: number[] = [];
globalThis.setTimeout = ((...args: Parameters<typeof setTimeout>) => {
  const delay = args[1];
  if (holdRevealTimers && delay !== undefined && delay >= 100 && delay <= 300) {
    revealDelays.push(delay);
    return realSetTimeout(() => {}, 0);
  }
  return realSetTimeout(...args);
}) as typeof setTimeout;
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
globalThis.setTimeout = realSetTimeout;
const originalLoad = xlsxPeerlessEditorLoader.load;
const props = { backend: 'peerless' as const, file: new Uint8Array([1]) };
type EditorModule = Awaited<ReturnType<typeof xlsxPeerlessEditorLoader.load>>;

afterEach(() => {
  holdRevealTimers = false;
  revealDelays.length = 0;
  cleanup();
  mock.restore();
  xlsxPeerlessEditorLoader.load = originalLoad;
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

function editor(label: string): EditorModule {
  return { XlsxPeerlessWorkerEditor: () => <div data-testid="loaded-peerless-editor">{label}</div> };
}

function chunk() {
  let resolve!: (value: EditorModule) => void;
  let reject!: (cause: unknown) => void;
  const promise = new Promise<EditorModule>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function commitLog() {
  const container = document.body.appendChild(document.createElement('div'));
  const commits: string[] = [];
  const onRender = () => {
    commits.push(container.querySelector('[data-testid="loaded-peerless-editor"]') ? 'editor'
      : container.childNodes.length === 0 ? 'fallback' : 'other');
  };
  return { container, commits, onRender };
}

test('a cold chunk commits the editor without a reveal delay', async () => {
  const pending = chunk();
  const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  spyOn(performance, 'now').mockReturnValue(0);
  spyOn(console, 'error').mockImplementation(() => {});
  holdRevealTimers = true;
  const log = commitLog();
  const view = render(<Profiler id="cold" onRender={log.onRender}><XlsxWorkerEditor {...props} /></Profiler>,
    { container: log.container });
  expect(log.commits.length).toBeGreaterThan(0);
  expect(log.commits.every((commit) => commit === 'fallback')).toBe(true);
  const fallbackCommits = log.commits.length;
  pending.resolve(editor('Cold editor'));
  for (let turn = 0; turn < 6; turn++) await new Promise<void>((resolve) => setImmediate(resolve));
  expect(view.queryByTestId('loaded-peerless-editor')?.textContent).toBe('Cold editor');
  expect(log.commits.slice(fallbackCommits)).toEqual(['editor']);
  expect(revealDelays).toEqual([]);
  expect(load).toHaveBeenCalledTimes(1);
});

test('a loaded chunk renders the editor in the first commit of a remount', async () => {
  const pending = chunk();
  const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  const first = render(<XlsxWorkerEditor {...props} />);
  await act(async () => { pending.resolve(editor('Loaded editor')); });
  expect(first.getByTestId('loaded-peerless-editor').textContent).toBe('Loaded editor');
  first.unmount();

  const log = commitLog();
  const second = render(<Profiler id="remount" onRender={log.onRender}><XlsxWorkerEditor {...props} /></Profiler>,
    { container: log.container });
  expect(log.commits).toEqual(['editor']);
  await act(async () => {});
  expect(log.commits).toEqual(['editor']);
  expect(second.getByTestId('loaded-peerless-editor').textContent).toBe('Loaded editor');
  expect(load).toHaveBeenCalledTimes(1);
});

for (const failure of ['rejected', 'thrown'] as const) {
  test(`a ${failure} chunk load reports once and retries the same loader on remount`, async () => {
    const cause = new TypeError('Chunk fetch failed');
    const load = spyOn(xlsxPeerlessEditorLoader, 'load');
    if (failure === 'rejected') load.mockRejectedValueOnce(cause);
    else load.mockImplementationOnce(() => { throw cause; });
    load.mockResolvedValue(editor('Retry opened'));
    const onError = mock((_error: Error) => { throw new Error('Host callback failed'); });
    spyOn(console, 'error').mockImplementation(() => {});
    const failed = render(<StrictMode><XlsxWorkerEditor {...props} onError={onError} /></StrictMode>);
    await waitFor(() => expect(failed.queryByTestId('xlsx-error') !== null).toBe(true));
    expect(failed.getByTestId('xlsx-error').getAttribute('role')).toBe('alert');
    expect(failed.getByRole('alert').textContent).toContain("The editor couldn't load");
    expect(failed.getByTestId('xlsx-error').querySelector('details pre')?.textContent).toContain('Remount the editor to retry');
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0][0]).toBeInstanceOf(XlsxWorkerEditorChunkLoadError);
    expect(onError.mock.calls[0][0]).toMatchObject({ code: 'chunk-load-failed', cause });
    failed.unmount();

    const retry = render(<XlsxWorkerEditor {...props} onError={onError} />);
    await waitFor(() => expect(retry.queryByTestId('loaded-peerless-editor')?.textContent).toBe('Retry opened'));
    expect(load).toHaveBeenCalledTimes(2);
    expect(onError).toHaveBeenCalledTimes(1);
  });
}

test('Retry on a failed chunk load remounts the editor', async () => {
  const load = spyOn(xlsxPeerlessEditorLoader, 'load');
  load.mockRejectedValueOnce(new TypeError('Chunk fetch failed'));
  load.mockResolvedValue(editor('Retry opened'));
  spyOn(console, 'error').mockImplementation(() => {});
  const onError = mock((_error: Error) => {});
  const view = render(<XlsxWorkerEditor {...props} onError={onError} />);
  await waitFor(() => expect(view.queryByTestId('xlsx-error') !== null).toBe(true));
  expect(view.getByTestId('xlsx-error').getAttribute('data-failure')).toBe('load');
  await act(async () => { fireEvent.click(view.getByRole('button', { name: 'Retry' })); });
  await waitFor(() => expect(view.queryByTestId('loaded-peerless-editor')?.textContent).toBe('Retry opened'));
  expect(view.queryByTestId('xlsx-error')).toBeNull();
  expect(load).toHaveBeenCalledTimes(2);
  expect(onError).toHaveBeenCalledTimes(1);
});

for (const outcome of ['resolve', 'reject'] as const) {
  test(`unmount during a chunk load ignores its ${outcome}`, async () => {
    const pending = chunk();
    spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
    const onError = mock((_error: Error) => {});
    const errors = spyOn(console, 'error').mockImplementation(() => {});
    let unmounted = false;
    let stateHooks = 0;
    let updatesAfterUnmount = 0;
    const useState = React.useState;
    spyOn(React, 'useState').mockImplementation(((initial: unknown) => {
      stateHooks++;
      const [value, update] = useState(initial);
      return [value, (next: unknown) => {
        if (unmounted) updatesAfterUnmount++;
        update(next);
      }];
    }) as typeof React.useState);
    const Editor = mock(() => <span data-testid="loaded-peerless-editor">Unmounted editor</span>);
    const view = render(<XlsxWorkerEditor {...props} onError={onError} />);
    expect(view.container.childNodes.length).toBe(0);
    expect(stateHooks).toBeGreaterThan(0);
    view.unmount();
    unmounted = true;
    await act(async () => {
      if (outcome === 'resolve') pending.resolve({ XlsxPeerlessWorkerEditor: Editor });
      else pending.reject(new TypeError('Unmounted chunk failure'));
    });
    expect(updatesAfterUnmount).toBe(0);
    expect(Editor).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
    expect(errors).not.toHaveBeenCalled();
  });
}

test('StrictMode loads the chunk once and renders one editor', async () => {
  const pending = chunk();
  const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  const view = render(<StrictMode><XlsxWorkerEditor {...props} /></StrictMode>);
  expect(view.container.childNodes.length).toBe(0);
  await act(async () => { pending.resolve(editor('Strict editor')); });
  expect(view.queryAllByTestId('loaded-peerless-editor')).toHaveLength(1);
  expect(view.getByTestId('loaded-peerless-editor').textContent).toBe('Strict editor');
  expect(load).toHaveBeenCalledTimes(1);
});

test('a replacement loader gets a fresh load after an earlier loader succeeded', async () => {
  const firstLoad = spyOn(xlsxPeerlessEditorLoader, 'load').mockResolvedValue(editor('First loader'));
  const first = render(<XlsxWorkerEditor {...props} />);
  await waitFor(() => expect(first.queryByTestId('loaded-peerless-editor')?.textContent).toBe('First loader'));
  first.unmount();

  const pending = chunk();
  const replacement = mock(() => pending.promise);
  xlsxPeerlessEditorLoader.load = replacement;
  const second = render(<XlsxWorkerEditor {...props} />);
  expect(second.container.childNodes.length).toBe(0);
  expect(replacement).toHaveBeenCalledTimes(1);
  await act(async () => { pending.resolve(editor('Replacement loader')); });
  expect(second.getByTestId('loaded-peerless-editor').textContent).toBe('Replacement loader');
  expect(firstLoad).toHaveBeenCalledTimes(1);
});

for (const outcome of ['resolve', 'reject'] as const) {
  test(`a stale load that later ${outcome}s cannot replace or fail the editor after a loader change`, async () => {
    const stale = chunk();
    const current = chunk();
    const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(stale.promise);
    const onError = mock((_error: Error) => {});
    spyOn(console, 'error').mockImplementation(() => {});
    const view = render(<XlsxWorkerEditor {...props} onError={onError} />);
    const replacement = mock(() => current.promise);
    xlsxPeerlessEditorLoader.load = replacement;
    view.rerender(<XlsxWorkerEditor {...props} onError={onError} />);
    expect(view.container.childNodes.length).toBe(0);
    await act(async () => { current.resolve(editor('Current editor')); });
    expect(view.queryByTestId('loaded-peerless-editor')?.textContent).toBe('Current editor');
    await act(async () => {
      if (outcome === 'resolve') stale.resolve(editor('Stale editor'));
      else stale.reject(new TypeError('Stale chunk failure'));
    });
    expect(view.queryByTestId('loaded-peerless-editor')?.textContent).toBe('Current editor');
    expect(onError).not.toHaveBeenCalled();
    view.unmount();

    const remount = render(<XlsxWorkerEditor {...props} onError={onError} />);
    expect(remount.queryByTestId('loaded-peerless-editor')?.textContent).toBe('Current editor');
    expect(load).toHaveBeenCalledTimes(1);
    expect(replacement).toHaveBeenCalledTimes(1);
  });
}

test('renders the supplied indicator while a file chunk loads', async () => {
  const pending = chunk();
  spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  const view = render(<XlsxWorkerEditor {...props} loadingIndicator={<span>Opening workbook</span>} />);
  expect(view.queryByText('Opening workbook')).not.toBeNull();
  await act(async () => { pending.resolve(editor('Opened')); });
  expect(view.queryByText('Opening workbook')).toBeNull();
});

test('merges style on a chunk error', async () => {
  spyOn(xlsxPeerlessEditorLoader, 'load').mockRejectedValue(new Error('Unavailable chunk'));
  spyOn(console, 'error').mockImplementation(() => {});
  const view = render(<XlsxWorkerEditor {...props} style={{ height: 123 }} />);
  await waitFor(() => expect(view.queryByTestId('xlsx-error')).not.toBeNull());
  expect(view.getByTestId('xlsx-error').style.height).toBe('123px');
});
