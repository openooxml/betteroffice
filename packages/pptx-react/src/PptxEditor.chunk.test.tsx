import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { Profiler, StrictMode } from 'react';
import * as React from 'react';
import { PptxEditor, PptxWorkerEditorChunkLoadError, pptxPeerlessEditorLoader } from './PptxEditor';

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
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
globalThis.setTimeout = realSetTimeout;
const originalLoad = pptxPeerlessEditorLoader.load;
const capabilities = new Map<string, PropertyDescriptor | undefined>();
const props = { file: new Uint8Array([1]), fonts: [], i18n: { worker: { preparing: 'Preparing test editor' } } };
type SurfaceModule = Awaited<ReturnType<typeof pptxPeerlessEditorLoader.load>>;

beforeEach(() => {
  for (const name of ['Worker', 'OffscreenCanvas']) {
    capabilities.set(name, Object.getOwnPropertyDescriptor(globalThis, name));
    Object.defineProperty(globalThis, name, { configurable: true, value: class {} });
  }
});
afterEach(() => {
  holdRevealTimers = false;
  revealDelays.length = 0;
  cleanup();
  mock.restore();
  pptxPeerlessEditorLoader.load = originalLoad;
  for (const [name, descriptor] of capabilities) {
    if (descriptor) Object.defineProperty(globalThis, name, descriptor);
    else Reflect.deleteProperty(globalThis, name);
  }
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

function surface(label: string): SurfaceModule {
  return { PeerlessPptxSurface: () => <div data-testid="loaded-peerless-surface">{label}</div> };
}

function chunk() {
  let resolve!: (value: SurfaceModule) => void;
  let reject!: (cause: unknown) => void;
  const promise = new Promise<SurfaceModule>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

test('a cold chunk commits the surface without advancing reveal timers', async () => {
  const pending = chunk();
  const load = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  spyOn(performance, 'now').mockReturnValue(0);
  spyOn(console, 'error').mockImplementation(() => {});
  holdRevealTimers = true;
  const container = document.createElement('div');
  document.body.append(container);
  const commits: string[] = [];
  const view = render(<Profiler id="cold" onRender={() => {
    commits.push(container.querySelector('[data-testid="loaded-peerless-surface"]') ? 'surface' : 'loading');
  }}><PptxEditor {...props} className="cold-editor" /></Profiler>, { container });
  expect(view.getByRole('status').outerHTML).toBe('<div class="cold-editor" role="status">Preparing test editor</div>');
  expect(commits.every((commit) => commit === 'loading')).toBe(true);
  const loadingCommits = commits.length;
  pending.resolve(surface('Cold editor'));
  for (let turn = 0; turn < 6; turn++) await new Promise<void>((resolve) => setImmediate(resolve));
  expect(view.queryByTestId('loaded-peerless-surface')?.textContent).toBe('Cold editor');
  expect(commits.slice(loadingCommits)).toEqual(['surface']);
  expect(revealDelays).toEqual([]);
  expect(load).toHaveBeenCalledTimes(1);
});

for (const outcome of ['resolve', 'reject'] as const) {
  test(`unmount during a chunk load ignores its ${outcome}`, async () => {
    const pending = chunk();
    spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
    const onError = mock((_error: Error) => {});
    const errors = spyOn(console, 'error').mockImplementation(() => {});
    let unmounted = false;
    let updatesAfterUnmount = 0;
    const useState = React.useState;
    spyOn(React, 'useState').mockImplementation(((initial: unknown) => {
      const [value, update] = useState(initial);
      return [value, (next: unknown) => {
        if (unmounted) updatesAfterUnmount++;
        update(next);
      }];
    }) as typeof React.useState);
    const Surface = mock(() => <span data-testid="loaded-peerless-surface">Unmounted editor</span>);
    const view = render(<PptxEditor {...props} onError={onError} />);
    expect(view.getByRole('status').textContent).toBe('Preparing test editor');
    view.unmount();
    unmounted = true;
    await act(async () => {
      if (outcome === 'resolve') pending.resolve({ PeerlessPptxSurface: Surface });
      else pending.reject(new TypeError('Unmounted chunk failure'));
    });
    expect(updatesAfterUnmount).toBe(0);
    expect(Surface).not.toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
    expect(errors).not.toHaveBeenCalled();
  });
}

test('StrictMode chunk loading renders one surface', async () => {
  const pending = chunk();
  const load = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  const view = render(<StrictMode><PptxEditor {...props} /></StrictMode>);
  expect(view.getByRole('status').textContent).toBe('Preparing test editor');
  expect(load.mock.calls.length).toBeGreaterThanOrEqual(1);
  expect(load.mock.calls.length).toBeLessThanOrEqual(2);
  await act(async () => { pending.resolve(surface('Strict editor')); });
  expect(view.queryAllByTestId('loaded-peerless-surface')).toHaveLength(1);
  expect(view.getByTestId('loaded-peerless-surface').textContent).toBe('Strict editor');
});

for (const change of ['loader', 'file'] as const) {
  test(`a stale chunk cannot replace or fail the current editor after a ${change} change`, async () => {
    const stale = chunk();
    const current = chunk();
    const load = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValueOnce(stale.promise).mockReturnValue(current.promise);
    const onError = mock((_error: Error) => {});
    spyOn(console, 'error').mockImplementation(() => {});
    const view = render(<PptxEditor {...props} onError={onError} />);
    const next = change === 'file' ? { ...props, file: new Uint8Array([2]) } : props;
    if (change === 'loader') pptxPeerlessEditorLoader.load = mock(() => current.promise);
    view.rerender(<PptxEditor {...next} onError={onError} />);
    expect(view.getByRole('status').textContent).toBe('Preparing test editor');
    await act(async () => { current.resolve(surface('Current editor')); });
    expect(view.queryByTestId('loaded-peerless-surface')?.textContent).toBe('Current editor');
    await act(async () => { stale.reject(new TypeError('Stale chunk failure')); });
    expect(view.queryByTestId('loaded-peerless-surface')?.textContent).toBe('Current editor');
    expect(onError).not.toHaveBeenCalled();
    view.unmount();
    const remount = render(<PptxEditor {...next} onError={onError} />);
    expect(remount.queryByTestId('loaded-peerless-surface')?.textContent).toBe('Current editor');
    expect(load).toHaveBeenCalledTimes(change === 'file' ? 2 : 1);
  });
}

test('a successful chunk renders synchronously on remount without a preparing fallback commit', async () => {
  const pending = chunk();
  const load = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  const first = render(<PptxEditor {...props} />);
  expect(first.getByRole('status').textContent).toBe('Preparing test editor');
  await act(async () => { pending.resolve(surface('Loaded editor')); });
  expect(first.getByTestId('loaded-peerless-surface').textContent).toBe('Loaded editor');
  first.unmount();

  const container = document.createElement('div');
  document.body.append(container);
  const preparingCommits: string[] = [];
  const second = render(<Profiler id="remount" onRender={() => {
    const status = container.querySelector('[role="status"]');
    if (status) preparingCommits.push(status.textContent ?? '');
  }}><PptxEditor {...props} /></Profiler>, { container });
  expect(preparingCommits).toEqual([]);
  expect(second.queryByTestId('loaded-peerless-surface') !== null).toBe(true);
  expect(second.queryByRole('status') === null).toBe(true);
  await act(async () => {});
  expect(preparingCommits).toEqual([]);
  expect(load).toHaveBeenCalledTimes(1);
});

test('a failed chunk reports its error and retries the same loader on remount', async () => {
  const cause = new TypeError('Chunk fetch failed');
  const load = spyOn(pptxPeerlessEditorLoader, 'load').mockRejectedValueOnce(cause)
    .mockResolvedValue(surface('Retry opened'));
  const onError = mock((_error: Error) => { throw new Error('Host callback failed'); });
  spyOn(console, 'error').mockImplementation(() => {});
  const failed = render(<StrictMode><PptxEditor {...props} onError={onError} /></StrictMode>);
  await waitFor(() => expect(failed.queryByRole('alert') !== null).toBe(true));
  expect(failed.getByRole('alert').textContent).toContain('Remount the editor to retry');
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]![0]).toBeInstanceOf(PptxWorkerEditorChunkLoadError);
  expect(onError.mock.calls[0]![0]).toMatchObject({ code: 'chunk-load-failed', cause });
  failed.unmount();

  const retry = render(<PptxEditor {...props} onError={onError} />);
  await waitFor(() => expect(retry.queryByTestId('loaded-peerless-surface')?.textContent).toBe('Retry opened'));
  expect(load).toHaveBeenCalledTimes(2);
  expect(onError).toHaveBeenCalledTimes(1);
});

test('a replacement loader gets a fresh chunk after an earlier loader succeeds', async () => {
  const firstLoad = spyOn(pptxPeerlessEditorLoader, 'load').mockResolvedValue(surface('First loader'));
  const first = render(<PptxEditor {...props} />);
  await waitFor(() => expect(first.queryByTestId('loaded-peerless-surface')?.textContent).toBe('First loader'));
  first.unmount();

  const pending = chunk();
  const replacement = mock(() => pending.promise);
  pptxPeerlessEditorLoader.load = replacement;
  const second = render(<PptxEditor {...props} />);
  expect(second.getByRole('status').textContent).toBe('Preparing test editor');
  expect(second.queryByTestId('loaded-peerless-surface') === null).toBe(true);
  expect(replacement).toHaveBeenCalledTimes(1);
  await act(async () => { pending.resolve(surface('Replacement loader')); });
  expect(second.getByTestId('loaded-peerless-surface').textContent).toBe('Replacement loader');
  expect(firstLoad).toHaveBeenCalledTimes(1);
});
