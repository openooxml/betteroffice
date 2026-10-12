import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { Profiler, StrictMode } from 'react';
import { PptxEditor, PptxWorkerEditorChunkLoadError, pptxPeerlessEditorLoader } from './PptxEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
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
  const promise = new Promise<SurfaceModule>((yes) => { resolve = yes; });
  return { promise, resolve };
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
