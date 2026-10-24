import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { createRef, forwardRef, lazy, Profiler, startTransition, StrictMode, Suspense, useImperativeHandle, type ComponentType, type ReactElement } from 'react';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import { docxPeerlessEditorLoader, DocxWorkerEditorChunkLoadError, LazyPeerlessDocxEditor } from './PeerlessDocxEditorLoader';

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
const originalLoad = docxPeerlessEditorLoader.load;

afterEach(() => {
  holdRevealTimers = false;
  revealDelays.length = 0;
  cleanup();
  mock.restore();
  docxPeerlessEditorLoader.load = originalLoad;
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const loaded = {
  PeerlessDocxEditor: forwardRef<DocxEditorRef, DocxEditorProps>(function Editor(props, ref) {
    useImperativeHandle(ref, () => ({ focus() {} }) as DocxEditorRef);
    return <span>{props.documentName}</span>;
  }),
};
type EditorModule = Awaited<ReturnType<typeof docxPeerlessEditorLoader.load>>;

function chunk() {
  let resolve!: (value: EditorModule) => void;
  let reject!: (cause: unknown) => void;
  const promise = new Promise<EditorModule>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function profiled(id: string, element: ReactElement) {
  const container = document.body.appendChild(document.createElement('div'));
  const commits: string[] = [];
  const view = render(<Profiler id={id} onRender={() => { commits.push(container.textContent ?? ''); }}>{element}</Profiler>, { container });
  return { view, commits };
}

test('chunk loading retains custom loading content and forwards editor props and ref', async () => {
  let resolve!: (value: typeof loaded) => void;
  const load = spyOn(docxPeerlessEditorLoader, 'load').mockImplementation(() => new Promise((yes) => { resolve = yes; }));
  const ref = createRef<DocxEditorRef>();
  const view = render(<LazyPeerlessDocxEditor ref={ref} documentName="Loaded editor" loadingIndicator={<span>Loading chunk</span>} />);
  expect(view.getByText('Loading chunk')).toBeTruthy();
  expect(ref.current).toBeNull();
  await act(async () => { resolve(loaded); });
  expect(view.getByText('Loaded editor')).toBeTruthy();
  expect(typeof ref.current?.focus).toBe('function');
  expect(load).toHaveBeenCalledTimes(1);
});

test('StrictMode imports a pending chunk once', async () => {
  const pending = chunk();
  const load = spyOn(docxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  const view = render(<StrictMode><LazyPeerlessDocxEditor documentName="Strict editor" /></StrictMode>);
  expect(load).toHaveBeenCalledTimes(1);
  await act(async () => { pending.resolve(loaded); });
  expect(view.getByText('Strict editor')).toBeTruthy();
  expect(load).toHaveBeenCalledTimes(1);
});

test('a discarded loader transition reports the committed import failure once without retrying', async () => {
  const pending = chunk();
  const retry = chunk();
  const replacement = chunk();
  const loadA = mock(() => retry.promise).mockReturnValueOnce(pending.promise);
  const loadB = mock(() => replacement.promise);
  docxPeerlessEditorLoader.load = loadA;
  const onError = mock((_error: Error) => {});
  const suspended = mock(() => {});
  const never = new Promise<never>(() => {});
  spyOn(console, 'error').mockImplementation(() => {});
  function Block({ active }: { active: boolean }) {
    if (active) { suspended(); throw never; }
    return null;
  }
  const tree = (active: boolean) => <StrictMode><Suspense fallback={<span>Host suspended</span>}>
    <LazyPeerlessDocxEditor onError={onError} loadingIndicator={<span>Loading chunk</span>} />
    <Block active={active} />
  </Suspense></StrictMode>;
  const committed = tree(false);
  const view = render(committed);
  expect(view.getByText('Loading chunk')).toBeTruthy();
  expect(loadA).toHaveBeenCalledTimes(1);
  await act(async () => {
    docxPeerlessEditorLoader.load = loadB;
    startTransition(() => { view.rerender(tree(true)); });
  });
  expect(loadB).toHaveBeenCalledTimes(1);
  expect(suspended).toHaveBeenCalled();
  expect(view.queryByText('Host suspended')).toBeNull();
  docxPeerlessEditorLoader.load = loadA;
  view.rerender(committed);
  const cause = new Error('Committed chunk failed');
  await act(async () => { pending.reject(cause); });
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]![0]).toBeInstanceOf(DocxWorkerEditorChunkLoadError);
  expect(onError.mock.calls[0]![0]).toMatchObject({ code: 'chunk-load-failed', cause });
  expect(view.getByRole('alert')).toBeTruthy();
  expect(loadA).toHaveBeenCalledTimes(1);
});

test('a failed chunk reports once, contains throwing callbacks, and retries on remount', async () => {
  const cause = new Error('chunk unavailable');
  const load = spyOn(docxPeerlessEditorLoader, 'load').mockRejectedValue(cause);
  const onError = mock((_error: Error) => { throw new Error('host callback failed'); });
  spyOn(console, 'error').mockImplementation(() => {});
  const view = render(<StrictMode><LazyPeerlessDocxEditor onError={onError} /></StrictMode>);
  await waitFor(() => expect(view.getByRole('alert')).toBeTruthy());
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]![0]).toMatchObject({ name: 'DocxWorkerEditorChunkLoadError', code: 'chunk-load-failed', cause });
  expect(onError.mock.calls[0]![0]).toBeInstanceOf(DocxWorkerEditorChunkLoadError);
  view.unmount();
  load.mockResolvedValue(loaded);
  const retry = render(<LazyPeerlessDocxEditor documentName="Retry opened" />);
  await waitFor(() => expect(retry.getByText('Retry opened')).toBeTruthy());
  expect(load).toHaveBeenCalledTimes(2);
});

test('a failed chunk offers a page reload, since the browser keeps failing the same chunk URL', async () => {
  const specifier = './PeerlessDocxEditor.js';
  const failedFetches = new Set<string>();
  const imported: string[] = [];
  let online = false;
  const importModule = (url: string): Promise<EditorModule> => {
    imported.push(url);
    if (!online || failedFetches.has(url)) {
      failedFetches.add(url);
      return Promise.reject(new TypeError(`Failed to fetch dynamically imported module: ${url}`));
    }
    return Promise.resolve(loaded);
  };
  spyOn(docxPeerlessEditorLoader, 'load').mockImplementation(() => importModule(specifier));
  const reload = spyOn(window.location, 'reload').mockImplementation(() => {});
  const onError = mock((_error: Error) => {});
  spyOn(console, 'error').mockImplementation(() => {});
  const view = render(<LazyPeerlessDocxEditor documentName="Reloaded editor" onError={onError} />);
  await waitFor(() => expect(view.getByRole('heading', { name: "The editor couldn't load" })).toBeTruthy());
  expect(view.getByText('Check your connection, then reload the page.')).toBeTruthy();
  expect(view.getByText('Failed to load the editor code. Reload the page to retry.')).toBeTruthy();
  expect(view.queryByRole('button', { name: 'Retry' })).toBeNull();
  online = true;
  await expect(importModule(specifier)).rejects.toThrow(specifier);
  imported.length = 0;
  await act(async () => { fireEvent.click(view.getByRole('button', { name: 'Reload page' })); });
  expect(reload).toHaveBeenCalledTimes(1);
  expect(imported).toEqual([]);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]![0]).toBeInstanceOf(DocxWorkerEditorChunkLoadError);
});

test('experimentalPrewarm warns once that it has no effect on the worker editor', async () => {
  spyOn(docxPeerlessEditorLoader, 'load').mockResolvedValue(loaded);
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const first = render(<LazyPeerlessDocxEditor experimentalPrewarm documentName="Prewarmed" />);
  await waitFor(() => expect(first.getByText('Prewarmed')).toBeTruthy());
  first.unmount();
  const second = render(<LazyPeerlessDocxEditor experimentalPrewarm documentName="Prewarmed again" />);
  await waitFor(() => expect(second.getByText('Prewarmed again')).toBeTruthy());
  const prewarm = warn.mock.calls.filter(([message]) => String(message).includes('experimentalPrewarm'));
  expect(prewarm).toEqual([['[DocxEditor] experimentalPrewarm has no effect on the worker editor; the worker loads the engine when a document opens.']]);
});

test('a loaded chunk mounts the editor in the first commit of a remount', async () => {
  const load = spyOn(docxPeerlessEditorLoader, 'load').mockResolvedValue(loaded);
  const first = render(<LazyPeerlessDocxEditor documentName="First mount" loadingIndicator={<span>Loading chunk</span>} />);
  await waitFor(() => expect(first.getByText('First mount')).toBeTruthy());
  first.unmount();
  const second = profiled('remount', <LazyPeerlessDocxEditor documentName="Second mount" loadingIndicator={<span>Loading chunk</span>} />);
  expect(second.commits[0]).toBe('Second mount');
  await act(async () => {});
  expect(second.commits.every((text) => text === 'Second mount')).toBe(true);
  expect(second.view.queryByText('Loading chunk')).toBeNull();
  expect(load).toHaveBeenCalledTimes(1);
});

test('a cold chunk commits the editor once its import resolves, without a reveal timer', async () => {
  const pending = chunk();
  const load = spyOn(docxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  spyOn(performance, 'now').mockReturnValue(0);
  spyOn(console, 'error').mockImplementation(() => {});
  holdRevealTimers = true;
  const cold = profiled('cold', <LazyPeerlessDocxEditor documentName="Cold editor" loadingIndicator={<span>Loading chunk</span>} />);
  expect(cold.view.getByText('Loading chunk')).toBeTruthy();
  expect(cold.commits.every((text) => text === 'Loading chunk')).toBe(true);
  const loadingCommits = cold.commits.length;
  pending.resolve(loaded);
  for (let turn = 0; turn < 6; turn++) await new Promise<void>((resolve) => setImmediate(resolve));
  expect(cold.view.queryByText('Cold editor') !== null).toBe(true);
  expect(cold.commits.slice(loadingCommits)).toEqual(['Cold editor']);
  expect(revealDelays).toEqual([]);
  expect(load).toHaveBeenCalledTimes(1);
});

test('a replacement loader bypasses the cached chunk, retries its failure on remount, then reuses its success', async () => {
  const cached = spyOn(docxPeerlessEditorLoader, 'load').mockResolvedValue(loaded);
  const first = render(<LazyPeerlessDocxEditor documentName="Cached editor" />);
  await waitFor(() => expect(first.getByText('Cached editor')).toBeTruthy());
  first.unmount();

  const cause = new TypeError('Replacement chunk failed');
  const replacement = mock((): Promise<EditorModule> => Promise.resolve(loaded)).mockRejectedValueOnce(cause);
  docxPeerlessEditorLoader.load = replacement;
  const onError = mock((_error: Error) => {});
  spyOn(console, 'error').mockImplementation(() => {});
  const props = { documentName: 'Replacement editor', loadingIndicator: <span>Loading chunk</span>, onError };
  const failed = render(<StrictMode><LazyPeerlessDocxEditor {...props} /></StrictMode>);
  expect(failed.getByText('Loading chunk')).toBeTruthy();
  await waitFor(() => expect(failed.getByRole('alert')).toBeTruthy());
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]![0]).toBeInstanceOf(DocxWorkerEditorChunkLoadError);
  expect(onError.mock.calls[0]![0]).toMatchObject({ code: 'chunk-load-failed', cause });
  failed.unmount();

  const retried = render(<LazyPeerlessDocxEditor {...props} />);
  await waitFor(() => expect(retried.getByText('Replacement editor')).toBeTruthy());
  retried.unmount();
  const remounted = profiled('cached', <LazyPeerlessDocxEditor {...props} />);
  expect(remounted.commits[0]).toBe('Replacement editor');
  expect(remounted.view.queryByText('Loading chunk')).toBeNull();
  expect(replacement).toHaveBeenCalledTimes(2);
  expect(cached).toHaveBeenCalledTimes(1);
  expect(onError).toHaveBeenCalledTimes(1);
});

test('the chunk import starts while a host Suspense reveal is still throttled', async () => {
  const pending = chunk();
  const load = spyOn(docxPeerlessEditorLoader, 'load').mockReturnValue(pending.promise);
  spyOn(performance, 'now').mockReturnValue(0);
  spyOn(console, 'error').mockImplementation(() => {});
  let reveal!: (module: { default: ComponentType }) => void;
  const Host = lazy(() => new Promise<{ default: ComponentType }>((resolve) => { reveal = resolve; }));
  holdRevealTimers = true;
  const view = render(<Suspense fallback={<span>Opening host</span>}><Host /></Suspense>);
  expect(view.getByText('Opening host')).toBeTruthy();
  reveal({ default: () => <LazyPeerlessDocxEditor documentName="Hosted editor" loadingIndicator={<span>Loading chunk</span>} /> });
  for (let turn = 0; turn < 6; turn++) await new Promise<void>((resolve) => setImmediate(resolve));
  expect(revealDelays.length > 0).toBe(true);
  expect(view.queryByText('Opening host') !== null).toBe(true);
  expect(load).toHaveBeenCalledTimes(1);
});
