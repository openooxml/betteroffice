import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { createRef, forwardRef, StrictMode, useImperativeHandle } from 'react';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import { docxPeerlessEditorLoader, DocxWorkerEditorChunkLoadError, LazyPeerlessDocxEditor } from './PeerlessDocxEditorLoader';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');

afterEach(() => { cleanup(); mock.restore(); });
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
