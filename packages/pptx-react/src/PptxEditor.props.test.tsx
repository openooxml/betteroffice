import { open } from './worker/peerlessTestHarness';
import { expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { act, render, waitFor } from '@testing-library/react';
import { PptxEditor } from './PptxEditor';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';

test('in-thread async props report the handle version and control zoom, style and saving', async () => {
  const initial = await open();
  initial.view.unmount();
  const changed = mock((_change: { version: string }) => {});
  const saved = mock(() => {});
  const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
  const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }];
  let api: PptxAsyncEditorApi | undefined;
  try {
    const view = render(<PptxEditor file={initial.bytes} fonts={fonts} experimentalWorkerOpen={false}
      initialZoom={1.5} style={{ backgroundColor: 'red' }} onDocumentChange={changed}
      onSave={saved} downloadOnSave onReady={(value) => { api = value; }} />);
    await waitFor(() => expect(api).toBeDefined());
    await waitFor(() => expect(api!.commands!.getState('save').enabled).toBe(true));
    expect(changed).not.toHaveBeenCalled();
    expect(api!.commands!.getState('zoom').value).toBe(1.5);
    expect((view.container.firstElementChild as HTMLElement).style.backgroundColor).toBe('red');
    await act(async () => { expect(await api!.commands!.execute('insertSlide', { layoutPartPath: null })).toMatchObject({ ok: true }); });
    await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
    expect(changed.mock.calls[0]![0].version).toBe(await api!.version());
    await act(async () => { expect(await api!.commands!.execute('save', null)).toMatchObject({ ok: true }); });
    expect(saved).toHaveBeenCalledTimes(1);
    expect(click).toHaveBeenCalledTimes(1);
  } finally { click.mockRestore(); }
});
