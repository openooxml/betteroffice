import { open, story, type } from './peerlessTestHarness';
import { expect, mock, spyOn, test } from 'bun:test';
import { act, render, waitFor } from '@testing-library/react';
import { PptxEditor, type PptxAsyncEditorProps } from '../PptxEditor';
import { peerlessSessionOpener } from './usePeerlessPresentation';

for (const snapshot of [false, true]) {
  test(`document changes report committed versions with snapshot=${snapshot}`, async () => {
    const changed = mock((_change: { version: string }) => {});
    const snapshots = mock(() => {});
    const value = await open({ onDocumentChange: changed, onChange: snapshot ? snapshots : undefined });
    expect(changed).not.toHaveBeenCalled();
    const target = await story(value.api);
    await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
    const before = await value.api.version();
    await act(async () => { type(value.view.getByLabelText('Slide text'), 'first'); await value.api.flushPendingInput(); });
    await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
    expect(changed.mock.calls[0]).toEqual([{ version: await value.api.version() }]);
    expect(await value.api.version()).not.toBe(before);
    expect(snapshots).toHaveBeenCalledTimes(snapshot ? 1 : 0);
    expect(value.client.state.version).toBe(await value.api.version());
  });
}

test('throwing document listener permits the next edit', async () => {
  const diagnostic = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const changed = mock(() => { throw new Error('listener'); });
    const value = await open({ onDocumentChange: changed });
    const target = await story(value.api);
    await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
    for (let count = 1; count <= 2; count++) {
      await act(async () => { type(value.view.getByLabelText('Slide text'), 'next'); await value.api.flushPendingInput(); });
      await waitFor(() => expect(changed).toHaveBeenCalledTimes(count));
    }
    expect(diagnostic.mock.calls.filter((call) => call[0] === 'documentChange listener threw:')).toHaveLength(2);
  } finally { diagnostic.mockRestore(); }
});

for (const [onSave, downloadOnSave, downloads] of [[true, undefined, 0], [false, undefined, 1], [true, true, 1], [false, false, 0]] as const) {
  test(`save callback=${onSave} download=${downloadOnSave}`, async () => {
    const saved = mock(() => {});
    const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
    try {
      const value = await open({ onSave: onSave ? saved : undefined, downloadOnSave });
      await act(async () => { expect(await value.api.commands!.execute('save', null)).toMatchObject({ ok: true }); });
      expect(saved).toHaveBeenCalledTimes(onSave ? 1 : 0);
      expect(click).toHaveBeenCalledTimes(downloads);
    } finally { click.mockRestore(); }
  });
}

for (const [initialZoom, expected] of [[1.5, 1.5], [20, 4], [0.1, 0.25], [NaN, 'fit'], [0, 'fit'], [undefined, 'fit']] as const) {
  test(`initial zoom ${initialZoom} becomes ${expected}`, async () => {
    const value = await open({ initialZoom });
    expect(value.api.commands!.getState('zoom').value).toBe(expected);
  });
}

test('style overrides the peerless root style', async () => {
  const value = await open({ style: { position: 'absolute', backgroundColor: 'red' } });
  const root = value.view.getByTestId('pptx-peerless-editor');
  expect(root.style.position).toBe('absolute');
  expect(root.style.backgroundColor).toBe('red');
});

test('ArrayBuffer opens once and preserves the view identity', async () => {
  const initial = await open();
  initial.view.unmount();
  const file = initial.bytes.slice().buffer as ArrayBuffer;
  const ready = mock(() => {});
  const opened = spyOn(peerlessSessionOpener, 'open');
  opened.mockClear();
  try {
    const props: PptxAsyncEditorProps = { fonts: [], file, onReady: ready };
    const view = render(<PptxEditor {...props} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(opened.mock.calls[0]![0]).toBeInstanceOf(Uint8Array);
    expect((opened.mock.calls[0]![0] as Uint8Array).buffer).toBe(file);
    view.rerender(<PptxEditor {...props} />);
    expect(opened).toHaveBeenCalledTimes(1);
  } finally { opened.mockRestore(); }
});

test('custom loading indicator replaces preparing and disappears when ready', async () => {
  const initial = await open();
  initial.view.unmount();
  const opened = spyOn(peerlessSessionOpener, 'open');
  const original = opened.getMockImplementation()!;
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  opened.mockImplementation((...args) => {
    const owner = original(...args);
    return { ...owner, ready: gate.then(() => owner.ready) };
  });
  try {
    const ready = mock(() => {});
    const view = render(<PptxEditor fonts={[]} file={initial.bytes} loadingIndicator={<span>Custom opening</span>} onReady={ready} />);
    await waitFor(() => expect(view.getByText('Custom opening')).toBeTruthy());
    await act(async () => { release(); });
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(view.queryByText('Custom opening')).toBeNull();
  } finally { release(); opened.mockRestore(); }
});
