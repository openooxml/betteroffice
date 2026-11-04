import { expect, mock, spyOn, test } from 'bun:test';
import { open } from './peerlessTestHarness';

const { act } = await import('@testing-library/react');

async function exporting(png: Blob | null, run: () => Promise<unknown>) {
  const names: string[] = [];
  const encode = spyOn(HTMLCanvasElement.prototype, 'toBlob').mockImplementation((callback) => callback(png));
  const createUrl = spyOn(URL, 'createObjectURL').mockReturnValue('blob:slide');
  const revokeUrl = spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
  const download = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) { names.push(this.download); });
  try {
    let result: unknown;
    await act(async () => { result = await run(); });
    return { names, result };
  } finally { encode.mockRestore(); createUrl.mockRestore(); revokeUrl.mockRestore(); download.mockRestore(); }
}

const png = new Blob([new Uint8Array([137, 80, 78, 71])], { type: 'image/png' });

test('PNG export names the download after the file and the slide', async () => {
  const value = await open({ fileName: 'Quarterly.pptx' });
  await act(async () => { expect(await value.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(3);
  await act(async () => { expect(await value.api.goToSlide(3)).toBe(true); });
  const { names, result } = await exporting(png, () => value.api.commands.execute('exportPng', null));
  expect(result).toMatchObject({ ok: true });
  expect(names).toEqual(['Quarterly-slide-3.png']);
});

test('PNG export without a file name falls back to the presentation name', async () => {
  const value = await open();
  const { names, result } = await exporting(png, () => value.api.commands.execute('exportPng', null));
  expect(result).toMatchObject({ ok: true });
  expect(names).toEqual(['presentation-slide-1.png']);
});

test('a failed PNG export reports the error once and does not download', async () => {
  const error = mock((_: Error) => {});
  const value = await open({ fileName: 'Quarterly.pptx', onError: error });
  const { names, result } = await exporting(null, () => value.api.commands.execute('exportPng', null));
  expect(result).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
  expect(names).toEqual([]);
  expect(error).toHaveBeenCalledTimes(1);
  expect(error.mock.calls[0]![0]).toBeInstanceOf(Error);
  expect(error.mock.calls[0]![0].message).toBe('PNG encoding failed');
});
