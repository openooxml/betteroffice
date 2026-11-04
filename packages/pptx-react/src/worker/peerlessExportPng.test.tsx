import { expect, mock, spyOn, test } from 'bun:test';
import { open } from './peerlessTestHarness';

const { act } = await import('@testing-library/react');

async function exporting(value: Awaited<ReturnType<typeof open>>, png: Uint8Array | Error, run: () => Promise<unknown>) {
  const names: string[] = [];
  const targets: unknown[] = [];
  const encode = spyOn(value.client, 'renderPng').mockImplementation(async (target) => {
    targets.push(target);
    if (png instanceof Error) throw png;
    return png;
  });
  const createUrl = spyOn(URL, 'createObjectURL').mockReturnValue('blob:slide');
  const revokeUrl = spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
  const download = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) { names.push(this.download); });
  try {
    let result: unknown;
    await act(async () => { result = await run(); });
    return { names, result, targets };
  } finally { encode.mockRestore(); createUrl.mockRestore(); revokeUrl.mockRestore(); download.mockRestore(); }
}

const png = new Uint8Array([137, 80, 78, 71]);

test('PNG export names the download after the file and the slide', async () => {
  const value = await open({ fileName: 'Quarterly.pptx' });
  await act(async () => { expect(await value.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(3);
  await act(async () => { expect(await value.api.goToSlide(3)).toBe(true); });
  const { names, result, targets } = await exporting(value, png, () => value.api.commands.execute('exportPng', null));
  expect(result).toMatchObject({ ok: true });
  expect(names).toEqual(['Quarterly-slide-3.png']);
  expect(targets).toEqual([{ slide: 2 }]);
});

test('PNG export without a file name falls back to the presentation name', async () => {
  const value = await open();
  const { names, result } = await exporting(value, png, () => value.api.commands.execute('exportPng', null));
  expect(result).toMatchObject({ ok: true });
  expect(names).toEqual(['presentation-slide-1.png']);
});

test('a failed PNG export reports the error once and does not download', async () => {
  const error = mock((_: Error) => {});
  const value = await open({ fileName: 'Quarterly.pptx', onError: error });
  const { names, result } = await exporting(value, new Error('PNG encoding failed'), () => value.api.commands.execute('exportPng', null));
  expect(result).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
  expect(names).toEqual([]);
  expect(error).toHaveBeenCalledTimes(1);
  expect(error.mock.calls[0]![0]).toBeInstanceOf(Error);
  expect(error.mock.calls[0]![0].message).toBe('PNG encoding failed');
});
