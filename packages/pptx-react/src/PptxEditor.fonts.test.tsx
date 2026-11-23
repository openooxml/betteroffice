import { deck } from './worker/peerlessTestHarness';
import { afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { act, render, waitFor } from '@testing-library/react';
import * as bundledFonts from '@betteroffice/fonts';
import * as pptxWorker from '@betteroffice/pptx/worker';
import type { BundledFontFace, PptxBundledFont } from './index';
import { PptxEditor, pptxBundledFontsLoader, pptxPeerlessEditorLoader, PptxWorkerEditorChunkLoadError, type PptxAsyncEditorProps } from './PptxEditor';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import { peerlessSessionOpener } from './worker/usePeerlessPresentation';
import { FONT_FALLBACK_MESSAGE } from './fontFaces';
import { MOUNTED } from './__fixtures__/mounted';

const face: BundledFontFace = { family: 'Liberation Sans', weight: 400, style: 'normal',
  file: 'LiberationSans-Regular.ttf.zst', byteLength: 1 };
const baseUrl = 'https://example.test/fonts/';
let fontBytes: Uint8Array;
const decoder = mock(async () => ({ font: new ArrayBuffer(0) }));
const release = mock(() => {});
const take = mock(() => ({} as Worker));
let prepare: ReturnType<typeof spyOn<typeof pptxWorker, 'preparePresentationFontDecoder'>>;
let load: ReturnType<typeof spyOn<typeof bundledFonts, 'loadBundledFontBytes'>>;
let opener: ReturnType<typeof spyOn<typeof peerlessSessionOpener, 'open'>>;
const restorers: (() => void)[] = [];

beforeAll(async () => {
  fontBytes = new Uint8Array(await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
});
beforeEach(() => {
  decoder.mockClear(); release.mockClear(); take.mockClear();
  prepare = spyOn(pptxWorker, 'preparePresentationFontDecoder').mockReturnValue({ decoder, release, take });
  load = spyOn(bundledFonts, 'loadBundledFontBytes').mockImplementation(async () => fontBytes.slice().buffer as ArrayBuffer);
  opener = spyOn(peerlessSessionOpener, 'open');
});
afterEach(() => {
  for (const restore of restorers.splice(0).reverse()) restore();
  opener.mockRestore(); load.mockRestore(); prepare.mockRestore();
});

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: Error) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

async function mount(fonts: PptxAsyncEditorProps['fonts'], options: Partial<PptxAsyncEditorProps> = {}) {
  const file = await deck();
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const props = { file, fonts, onReady: ready, defaultFonts: false, ...options };
  const view = render(<PptxEditor {...props} />);
  return { view, props, ready, file };
}

test('bundled descriptors wait for the prepared decoder and open with byte faces', async () => {
  const pending = deferred<ArrayBuffer>();
  load.mockReturnValue(pending.promise);
  const boldItalic = { ...face, weight: 700 as const, style: 'italic' as const };
  const { view, file, ready } = await mount([{ face: boldItalic, baseUrl }], { loadingIndicator: 'Loading fonts' });
  await waitFor(() => expect(load).toHaveBeenCalledTimes(1));
  expect(prepare).toHaveBeenCalledWith(file, { defaultFonts: false, activeSlideIndex: 0 });
  expect(load).toHaveBeenCalledWith(boldItalic, { baseUrl }, { decoder });
  expect(load.mock.calls.every((call) => call[2]?.decoder === decoder && !('retain' in call[2]!))).toBe(true);
  expect(opener).not.toHaveBeenCalled();
  expect(ready).not.toHaveBeenCalled();
  expect(view.getByRole('status').textContent).toBe('Loading fonts');
  await act(async () => { pending.resolve(fontBytes.slice().buffer as ArrayBuffer); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(opener).toHaveBeenCalledWith(file, expect.objectContaining({ fonts: [
    { family: face.family, bold: true, italic: true, bytes: fontBytes },
  ] }));
  expect(take).not.toHaveBeenCalled();
  expect(release).toHaveBeenCalledTimes(1);
});

test('byte faces pass through unchanged without preparing a font decoder', async () => {
  const fonts = [{ family: face.family, bytes: fontBytes }];
  const { ready } = await mount(fonts);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(opener.mock.calls[0]![1]!.fonts).toBe(fonts);
  expect(opener.mock.calls[0]![1]!.fonts![0]).toBe(fonts[0]);
  expect(prepare).not.toHaveBeenCalled();
  expect(load).not.toHaveBeenCalled();
});

test('equivalent inline descriptors stay stable during loading and after opening', async () => {
  const pending = deferred<ArrayBuffer>();
  load.mockReturnValue(pending.promise);
  const { view, props, ready } = await mount([{ face, baseUrl }]);
  await waitFor(() => expect(load).toHaveBeenCalledTimes(1));
  const ui = () => <PptxEditor {...props} fonts={[{ face: { ...face }, baseUrl: new URL(baseUrl) }]} />;
  await act(async () => { view.rerender(ui()); });
  expect(prepare).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledTimes(1);
  await act(async () => { pending.resolve(fontBytes.slice().buffer as ArrayBuffer); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  await act(async () => { view.rerender(ui()); });
  expect(ready).toHaveBeenCalledTimes(1);
  expect(opener).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledTimes(1);
  expect(prepare).toHaveBeenCalledTimes(1);
});

test('mixed byte and bundled faces preserve both in input order', async () => {
  const byteFace = { family: 'Host Sans', bytes: fontBytes };
  const { ready } = await mount([byteFace, { face, baseUrl }]);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  const fonts = opener.mock.calls[0]![1]!.fonts!;
  expect(fonts).toEqual([byteFace, { family: face.family, bold: false, italic: false, bytes: fontBytes }]);
  expect(fonts[0]).toBe(byteFace);
  expect(load).toHaveBeenCalledWith(face, { baseUrl }, { decoder });
});

test('the prepared decoder gets the default fonts and opening slide the editor will use', async () => {
  const { ready, file } = await mount([{ face, baseUrl }], { defaultFonts: undefined, initialSlide: 3 });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(prepare).toHaveBeenCalledWith(file, { defaultFonts: undefined, activeSlideIndex: 2 });
});

test('metric-compatible descriptors open as the font they substitute', async () => {
  const arial = { ...face, metricCompatWith: 'Arial', weight: 700 as const };
  const { ready } = await mount([{ face: arial, baseUrl }]);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(opener.mock.calls[0]![1]!.fonts).toEqual([{ family: 'Arial', bold: true, italic: false, bytes: fontBytes }]);
});

test('failed descriptors warn once and the presentation opens with the remaining faces', async () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  const onError = mock((_error: Error) => {});
  load.mockRejectedValueOnce(new Error('Font download failed')).mockRejectedValueOnce(new Error('Another font failed'));
  const byteFace = { family: face.family, bytes: fontBytes };
  const { ready } = await mount([{ face, baseUrl }, { face: { ...face, file: 'missing.ttf.zst' } }, byteFace], { onError });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(warn).toHaveBeenCalledTimes(1);
  expect(warn).toHaveBeenCalledWith(FONT_FALLBACK_MESSAGE);
  expect(onError).not.toHaveBeenCalled();
  expect(opener.mock.calls[0]![1]!.fonts).toEqual([byteFace]);
  expect(load.mock.calls.every((call) => call[2]?.decoder === decoder && !('retain' in call[2]!))).toBe(true);
});

for (const outcome of ['resolve', 'reject'] as const) {
  test(`unmount before font ${outcome} releases the prepared worker without opening`, async () => {
    const pending = deferred<ArrayBuffer>();
    load.mockReturnValue(pending.promise);
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    restorers.push(() => warn.mockRestore());
    const { view, ready } = await mount([{ face, baseUrl }]);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(1));
    view.unmount();
    expect(release).toHaveBeenCalledTimes(1);
    await act(async () => {
      if (outcome === 'resolve') pending.resolve(fontBytes.slice().buffer as ArrayBuffer);
      else pending.reject(new Error('Late font failure'));
    });
    expect(opener).not.toHaveBeenCalled();
    expect(ready).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
  });
}

test('changing the file releases the old worker and ignores its pending fonts', async () => {
  const first = deferred<ArrayBuffer>();
  const second = deferred<ArrayBuffer>();
  load.mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);
  const secondRelease = mock(() => {});
  prepare.mockReturnValueOnce({ decoder, release, take }).mockReturnValueOnce({ decoder, release: secondRelease, take });
  const { view, props, ready, file } = await mount([{ face, baseUrl }]);
  await waitFor(() => expect(load).toHaveBeenCalledTimes(1));
  const nextFile = file.slice();
  await act(async () => { view.rerender(<PptxEditor {...props} file={nextFile} />); });
  await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
  expect(release).toHaveBeenCalledTimes(1);
  expect(prepare).toHaveBeenNthCalledWith(2, nextFile, { defaultFonts: false, activeSlideIndex: 0 });
  await act(async () => { first.resolve(fontBytes.slice().buffer as ArrayBuffer); });
  expect(opener).not.toHaveBeenCalled();
  await act(async () => { second.resolve(fontBytes.slice().buffer as ArrayBuffer); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(opener.mock.calls[0]![0]).toBe(nextFile);
  expect(secondRelease).toHaveBeenCalledTimes(1);
});

test('two editors on one file each release their own handle, so unmounting one keeps the other loading', async () => {
  const pending = deferred<ArrayBuffer>();
  load.mockReturnValue(pending.promise);
  const firstRelease = mock(() => {});
  const secondRelease = mock(() => {});
  prepare.mockReturnValueOnce({ decoder, release: firstRelease, take })
    .mockReturnValueOnce({ decoder, release: secondRelease, take });
  const file = await deck();
  const firstReady = mock((_api: PptxAsyncEditorApi) => {});
  const secondReady = mock((_api: PptxAsyncEditorApi) => {});
  const first = render(<PptxEditor file={file} fonts={[{ face, baseUrl }]} defaultFonts={false} onReady={firstReady} />);
  const second = render(<PptxEditor file={file} fonts={[{ face, baseUrl }]} defaultFonts={false} onReady={secondReady} />);
  await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
  first.unmount();
  expect(firstRelease).toHaveBeenCalledTimes(1);
  expect(secondRelease).not.toHaveBeenCalled();
  await act(async () => { pending.resolve(fontBytes.slice().buffer as ArrayBuffer); });
  await waitFor(() => expect(secondReady).toHaveBeenCalledTimes(1), MOUNTED);
  expect(opener.mock.calls[opener.mock.calls.length - 1]![1]!.fonts).toEqual([{ family: face.family, bold: false, italic: false, bytes: fontBytes }]);
  expect(firstReady).not.toHaveBeenCalled();
  second.unmount();
});

test('an in-thread editor releases the prepared worker after decoding', async () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  const { ready } = await mount([{ face, baseUrl }], { experimentalWorkerOpen: false });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(load).toHaveBeenCalledWith(face, { baseUrl }, { decoder });
  expect(release).toHaveBeenCalledTimes(1);
  expect(opener).not.toHaveBeenCalled();
});

test('a failed session open releases the prepared font worker', async () => {
  opener.mockImplementation(() => { throw new Error('Open failed'); });
  const onError = mock((_error: Error) => {});
  const { view, ready } = await mount([{ face, baseUrl }], { onError });
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(view.queryByRole('alert')).not.toBeNull();
  expect(release).toHaveBeenCalledTimes(1);
  expect(ready).not.toHaveBeenCalled();
});

test('a prepared worker is released when the editor chunk never opens a session', async () => {
  const chunk = deferred<Awaited<ReturnType<typeof pptxPeerlessEditorLoader.load>>>();
  const loader = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(chunk.promise);
  restorers.push(() => loader.mockRestore());
  const { view } = await mount([{ face, baseUrl }]);
  await waitFor(() => expect(loader).toHaveBeenCalledTimes(1));
  expect(opener).not.toHaveBeenCalled();
  view.unmount();
  expect(release).toHaveBeenCalledTimes(1);
});

test('descriptors without a file do not prepare a worker or load fonts', async () => {
  const { view, ready } = await mount([{ face, baseUrl }], { file: undefined });
  await act(async () => {});
  expect(prepare).not.toHaveBeenCalled();
  expect(load).not.toHaveBeenCalled();
  expect(ready).not.toHaveBeenCalled();
  view.unmount();
});

const changedDescriptors: { name: string; font: PptxBundledFont }[] = [
  { name: 'asset', font: { face: { ...face, file: 'Another-Regular.ttf.zst' }, baseUrl } },
  { name: 'family', font: { face: { ...face, family: 'Another Sans' }, baseUrl } },
  { name: 'weight', font: { face: { ...face, weight: 700 }, baseUrl } },
  { name: 'style', font: { face: { ...face, style: 'italic' }, baseUrl } },
  { name: 'base URL', font: { face, baseUrl: 'https://example.test/other-fonts/' } },
  { name: 'substituted font', font: { face: { ...face, metricCompatWith: 'Arial' }, baseUrl } },
];
for (const { name, font } of changedDescriptors) {
  test(`changing a descriptor's ${name} resolves new faces and releases the old worker`, async () => {
    load.mockImplementation(() => new Promise<ArrayBuffer>(() => {}));
    const { view, props } = await mount([{ face, baseUrl }]);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(1));
    await act(async () => { view.rerender(<PptxEditor {...props} fonts={[font]} />); });
    await waitFor(() => expect(load).toHaveBeenCalledTimes(2));
    expect(load).toHaveBeenLastCalledWith(font.face, { baseUrl: font.baseUrl }, { decoder });
    expect(prepare).toHaveBeenCalledTimes(2);
    expect(release).toHaveBeenCalledTimes(1);
    expect(opener).not.toHaveBeenCalled();
  });
}

test('decoder preparation failures warn and preserve host byte faces', async () => {
  prepare.mockImplementation(() => { throw new Error('Worker unavailable'); });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  const byteFace = { family: face.family, bytes: fontBytes };
  const { ready } = await mount([byteFace, { face, baseUrl }]);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1), MOUNTED);
  expect(warn).toHaveBeenCalledWith(FONT_FALLBACK_MESSAGE);
  expect(load).not.toHaveBeenCalled();
  expect(opener.mock.calls[0]![1]!.fonts).toEqual([byteFace]);
});

test('a failed descriptor chunk load reports through onError without crashing the host', async () => {
  const failed = spyOn(pptxBundledFontsLoader, 'load').mockRejectedValue(new Error('chunk unavailable'));
  restorers.push(() => failed.mockRestore());
  const onError = mock((_error: Error) => {});
  const view = render(<div data-testid="host"><PptxEditor file={await deck()} fonts={[{ face, baseUrl }]} onError={onError} /></div>);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0]![0]).toBeInstanceOf(PptxWorkerEditorChunkLoadError);
  expect(view.getByTestId('host')).toBeDefined();
  expect(prepare).not.toHaveBeenCalled();
  view.unmount();
});
