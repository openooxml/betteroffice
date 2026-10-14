import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DocxCommandResult } from '../../../commands/types';
import type { PagedEditorRef } from '../PagedEditor';
import { useFileIO } from './useFileIO';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

const originals = { FileReader: globalThis.FileReader, Image: globalThis.Image };
afterEach(() => {
  cleanup();
  globalThis.FileReader = originals.FileReader;
  globalThis.Image = originals.Image;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a picture decoded after another picker opened goes to the picker it was chosen in', async () => {
  const reads: (() => void)[] = [];
  class SlowReader {
    result: string | null = null;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    readAsDataURL() {
      reads.push(() => {
        this.result = 'data:image/png;base64,AAAA';
        this.onload?.();
      });
    }
  }
  class LoadedImage {
    naturalWidth = 40;
    naturalHeight = 20;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }
  globalThis.FileReader = SlowReader as never;
  globalThis.Image = LoadedImage as never;

  const inserted: string[] = [];
  const insert =
    (name: string) =>
    async (image: Readonly<Record<string, unknown>>): Promise<DocxCommandResult> => {
      inserted.push(`${name}:${image.alt}`);
      return { ok: true, status: 'executed' };
    };
  const { result } = renderHook(() =>
    useFileIO({
      pagedEditorRef: { current: null as PagedEditorRef | null },
      resolveImage: () => null,
      comments: [],
      documentName: 'doc',
      onSave: undefined,
      onOpen: undefined,
      onError: undefined,
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    })
  );
  const choose = (name: string) =>
    result.current.handleImageFileChange({
      target: { files: [new File(['x'], name)], value: name },
    } as never);

  act(() => result.current.handleInsertImageClick(insert('first')));
  choose('a.png');
  act(() => result.current.handleInsertImageClick(insert('second')));
  choose('b.png');
  reads[1]();
  reads[0]();
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
  expect(inserted.sort()).toEqual(['first:a.png', 'second:b.png']);
});
