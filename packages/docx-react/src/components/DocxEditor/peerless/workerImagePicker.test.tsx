import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { DocxEditorRef } from '../../DocxEditor';
import { useWorkerFileIO } from './useWorkerFileIO';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalImage = globalThis.Image;
afterEach(() => { cleanup(); globalThis.Image = originalImage; });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

test('worker image picker decodes the picked file into the hybrid image payload', async () => {
  globalThis.Image = class {
    naturalWidth = 1224;
    naturalHeight = 800;
    onload: (() => void) | null = null;
    set src(_value: string) { queueMicrotask(() => this.onload?.()); }
  } as never;
  const focus = mock(() => {});
  const insert = mock(async (_image: Readonly<Record<string, unknown>>) => ({ ok: true as const, status: 'executed' as const }));
  const hook = renderHook(() => useWorkerFileIO({}, { focus } as unknown as DocxEditorRef, async () => {}, () => ({})));
  const input = document.createElement('input');
  input.type = 'file';
  hook.result.current.imageInput.current = input;
  const click = mock(() => {});
  input.click = click;
  hook.result.current.pickImage(insert);
  expect(click).toHaveBeenCalledTimes(1);
  const event = { currentTarget: { files: [new File(['image'], 'pixel.png', { type: 'image/png' })], value: 'chosen' } };
  hook.result.current.imageChange(event as never);
  expect(event.currentTarget.value).toBe('');
  await waitFor(() => expect(insert).toHaveBeenCalledWith({ src: 'data:image/png;base64,aW1hZ2U=', alt: 'pixel.png',
    width: 612, height: 400, wrapType: 'inline', displayMode: 'inline', distTop: 0, distBottom: 0, distLeft: 0, distRight: 0 }));
  expect(focus).toHaveBeenCalledTimes(1);
});
