import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { PeerlessPresentation } from '../worker/usePeerlessPresentation';
import { pptxCommandController } from './createPptxCommandStore';
import { usePeerlessCommandBinding } from './usePeerlessCommandBinding';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('scoped plain-text paste starts the clipboard read synchronously', async () => {
  const selection = { kind: 'text', slideId: 'slide', shapeId: 'shape', storyId: 'story', anchor: 0, focus: 3 };
  const view = {
    selection, composition: null, pending: [], waitingInputCount: 0,
    confirmed: { frame: { slideId: 'slide' } },
    editState: { readOnly: false, canUndo: false, canRedo: false,
      toolbar: { target: selection, formatting: null, alignment: null, table: null } },
  };
  const run = {
    current: true, require() {},
    client: { state: { version: 'v1', slides: [{ id: 'slide', layoutPartPath: null }] } },
    store: {
      getSnapshot: () => view,
      subscribeEditState: () => () => {},
      subscribePresentation: () => () => {},
    },
  } as unknown as PeerlessPresentation;
  let reject!: (error: unknown) => void;
  const clipboardRead = new Promise<string>((_resolve, fail) => { reject = fail; });
  const readText = mock(() => clipboardRead);
  const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText } });
  try {
    const hook = renderHook(() => usePeerlessCommandBinding(run, {
      tool: 'select', setTool() {}, openPicture() {}, save: async () => {}, focus() {},
      zoom: 1, setZoom() {}, review: { open: false, selectedId: null, diff: false, presenting: false },
      proposals: [], setReview() {}, exportPng: async () => {},
    }));
    let gesture = true;
    const scoped = pptxCommandController(hook.result.current)!.scoped({
      gesture: () => gesture, deny: () => null, subscribe: () => () => {},
    });
    expect(scoped.getState('pasteAsPlainText').enabled).toBe(true);
    const pending = scoped.execute('pasteAsPlainText', null);
    expect(readText).toHaveBeenCalledTimes(1);
    gesture = false;
    reject(new DOMException('Unavailable', 'NotAllowedError'));
    expect(await pending).toMatchObject({ ok: false, failure: { code: 'clipboard-unavailable' } });
    expect(await scoped.execute('pasteAsPlainText', null)).toMatchObject({
      ok: false, failure: { code: 'unsupported-policy' },
    });
    expect(readText).toHaveBeenCalledTimes(1);
  } finally {
    if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
    else Reflect.deleteProperty(navigator, 'clipboard');
  }
});
