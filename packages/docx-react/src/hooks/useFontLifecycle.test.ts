import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { useEffect } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, configure, renderHook } = await import('@testing-library/react');
const { registerDocumentFaces } = await import('@betteroffice/docx/utils');
const { useFontLifecycle, useFontLoadScope } = await import('./useFontLifecycle');

afterEach(() => {
  configure({ reactStrictMode: false });
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('under StrictMode each editor hears its own font loads and not the other one\'s', async () => {
  configure({ reactStrictMode: true });
  const heard = { first: 0, second: 0 };
  const editor = (name: keyof typeof heard) =>
    renderHook(() => {
      const scope = useFontLoadScope();
      useFontLifecycle(undefined, () => (heard[name] += 1), undefined, scope);
      return scope;
    });
  const first = editor('first');
  const second = editor('second');
  expect(first.result.current.disposed).toBe(false);
  await act(() =>
    registerDocumentFaces([{ family: 'StrictMode Face', data: new ArrayBuffer(4) }], first.result.current)
  );
  expect(heard).toEqual({ first: 1, second: 0 });
  first.unmount();
  expect(first.result.current.disposed).toBe(true);
  second.unmount();
});

test('under StrictMode the scope is live again when the effects after it replay', () => {
  configure({ reactStrictMode: true });
  const seen: boolean[] = [];
  const { unmount } = renderHook(() => {
    const scope = useFontLoadScope();
    useEffect(() => void seen.push(scope.disposed), [scope]);
  });
  expect(seen).toEqual([false, false]);
  unmount();
});
