import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef } from 'react';
import { LayoutSelectionGate } from '../internals/LayoutSelectionGate';
import { useSelectionOverlay } from './useSelectionOverlay';

const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('without geometry a clear overlay does not resolve the selection', () => {
  let resolved = 0;
  const hook = renderHook(() =>
    useSelectionOverlay({
      layout: null,
      containerRef: createRef<HTMLDivElement>(),
      syncCoordinator: new LayoutSelectionGate(),
      displayListQueries: null,
      getYrsDisplaySelection: () => {
        resolved += 1;
        return { anchor: 3, head: 3 };
      },
    })
  );
  act(() => hook.result.current.updateSelectionOverlay());
  expect(resolved).toBe(0);

  // A caret left from before has to be cleared, which needs the selection.
  act(() => hook.result.current.setCaretPosition({ x: 1, y: 2, height: 3, pageIndex: 0 }));
  act(() => hook.result.current.updateSelectionOverlay());
  expect(resolved).toBe(1);
  expect(hook.result.current.caretPosition).toBeNull();
});
