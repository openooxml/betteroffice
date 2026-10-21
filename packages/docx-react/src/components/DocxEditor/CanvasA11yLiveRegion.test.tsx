import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { RefObject } from 'react';
import type { YrsSelectionContext } from '@betteroffice/docx/yrs';
import { CanvasA11yLiveRegion } from './CanvasA11yLiveRegion';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function selection(hasSelection: boolean) {
  return { hasSelection, inInsertion: false, inDeletion: false, paragraphProperties: {} } as unknown as YrsSelectionContext;
}

test('announces the first selection made after the region mounts', () => {
  let context = selection(false);
  const notifyRef: RefObject<(() => void) | null> = { current: null };
  const { container } = render(
    <CanvasA11yLiveRegion active getSelectionContext={() => context} notifyRef={notifyRef} />
  );
  const region = container.querySelector('.oox-canvas-live-region')!;
  expect(region.textContent).toBe('');

  context = selection(true);
  act(() => notifyRef.current?.());
  expect(region.textContent).toContain('Selected');
});
