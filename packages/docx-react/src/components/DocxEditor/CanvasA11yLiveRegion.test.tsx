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

function selection(hasSelection: boolean, inInsertion = false) {
  return { hasSelection, inInsertion, inDeletion: false, paragraphProperties: {} } as unknown as YrsSelectionContext;
}

function mount(initial: YrsSelectionContext) {
  let context = initial;
  let reads = 0;
  const notifyRef: RefObject<(() => void) | null> = { current: null };
  const getSelectionContext = () => {
    reads += 1;
    return context;
  };
  const { container, rerender } = render(
    <CanvasA11yLiveRegion active getSelectionContext={getSelectionContext} notifyRef={notifyRef} />
  );
  rerender(<CanvasA11yLiveRegion active getSelectionContext={() => getSelectionContext()} notifyRef={notifyRef} />);
  return {
    region: container.querySelector('.oox-canvas-live-region')!,
    reads: () => reads,
    select(next: YrsSelectionContext) {
      context = next;
      act(() => notifyRef.current?.());
    },
  };
}

test('announces the first selection made after the region mounts', () => {
  const live = mount(selection(false));
  expect(live.reads()).toBe(0);
  expect(live.region.textContent).toBe('');

  live.select(selection(true));
  expect(live.reads()).toBe(1);
  expect(live.region.textContent).toContain('Selected');
});

test('takes the first caret placement as the baseline', () => {
  const live = mount(selection(false));
  live.select(selection(false, true));
  expect(live.region.textContent).toBe('');

  live.select(selection(true, true));
  expect(live.region.textContent).toContain('Selected');
});
