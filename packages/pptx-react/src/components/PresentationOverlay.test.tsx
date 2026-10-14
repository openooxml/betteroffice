import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, spyOn } from 'bun:test';
import type { PresentationHandle, SlideDisplayList } from '@betteroffice/pptx';
import { PresentationOverlay } from './PresentationOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import(
  '@testing-library/react'
);
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const frame = (assetId: string): SlideDisplayList => ({
  contractVersion: 1,
  width: 320,
  height: 180,
  primitives: [
    {
      kind: 'image',
      name: 'Test image',
      assetId,
      objectId: 1,
      x: 0,
      y: 0,
      w: 320,
      h: 180,
    },
  ],
});

it('only presents the latest slide when an earlier image resolves late', async () => {
  const images = [{}, {}] as CanvasImageSource[];
  let finishFirst: (image: CanvasImageSource) => void = () => {};
  const pending = new Promise<CanvasImageSource>((resolve) => {
    finishFirst = resolve;
  });
  const draws = new Map<HTMLCanvasElement, unknown[][]>();
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    function (this: HTMLCanvasElement) {
      const calls: unknown[][] = [];
      draws.set(this, calls);
      return new Proxy(
        {},
        {
          get: (_target, key) =>
            key === 'drawImage' ? (...args: unknown[]) => calls.push(args) : () => {},
        }
      ) as CanvasRenderingContext2D;
    } as unknown as typeof HTMLCanvasElement.prototype.getContext
  );
  const handle = {
    layoutSlide: (index: number) => frame(String(index)),
  } as PresentationHandle;
  const errors: unknown[] = [];
  try {
    const view = render(
      <PresentationOverlay
        handle={handle}
        slideCount={2}
        startIndex={0}
        resolveImage={(id) => (id === '0' ? pending : images[1])}
        label="Presentation"
        counterLabel={(current, total) => `${current} / ${total}`}
        exitLabel="Exit"
        previousLabel="Previous"
        nextLabel="Next"
        onExit={() => {}}
        onError={(error) => errors.push(error)}
      />
    );
    const displayed = view.container.querySelector('canvas')!;
    fireEvent.click(view.getByRole('button', { name: 'Next' }));
    await waitFor(() => expect(draws.get(displayed)?.length).toBe(1));
    const painted = draws.get(displayed)![0][0] as HTMLCanvasElement;
    expect(draws.get(painted)?.[0][0]).toBe(images[1]);
    await act(async () => {
      finishFirst(images[0]);
      await pending;
    });
    expect(draws.get(displayed)?.length).toBe(1);
    expect(view.getByText('2 / 2')).toBeTruthy();
    expect(errors).toEqual([]);
  } finally {
    context.mockRestore();
  }
});

it('clamps navigation after slides are deleted and restores keyboard focus on exit', async () => {
  const focusTarget = document.createElement('button');
  document.body.append(focusTarget);
  focusTarget.focus();
  const layouts: number[] = [];
  const handle = {
    layoutSlide: (index: number) => {
      layouts.push(index);
      return { ...frame(''), primitives: [] };
    },
  } as unknown as PresentationHandle;
  let exits = 0;
  const props = {
    handle,
    slideCount: 3,
    startIndex: 2,
    resolveImage: () => null,
    label: 'Presentation',
    counterLabel: (current: number, total: number) => `${current} / ${total}`,
    exitLabel: 'Exit',
    previousLabel: 'Previous',
    nextLabel: 'Next',
    onExit: () => {
      exits++;
    },
    onError: () => {},
  };
  const view = render(<PresentationOverlay {...props} />);
  expect(view.getByText('3 / 3')).toBeTruthy();
  view.rerender(<PresentationOverlay {...props} slideCount={1} />);
  expect(view.getByText('1 / 1')).toBeTruthy();
  expect((view.getByRole('button', { name: 'Next' }) as HTMLButtonElement).disabled).toBe(
    true
  );
  const dialog = view.getByRole('dialog');
  fireEvent.keyDown(dialog, { key: 'Escape' });
  expect(exits).toBe(1);
  view.unmount();
  expect(document.activeElement).toBe(focusTarget);
  focusTarget.remove();
});
