import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, mock, spyOn } from 'bun:test';
import { useRef } from 'react';
import { pinchZoomScale, usePptxPinchZoom } from './usePptxPinchZoom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

it('clamps trackpad pinch scaling to the editor zoom range', () => {
  expect(pinchZoomScale(1, -10_000, WheelEvent.DOM_DELTA_PIXEL, 600)).toBeLessThanOrEqual(5);
  expect(pinchZoomScale(1, 10_000, WheelEvent.DOM_DELTA_PIXEL, 600)).toBeGreaterThanOrEqual(0.1);
  expect(pinchZoomScale(1, -10, WheelEvent.DOM_DELTA_PIXEL, 600)).toBeGreaterThan(1);
  expect(pinchZoomScale(1, 10, WheelEvent.DOM_DELTA_PIXEL, 600)).toBeLessThan(1);
});

it('zooms only for trackpad pinch and keeps its canvas anchor stationary', async () => {
  let renderedScale = 1;
  let requestedScale = 1;
  const setScale = mock((scale: number) => {
    requestedScale = scale;
  });
  let resize = () => {};
  const resizeObserver = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class {
    constructor(callback: ResizeObserverCallback) {
      resize = () => callback([], this as unknown as ResizeObserver);
    }
    observe() {}
    unobserve() {}
    disconnect() {}
  } as unknown as typeof ResizeObserver;
  const frame = { left: 10, top: 5, width: 280, height: 140 };
  function Harness() {
    const viewport = useRef<HTMLDivElement>(null);
    const content = useRef<HTMLDivElement>(null);
    usePptxPinchZoom(viewport, content, 1, setScale);
    return <div ref={viewport} data-testid="viewport"><div ref={content} data-testid="content" /></div>;
  }
  const view = render(<Harness />);
  await act(async () => {
    await Promise.resolve();
  });
  const viewport = view.getByTestId('viewport');
  const content = view.getByTestId('content');
  Object.defineProperties(viewport, {
    clientWidth: { configurable: true, value: 300 },
    clientHeight: { configurable: true, value: 150 },
  });
  spyOn(viewport, 'getBoundingClientRect').mockImplementation(() => ({
    left: 0,
    top: 0,
    width: 300,
    height: 150,
    right: 300,
    bottom: 150,
    x: 0,
    y: 0,
    toJSON: () => ({}),
  }));
  spyOn(content, 'getBoundingClientRect').mockImplementation(() => ({
    ...frame,
    width: frame.width * renderedScale,
    height: frame.height * renderedScale,
    right: frame.left + frame.width * renderedScale,
    bottom: frame.top + frame.height * renderedScale,
    x: frame.left,
    y: frame.top,
    toJSON: () => ({}),
  }));
  const frames: FrameRequestCallback[] = [];
  const animation = spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
    frames.push(callback);
    return frames.length;
  });
  const wheel = (ctrlKey: boolean) => {
    const event = new Event('wheel', { bubbles: true, cancelable: true }) as WheelEvent;
    Object.defineProperties(event, {
      clientX: { value: 240 },
      clientY: { value: 100 },
      ctrlKey: { value: ctrlKey },
      metaKey: { value: false },
      altKey: { value: false },
      deltaY: { value: -50 },
      deltaMode: { value: WheelEvent.DOM_DELTA_PIXEL },
    });
    return event;
  };
  try {
    const scroll = wheel(false);
    viewport.dispatchEvent(scroll);
    expect(scroll.defaultPrevented).toBe(false);
    expect(setScale).not.toHaveBeenCalled();

    const pinch = wheel(true);
    viewport.dispatchEvent(pinch);
    expect(pinch.defaultPrevented).toBe(true);
    viewport.dispatchEvent(wheel(true));
    expect(setScale).not.toHaveBeenCalled();
    frames.shift()?.(0);
    expect(setScale).toHaveBeenCalledTimes(1);
    expect(viewport.scrollLeft).toBe(0);
    expect(viewport.scrollTop).toBe(0);
    renderedScale = requestedScale;
    resize();
    expect(viewport.scrollLeft).toBeCloseTo(230 * renderedScale - 240);
    expect(viewport.scrollTop).toBeCloseTo(95 * renderedScale - 100);
  } finally {
    animation.mockRestore();
    globalThis.ResizeObserver = resizeObserver;
  }
});
