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

async function mountPinchZoom(initialScale: number | null = 1) {
  let renderedScale = initialScale ?? 1;
  const setScale = mock((_scale: number) => {});
  const disconnect = mock(() => {});
  let resize = () => {};
  const resizeObserver = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class {
    constructor(callback: ResizeObserverCallback) {
      resize = () => callback([], this as unknown as ResizeObserver);
    }
    observe() {}
    unobserve() {}
    disconnect = disconnect;
  } as unknown as typeof ResizeObserver;
  const frame = { left: 10, top: 5, width: 280, height: 140 };
  function Harness({ scale }: { scale: number | null }) {
    const viewport = useRef<HTMLDivElement>(null);
    const content = useRef<HTMLDivElement>(null);
    usePptxPinchZoom(viewport, content, scale, setScale, 'presentation');
    return <div ref={viewport} data-testid="viewport"><div ref={content} data-testid="content" /></div>;
  }
  const view = render(<Harness scale={initialScale} />);
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
  const frames = new Map<number, FrameRequestCallback>();
  let frameId = 0;
  const animation = spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
    frames.set(++frameId, callback);
    return frameId;
  });
  const cancelAnimation = spyOn(window, 'cancelAnimationFrame').mockImplementation((id) => {
    frames.delete(id);
  });
  const wheel = (ctrlKey: boolean, clientX = 240, clientY = 100, timeStamp = 0) => {
    const event = new Event('wheel', { bubbles: true, cancelable: true }) as WheelEvent;
    Object.defineProperties(event, {
      clientX: { value: clientX },
      clientY: { value: clientY },
      ctrlKey: { value: ctrlKey },
      metaKey: { value: false },
      altKey: { value: false },
      deltaY: { value: -50 },
      deltaMode: { value: WheelEvent.DOM_DELTA_PIXEL },
      timeStamp: { value: timeStamp },
    });
    viewport.dispatchEvent(event);
    return event;
  };
  return {
    viewport,
    setScale,
    disconnect,
    wheel,
    rerender: (scale: number | null) => view.rerender(<Harness scale={scale} />),
    flushFrame() {
      for (const [id, callback] of [...frames]) {
        if (frames.delete(id)) callback(0);
      }
    },
    paint(scale: number) {
      renderedScale = scale;
      resize();
    },
    restore() {
      animation.mockRestore();
      cancelAnimation.mockRestore();
      globalThis.ResizeObserver = resizeObserver;
    },
  };
}

it('keeps the latest pinch anchor pending through an earlier canvas paint', async () => {
  const { viewport, setScale, wheel, flushFrame, paint, restore } = await mountPinchZoom();
  try {
    expect(wheel(false).defaultPrevented).toBe(false);
    expect(setScale).not.toHaveBeenCalled();

    expect(wheel(true).defaultPrevented).toBe(true);
    const firstScale = pinchZoomScale(1, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    wheel(true);
    const nextScale = pinchZoomScale(firstScale, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).not.toHaveBeenCalled();
    flushFrame();
    expect(setScale).toHaveBeenCalledTimes(1);
    expect(setScale).toHaveBeenCalledWith(nextScale);
    expect(viewport.scrollLeft).toBe(0);
    expect(viewport.scrollTop).toBe(0);
    paint(firstScale);
    expect(viewport.scrollLeft).toBe(0);
    expect(viewport.scrollTop).toBe(0);
    paint(nextScale);
    expect(viewport.scrollLeft).toBeCloseTo(230 * (nextScale - 1));
    expect(viewport.scrollTop).toBeCloseTo(95 * (nextScale - 1));
  } finally {
    restore();
  }
});

it('ignores pinch before the first fit scale is available', async () => {
  const { viewport, setScale, wheel, rerender, flushFrame, paint, restore } = await mountPinchZoom(null);
  try {
    expect(wheel(true, 100, 50).defaultPrevented).toBe(false);
    flushFrame();
    expect(setScale).not.toHaveBeenCalled();
    expect(viewport.scrollLeft).toBe(0);
    expect(viewport.scrollTop).toBe(0);

    rerender(0.5);
    paint(0.5);
    expect(wheel(true, 100, 50).defaultPrevented).toBe(true);
    flushFrame();
    const nextScale = pinchZoomScale(0.5, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).toHaveBeenCalledTimes(1);
    expect(setScale).toHaveBeenCalledWith(nextScale);
    paint(nextScale);
    expect(viewport.scrollLeft).toBeCloseTo(90 * (nextScale / 0.5 - 1));
    expect(viewport.scrollTop).toBeCloseTo(45 * (nextScale / 0.5 - 1));
  } finally {
    restore();
  }
});

it('drops a pending pinch anchor when the scale changes externally', async () => {
  const { viewport, setScale, disconnect, wheel, rerender, flushFrame, paint, restore } = await mountPinchZoom();
  try {
    expect(wheel(true).defaultPrevented).toBe(true);
    flushFrame();
    const firstScale = pinchZoomScale(1, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).toHaveBeenCalledWith(firstScale);
    const disconnects = disconnect.mock.calls.length;

    rerender(2);
    expect(disconnect).toHaveBeenCalledTimes(disconnects + 1);
    paint(firstScale);
    expect(viewport.scrollLeft).toBe(0);
    expect(viewport.scrollTop).toBe(0);

    paint(2);
    expect(wheel(true, 120, 60).defaultPrevented).toBe(true);
    flushFrame();
    const nextScale = pinchZoomScale(2, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).toHaveBeenLastCalledWith(nextScale);
    paint(nextScale);
    expect(viewport.scrollLeft).toBeCloseTo(110 * (nextScale / 2 - 1));
    expect(viewport.scrollTop).toBeCloseTo(55 * (nextScale / 2 - 1));
  } finally {
    restore();
  }
});

it('preserves pinch steps across a render before the zoom frame commits', async () => {
  const { viewport, setScale, wheel, rerender, flushFrame, paint, restore } = await mountPinchZoom();
  try {
    expect(wheel(true).defaultPrevented).toBe(true);
    rerender(1);
    expect(wheel(true).defaultPrevented).toBe(true);
    expect(setScale).not.toHaveBeenCalled();
    flushFrame();
    const firstScale = pinchZoomScale(1, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    const nextScale = pinchZoomScale(firstScale, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).toHaveBeenCalledTimes(1);
    expect(setScale).toHaveBeenCalledWith(nextScale);
    paint(nextScale);
    expect(viewport.scrollLeft).toBeCloseTo(230 * (nextScale - 1));
    expect(viewport.scrollTop).toBeCloseTo(95 * (nextScale - 1));
  } finally {
    restore();
  }
});

it('starts a fresh anchor after an idle pinch with an unfinished repaint', async () => {
  const { viewport, setScale, disconnect, wheel, rerender, flushFrame, paint, restore } = await mountPinchZoom();
  try {
    expect(wheel(true).defaultPrevented).toBe(true);
    flushFrame();
    const firstScale = pinchZoomScale(1, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).toHaveBeenCalledWith(firstScale);
    const disconnects = disconnect.mock.calls.length;
    rerender(firstScale);
    expect(disconnect).toHaveBeenCalledTimes(disconnects);

    expect(wheel(true, 120, 60, 301).defaultPrevented).toBe(true);
    flushFrame();
    const nextScale = pinchZoomScale(firstScale, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    expect(setScale).toHaveBeenLastCalledWith(nextScale);
    const ratio = nextScale / firstScale;
    paint(ratio);
    expect(viewport.scrollLeft).toBeCloseTo(110 * (ratio - 1));
    expect(viewport.scrollTop).toBeCloseTo(55 * (ratio - 1));
  } finally {
    restore();
  }
});

it('drops a pending pinch anchor when the presentation changes', async () => {
  const setScale = mock(() => {});
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
  function Harness({ presentation }: { presentation: object }) {
    const viewport = useRef<HTMLDivElement>(null);
    const content = useRef<HTMLDivElement>(null);
    usePptxPinchZoom(viewport, content, 1, setScale, presentation);
    return <div ref={viewport} data-testid="viewport"><div ref={content} data-testid="content" /></div>;
  }
  const first = {};
  const view = render(<Harness presentation={first} />);
  await act(async () => {
    await Promise.resolve();
  });
  const viewport = view.getByTestId('viewport');
  const content = view.getByTestId('content');
  Object.defineProperty(viewport, 'clientHeight', { configurable: true, value: 150 });
  let rect = { left: 10, top: 5, width: 280, height: 140 };
  spyOn(content, 'getBoundingClientRect').mockImplementation(() => ({
    ...rect,
    right: rect.left + rect.width,
    bottom: rect.top + rect.height,
    x: rect.left,
    y: rect.top,
    toJSON: () => ({}),
  }));
  const frames: FrameRequestCallback[] = [];
  const animation = spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
    frames.push(callback);
    return frames.length;
  });
  const wheel = () => {
    const event = new Event('wheel', { bubbles: true, cancelable: true }) as WheelEvent;
    Object.defineProperties(event, {
      clientX: { value: 240 },
      clientY: { value: 100 },
      ctrlKey: { value: true },
      metaKey: { value: false },
      altKey: { value: false },
      deltaY: { value: -50 },
      deltaMode: { value: WheelEvent.DOM_DELTA_PIXEL },
    });
    return event;
  };
  try {
    viewport.dispatchEvent(wheel());
    view.rerender(<Harness presentation={{}} />);
    frames.length = 0;
    rect = { left: 20, top: 10, width: 560, height: 280 };
    resize();
    expect(viewport.scrollLeft).toBe(0);
    expect(viewport.scrollTop).toBe(0);

    viewport.dispatchEvent(wheel());
    frames.shift()?.(0);
    const next = pinchZoomScale(1, -50, WheelEvent.DOM_DELTA_PIXEL, 150);
    rect = { ...rect, width: 560 * next, height: 280 * next };
    resize();
    expect(viewport.scrollLeft).toBeCloseTo(20 + 220 * next - 240);
    expect(viewport.scrollTop).toBeCloseTo(10 + 90 * next - 100);
  } finally {
    animation.mockRestore();
    globalThis.ResizeObserver = resizeObserver;
  }
});

it('uses the resized oversized canvas geometry to keep the cursor anchor fixed', async () => {
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
  function Harness() {
    const viewport = useRef<HTMLDivElement>(null);
    const content = useRef<HTMLDivElement>(null);
    usePptxPinchZoom(viewport, content, 1, setScale, 'presentation');
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
  spyOn(content, 'getBoundingClientRect').mockImplementation(() => {
    const width = 280 * renderedScale;
    const height = 140 * renderedScale;
    const left = 100 + (300 - width) / 2 - viewport.scrollLeft;
    const top = 50 + (150 - height) / 2 - viewport.scrollTop;
    return {
      left,
      top,
      width,
      height,
      right: left + width,
      bottom: top + height,
      x: left,
      y: top,
      toJSON: () => ({}),
    };
  });
  const frames: FrameRequestCallback[] = [];
  const animation = spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
    frames.push(callback);
    return frames.length;
  });
  const event = new Event('wheel', { bubbles: true, cancelable: true }) as WheelEvent;
  Object.defineProperties(event, {
    clientX: { value: 340 },
    clientY: { value: 145 },
    ctrlKey: { value: true },
    metaKey: { value: false },
    altKey: { value: false },
    deltaY: { value: -50 },
    deltaMode: { value: WheelEvent.DOM_DELTA_PIXEL },
  });
  try {
    viewport.dispatchEvent(event);
    frames.shift()?.(0);
    renderedScale = requestedScale;
    resize();

    const resized = content.getBoundingClientRect();
    const anchorX = (340 - 110) / 280;
    const anchorY = (145 - 55) / 140;
    expect(resized.left).toBeLessThan(100);
    expect(resized.top).toBeLessThan(50);
    expect(resized.left + anchorX * resized.width).toBeCloseTo(340);
    expect(resized.top + anchorY * resized.height).toBeCloseTo(145);
  } finally {
    animation.mockRestore();
    globalThis.ResizeObserver = resizeObserver;
  }
});
