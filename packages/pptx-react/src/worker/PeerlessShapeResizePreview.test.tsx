import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, mock, spyOn } from 'bun:test';
import type { SlideDisplayList } from '@betteroffice/pptx';
import { useRef } from 'react';
import type { PeerlessViewStore } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { usePeerlessShapeGestures, type PeerlessShapePreview, type PeerlessShapePreviewStore } from './usePeerlessShapeGestures';
import { PeerlessShapeResizePreview, peerlessShapeResizePreviewPaint } from './PeerlessShapeResizePreview';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function request(width: number, gesture = 'gesture'): NonNullable<PeerlessShapePreview['request']> {
  return {
    gesture,
    frame: { slideId: 'slide', slideRevision: 1, coversThroughSeq: 0, documentEpoch: 1, incarnation: 'deck' },
    target: { slideId: 'slide', shapeId: 'shape', token: 'token' },
    rect: { x: 0, y: 0, width, height: 100 },
  };
}

const frame = (width: number): SlideDisplayList => ({
  contractVersion: 1,
  width,
  height: 100,
  primitives: [],
});

it.each(['resolved', 'refused'])('keeps the last painted preview after release when a late request is %s', async (outcome) => {
  const late = deferred<SlideDisplayList>();
  const committedPaint = deferred<void>();
  const previewShapeRect = mock(() => late.promise).mockResolvedValueOnce(frame(80));
  const { frame: tag, target } = request(100);
  const selection = { ...target, kind: 'object', storyId: 'story', canMove: true, canResize: true,
    units: { x: 10000, y: 10000 }, rect: { x: 0, y: 0, width: 1000000, height: 1000000 } };
  const store = {
    getSnapshot: () => ({ selection, editState: { readOnly: false } }),
    subscribePresentation: () => () => {}, subscribeEditState: () => () => {},
    dispatch: async () => ({ throughSeq: 1, outcomes: [{ seq: 1, committed: true }] }),
  } as unknown as PeerlessViewStore;
  const run = {
    current: true, selectionEnabled: true, require: () => {},
    getVisualSnapshot: () => ({ view: { confirmed: { frame: tag, displayList: frame(100) } } }),
    whenPaintedThrough: () => committedPaint.promise,
    client: { previewShapeRect },
    images: { pin: () => ({ ready: Promise.resolve(), resolve: () => null, release: () => {} }) },
  } as unknown as PeerlessPresentation;
  function Stage() {
    const canvas = useRef<HTMLCanvasElement>(null);
    const gestures = usePeerlessShapeGestures(store, { canvas, presentation: run });
    return <div onPointerDown={gestures.onPointerDown} onPointerMove={gestures.onPointerMove} onPointerUp={gestures.onPointerUp}>
      <canvas ref={canvas} data-testid="slide" />
      <button data-pptx-resize="e">Resize</button>
      <PeerlessShapeResizePreview run={run} preview={gestures.preview} scale={1} />
    </div>;
  }
  const drawImage = mock(() => {});
  const context = { setTransform: () => {}, clearRect: () => {}, drawImage } as unknown as CanvasRenderingContext2D;
  const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
  const paint = spyOn(peerlessShapeResizePreviewPaint, 'paint').mockResolvedValue();
  try {
    const view = render(<Stage />);
    const slide = view.getByTestId('slide');
    slide.getBoundingClientRect = () => new DOMRect(0, 0, 100, 100);
    const point = (clientX: number) => ({ pointerId: 1, isPrimary: true, button: 0, clientX, clientY: 50 });
    await act(async () => { fireEvent.pointerDown(view.getByRole('button'), point(100)); });
    fireEvent.pointerMove(slide, point(80));
    await waitFor(() => expect(drawImage).toHaveBeenCalledTimes(1));
    fireEvent.pointerMove(slide, point(60));
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(2));
    fireEvent.pointerMove(slide, point(50));
    fireEvent.pointerUp(slide, point(40));
    await act(async () => {
      if (outcome === 'resolved') late.resolve(frame(60));
      else late.reject(new Error('stale preview'));
    });
    expect(view.queryByTestId('pptx-shape-resize-preview')?.style.width).toBe('80px');
    expect(drawImage).toHaveBeenCalledTimes(1);
    expect(previewShapeRect).toHaveBeenCalledTimes(2);
    await act(async () => { committedPaint.resolve(); });
    expect(view.queryByTestId('pptx-shape-resize-preview')).toBeNull();
  } finally {
    paint.mockRestore();
    getContext.mockRestore();
  }
});

it('backs the resize preview with opaque white', () => {
  const value = { token: 'token', bounds: request(80).rect, request: request(80) };
  const preview = { getSnapshot: () => value, subscribe: () => () => {} };
  const run = { current: true, client: { previewShapeRect: () => deferred<SlideDisplayList>().promise } } as unknown as PeerlessPresentation;
  const view = render(<PeerlessShapeResizePreview run={run} preview={preview} scale={1} />);
  expect(view.getByTestId('pptx-shape-resize-preview').style.backgroundColor).toBe('white');
});

it('paints completed current-gesture previews while coalescing latest work and drops stale gestures', async () => {
  let value: PeerlessShapePreview | null = null;
  const listeners = new Set<() => void>();
  const store: PeerlessShapePreviewStore = {
    getSnapshot: () => value,
    subscribe(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  };
  const set = (next: PeerlessShapePreview | null) => act(() => {
    value = next;
    for (const listener of listeners) listener();
  });
  const pending: ReturnType<typeof deferred<SlideDisplayList>>[] = [];
  const previewShapeRect = mock((_frame: ReturnType<typeof request>['frame'],
    _target: ReturnType<typeof request>['target'], _rect: ReturnType<typeof request>['rect']) => {
    const next = deferred<SlideDisplayList>();
    pending.push(next);
    return next.promise;
  });
  const release = mock(() => {});
  const run = {
    current: true,
    require: () => {},
    client: { previewShapeRect },
    images: { pin: () => ({ ready: Promise.resolve(), resolve: () => null, release }) },
    ensureWholeFonts: async () => {},
  } as unknown as PeerlessPresentation;
  const context = {
    setTransform: mock(() => {}),
    clearRect: mock(() => {}),
    drawImage: mock(() => {}),
  } as unknown as CanvasRenderingContext2D;
  const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
  const paint = spyOn(peerlessShapeResizePreviewPaint, 'paint').mockResolvedValue();
  try {
    const view = render(<PeerlessShapeResizePreview run={run} preview={store} scale={1} />);
    set({ token: 'token', bounds: { x: 0, y: 0, width: 100, height: 100 } });
    expect(previewShapeRect).not.toHaveBeenCalled();
    expect(view.queryByTestId('pptx-shape-resize-preview')).toBeNull();

    set({ token: 'token', bounds: { x: 0, y: 0, width: 80, height: 100 }, request: request(80) });
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(1));
    set({ token: 'token', bounds: { x: 0, y: 0, width: 60, height: 100 }, request: request(60) });
    set({ token: 'token', bounds: { x: 0, y: 0, width: 40, height: 100 }, request: request(40) });
    expect(previewShapeRect).toHaveBeenCalledTimes(1);
    await act(async () => { pending[0]!.resolve(frame(80)); });
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(2));
    expect(previewShapeRect.mock.calls[1]![2]).toMatchObject({ width: 40 });
    expect(paint).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('pptx-shape-resize-preview').style.width).toBe('80px');
    await act(async () => { pending[1]!.resolve(frame(40)); });
    await waitFor(() => expect(paint).toHaveBeenCalledTimes(2));
    expect(view.getByTestId('pptx-shape-resize-preview').style.width).toBe('40px');

    set({ token: 'token', bounds: { x: 0, y: 0, width: 30, height: 100 }, request: request(30, 'old') });
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(3));
    set({ token: 'token', bounds: { x: 0, y: 0, width: 20, height: 100 }, request: request(20, 'new') });
    await act(async () => { pending[2]!.resolve(frame(30)); });
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(4));
    expect(paint).toHaveBeenCalledTimes(2);

    await act(async () => { pending[3]!.reject(new Error('preview unavailable')); });
    await waitFor(() => expect((view.getByTestId('pptx-shape-resize-preview') as HTMLCanvasElement).width).toBe(0));
  } finally {
    paint.mockRestore();
    getContext.mockRestore();
  }
});

it('uses the scale captured by each preview job and never lets an old pump paint a new job at old zoom', async () => {
  let value: PeerlessShapePreview | null = {
    token: 'token', bounds: { x: 0, y: 0, width: 80, height: 100 }, request: request(80),
  };
  const listeners = new Set<() => void>();
  const store: PeerlessShapePreviewStore = {
    getSnapshot: () => value,
    subscribe(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  };
  const pending: ReturnType<typeof deferred<SlideDisplayList>>[] = [];
  const previewShapeRect = mock(() => {
    const next = deferred<SlideDisplayList>();
    pending.push(next);
    return next.promise;
  });
  const run = {
    current: true,
    require: () => {},
    client: { previewShapeRect },
    images: { pin: () => ({ ready: Promise.resolve(), resolve: () => null, release: () => {} }) },
    ensureWholeFonts: async () => {},
  } as unknown as PeerlessPresentation;
  const context = {
    setTransform: mock(() => {}),
    clearRect: mock(() => {}),
    drawImage: mock(() => {}),
  } as unknown as CanvasRenderingContext2D;
  const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
  const paint = spyOn(peerlessShapeResizePreviewPaint, 'paint').mockResolvedValue();
  try {
    const view = render(<PeerlessShapeResizePreview run={run} preview={store} scale={1} />);
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(1));
    view.rerender(<PeerlessShapeResizePreview run={run} preview={store} scale={2} />);
    await act(async () => { pending[0]!.resolve(frame(80)); });
    await waitFor(() => expect(previewShapeRect).toHaveBeenCalledTimes(2));
    expect(paint).not.toHaveBeenCalled();
    await act(async () => { pending[1]!.resolve(frame(80)); });
    await waitFor(() => expect(paint).toHaveBeenCalledTimes(1));
    expect(paint.mock.calls[0]![3]).toBe(2);
    expect(view.getByTestId('pptx-shape-resize-preview').style.width).toBe('160px');
  } finally {
    paint.mockRestore();
    getContext.mockRestore();
  }
});
