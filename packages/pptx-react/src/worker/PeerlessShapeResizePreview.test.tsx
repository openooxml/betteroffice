import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, mock, spyOn } from 'bun:test';
import type { SlideDisplayList } from '@betteroffice/pptx';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import type { PeerlessShapePreview, PeerlessShapePreviewStore } from './usePeerlessShapeGestures';
import { PeerlessShapeResizePreview, peerlessShapeResizePreviewPaint } from './PeerlessShapeResizePreview';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
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
