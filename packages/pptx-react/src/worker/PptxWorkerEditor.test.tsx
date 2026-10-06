import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { StrictMode } from 'react';
import * as pptx from '@betteroffice/pptx';
import type {
  DeckSnapshot, PptxWorkerEditorAccess, PptxWorkerEditorFrame, PptxWorkerEditorOperation,
  PptxWorkerEditorSession, PptxWorkerEditorState, SlideDisplayList, StorySnapshot,
} from '@betteroffice/pptx';
import { PptxEditor } from '../PptxEditor';
import { definePptxPlugin } from '../plugins/definePptxPlugin';
import type { PptxPluginContext } from '../plugins/types';
import { presentationSessionOpener } from '../viewer/useSessionPresentation';
import { EditablePresentation, workerEditorSessionOpener } from './useEditableSessionPresentation';
import type { PptxWorkerEditorApi } from './createWorkerEditorApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const file = new Uint8Array([1]);
const restorers: (() => void)[] = [];
let painted: SlideDisplayList[];
let delayPaint: (() => Promise<void>) | undefined;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

const settle = () => act(() => new Promise<void>((done) => { setTimeout(done); }));

function session(ready = true, slideCount = 1) {
  let text = 'x';
  let epoch = 0;
  const listeners = new Set<() => void>();
  const methods: string[] = [];
  const hydrated = deferred<void>();
  const story = (): StorySnapshot => ({ id: 'story', length: text.length,
    paragraphs: [{ id: 'p', alignment: 'l', level: 0, bulletJson: null,
      runs: [{ text, style: { bold: null, italic: null, fontSizePt: null, color: null, fontFamily: null, underline: null } }] }] });
  const snapshot: DeckSnapshot = { widthEmu: 12700000, heightEmu: 7143750, slides: [{ id: 's',
    sourcePartPath: null, layoutPartPath: null, name: null, notes: '', shapes: [{ id: 'shape', sourceId: 1,
      kind: 'shape', name: 'text', x: 0, y: 0, width: 1000000, height: 1000000, rotationDeg: 0,
      flipH: false, flipV: false, geometry: 'rect', adjustValues: {}, placeholder: null, fill: null,
      resolvedFillColor: null, outline: null, resolvedOutlineColor: null, mediaPartPath: null,
      graphic: null, textStories: [story()], children: [] }] }] };
  for (let index = 1; index < slideCount; index += 1)
    snapshot.slides.push({ ...snapshot.slides[0], id: `s${index + 1}`, shapes: [] });
  const peerFrame: SlideDisplayList = { contractVersion: 1, width: 960, height: 540, primitives: [] };
  const state: PptxWorkerEditorState = { stage: ready ? 'ready' : 'hydrating', sequence: 0,
    acknowledgedSequence: 0, version: 'v0', projection: { format: 'pptx', stage: 'ready', version: 0,
      dirty: false, size: { width: snapshot.widthEmu, height: snapshot.heightEmu },
      slides: snapshot.slides.map((slide, index) => ({ id: slide.id, index, name: null, layoutPartPath: null })) } };
  let current = state;
  const publish = () => { for (const listener of listeners) listener(); };
  const access = {
    clientId: 1, snapshot: () => snapshot, story, version: () => current.version!,
    layoutSlide: mock(() => peerFrame), isProposalsAvailable: () => true, listProposals: () => [],
    canUndo: () => true, canRedo: () => false, onUpdate: () => () => {},
    readContent: () => ({ ok: true, version: current.version, slides: [], stories: [{ id: 'story', text }] }),
    interaction: { snapshot: () => ({ snapshot, keys: { s: 'k' } }), key: () => 'k', activate: () => true,
      hitTest: () => ({ shapeId: 'shape', storyId: 'story', textPosition: 0 }) },
    insertText() {}, deleteText() {}, formatText() {}, setSlideNotes() {}, applyEdits() {},
  } as unknown as PptxWorkerEditorAccess;
  Object.defineProperty(access, Symbol.for('@betteroffice/pptx/slide-layout-cache'), { value: access.interaction });
  const frames: PptxWorkerEditorFrame[] = [];
  const owner: PptxWorkerEditorSession = {
    get state() { return current; }, get hydrated() { return current.stage === 'ready'; },
    get failure() { return current.failure; }, subscribe(listener) { const notify = () => listener(current);
      listeners.add(notify); return () => { listeners.delete(notify); }; },
    whenHydrated: async () => { if (!owner.hydrated) await hydrated.promise; },
    handleAsync: async () => { await owner.whenHydrated(); return access; },
    apply: mock((op: PptxWorkerEditorOperation) => {
      methods.push(op.method);
      if (op.method === 'insertText') text = text.slice(0, op.args[1]) + op.args[2] + text.slice(op.args[1]);
      if (op.method === 'setSlideNotes') snapshot.slides[0].notes = op.args[1];
      const batch = op.method === 'applyEdits' ? op.args[0] : undefined;
      const refused = !!batch && batch.expectVersion !== current.version;
      const applied = !refused && (!batch || batch.steps.length > 0);
      if (batch && applied) for (const step of batch.steps) {
        if (step.op === 'setSlideNotes') snapshot.slides[0].notes = step.text;
      }
      const sequence = current.sequence + (refused ? 0 : 1);
      current = { ...current, sequence, acknowledgedSequence: sequence, version: applied ? `v${sequence}` : current.version };
      publish();
      return { sequence, consumed: !refused, revision: sequence, version: current.version!, engineVersion: current.version!,
        outcome: { result: batch ? refused
          ? { ok: false, version: current.version, failure: { code: 'stale-version', message: 'Version changed' } }
          : { ok: true, version: current.version, applied } : true,
          applied, changedTargets: [], canUndo: true, canRedo: false } };
    }),
    flush: mock(async () => {}), saveAsync: mock(async () => new TextEncoder().encode(text)),
    frame: mock(async (slideId: string) => {
      const frame = { slideId, slideIndex: snapshot.slides.findIndex((slide) => slide.id === slideId),
        sequence: current.sequence, version: current.version!,
        epoch: ++epoch, media: new Map(), displayList: { ...peerFrame, primitives: [] } };
      frames.push(frame);
      return frame;
    }),
    recoverySave: mock(async () => ({ bytes: new TextEncoder().encode(text), recovery: true as const })),
    dispose: mock(async () => { current = { ...current, stage: 'disposed' }; hydrated.resolve(); publish(); }),
  };
  return { owner, access, methods, peerFrame, frames, snapshot,
    hydrate() { current = { ...current, stage: 'ready' }; hydrated.resolve(); publish(); } };
}

function open(value: ReturnType<typeof session>) {
  const spy = spyOn(workerEditorSessionOpener, 'open').mockReturnValue(value.owner);
  restorers.push(() => spy.mockRestore());
  return spy;
}

beforeEach(() => {
  painted = [];
  delayPaint = undefined;
  const context = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, 'getContext')!;
  Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
    ...context, value(this: HTMLCanvasElement) {
      return { canvas: this, clearRect() {}, setTransform() {}, fillRect() {} } as unknown as CanvasRenderingContext2D;
    },
  });
  const paint = spyOn(pptx, 'paintSlide').mockImplementation(async (ctx, frame) => {
    if (delayPaint) await delayPaint();
    ctx.fillRect(0, 0, 1, 1);
    painted.push(frame);
  });
  const observer = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class { observe() {} disconnect() {} } as unknown as typeof ResizeObserver;
  restorers.push(() => Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', context),
    () => paint.mockRestore(), () => { globalThis.ResizeObserver = observer; });
});
afterEach(async () => {
  await settle();
  cleanup();
  await settle();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

test('dispatches_only_flagged_editable_props', async () => {
  const value = session();
  const worker = open(value);
  const viewer = spyOn(presentationSessionOpener, 'open').mockRejectedValue(new Error('viewer sentinel'));
  const main = spyOn(pptx, 'initWasm').mockResolvedValue(undefined);
  const local = spyOn(pptx, 'openPresentation').mockImplementation(() => { throw new Error('local sentinel'); });
  restorers.push(() => viewer.mockRestore(), () => main.mockRestore(), () => local.mockRestore());
  const view = render(<PptxEditor fonts={[]} file={file} clientId={42} experimentalWorkerOpen />);
  await waitFor(() => expect(worker).toHaveBeenCalledTimes(1));
  expect(main).not.toHaveBeenCalled();
  expect(local).not.toHaveBeenCalled();
  expect(viewer).not.toHaveBeenCalled();
  expect(worker.mock.calls[0][1]?.clientId).toBe(42);
  view.rerender(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen readOnly />);
  await waitFor(() => expect(viewer).toHaveBeenCalledTimes(1));
  expect(worker).toHaveBeenCalledTimes(1);
  expect(main).not.toHaveBeenCalled();
  expect(local).not.toHaveBeenCalled();
  view.rerender(<PptxEditor fonts={[]} file={file} />);
  await waitFor(() => expect(main).toHaveBeenCalledTimes(1));
  await waitFor(() => expect(local).toHaveBeenCalledTimes(1));
  expect(worker).toHaveBeenCalledTimes(1);
  expect(viewer).toHaveBeenCalledTimes(1);
  worker.mockReturnValueOnce(session().owner);
  view.rerender(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen readOnly={false} />);
  await waitFor(() => expect(worker).toHaveBeenCalledTimes(2));
  expect(viewer).toHaveBeenCalledTimes(1);
  expect(main).toHaveBeenCalledTimes(1);
  expect(local).toHaveBeenCalledTimes(1);
  const error = new pptx.PptxWorkerEditorCollaborationError();
  worker.mockImplementationOnce(() => { throw error; });
  const failed = mock(() => {});
  const ready = mock(() => {});
  view.rerender(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen collaboration={{ clientId: 1 }}
    onError={failed} onReady={ready} />);
  await waitFor(() => expect(failed).toHaveBeenCalledWith(error));
  expect(ready).not.toHaveBeenCalled();
  expect(main).toHaveBeenCalledTimes(1);
  expect(local).toHaveBeenCalledTimes(1);
});

test('preparing_state_disables_every_edit_entry', async () => {
  const value = session(false);
  open(value);
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen onReady={ready} />);
  await waitFor(() => expect(view.getByTestId('pptx-notes-textarea')).toBeDefined());
  expect((view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).disabled).toBe(true);
  expect((view.getByTestId('pptx-insert-image-input') as HTMLInputElement).disabled).toBe(true);
  fireEvent.keyDown(view.getByRole('application'), { key: 'a' });
  expect(value.methods).toEqual([]);
  expect(ready).not.toHaveBeenCalled();
  expect(view.getByTestId('pptx-worker-status').textContent).toBe('Preparing editor. Editing will be available shortly.');
  await act(async () => { value.hydrate(); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
});

test('hydration_keeps_navigation_and_the_painted_worker_slide', async () => {
  const value = session(false, 2);
  open(value);
  const ready = mock(() => {});
  let run!: EditablePresentation;
  let adopted!: PptxWorkerEditorFrame;
  const didPaint = EditablePresentation.prototype.didPaint;
  const spy = spyOn(EditablePresentation.prototype, 'didPaint').mockImplementation(function (
    this: EditablePresentation, frame, navigation
  ) {
    if (this.currentPaint(frame, navigation)) { run = this; adopted = frame; }
    didPaint.call(this, frame, navigation);
  });
  restorers.push(() => spy.mockRestore());
  const view = render(<PptxEditor fonts={[]} file={file} initialSlide={1} experimentalWorkerOpen onReady={ready} />);
  await waitFor(() => expect(view.container.querySelectorAll('aside button')).toHaveLength(2));
  const second = view.container.querySelectorAll('aside button')[1];
  fireEvent.click(second);
  await waitFor(() => expect(second.getAttribute('aria-current')).toBe('page'));
  await waitFor(() => expect(adopted?.slideId).toBe('s2'));
  expect(ready).not.toHaveBeenCalled();
  await act(async () => { value.hydrate(); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  expect(second.getAttribute('aria-current')).toBe('page');
  expect(run.active).toBe(1);
  expect(adopted.slideId).toBe('s2');
  expect(adopted.slideIndex).toBe(1);
  expect(adopted).toBe(run.frame(1)!);
  expect(painted).toContain(adopted.displayList);
  expect(painted).not.toContain(value.peerFrame);
});

for (const initialSlide of [Number.NaN, 1.5]) test(`invalid_initial_slide_${initialSlide}_opens_the_first_slide`, async () => {
  const value = session(false, 2);
  open(value);
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={file} initialSlide={initialSlide} experimentalWorkerOpen
    onReady={ready} />);
  await waitFor(() => expect(view.container.querySelectorAll('aside button')).toHaveLength(2));
  await act(async () => { value.hydrate(); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  expect(view.container.querySelectorAll('aside button')[0].getAttribute('aria-current')).toBe('page');
  expect(painted).toContain(value.frames.find((frame) => frame.slideId === 's')!.displayList);
});

test('applyEdits_publishes_only_applied_batches_and_waits_for_replay', async () => {
  const value = session();
  open(value);
  let api!: PptxWorkerEditorApi;
  const changes = mock((_snapshot: DeckSnapshot) => {});
  render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen onChange={changes}
    onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).toBeDefined());
  const held = deferred<void>();
  const flush = spyOn(value.owner, 'flush').mockImplementationOnce(() => held.promise);
  restorers.push(() => flush.mockRestore());
  let settled = false;
  let pending!: ReturnType<PptxWorkerEditorApi['applyEdits']>;
  await act(async () => {
    pending = api.applyEdits({ expectVersion: value.owner.state.version!,
      steps: [{ op: 'setSlideNotes', target: { slideId: 's' }, text: 'Batch notes' }] });
    void pending.then(() => { settled = true; });
  });
  await waitFor(() => expect(flush).toHaveBeenCalledTimes(1));
  expect(changes).toHaveBeenCalledTimes(1);
  expect(changes.mock.calls[0][0].slides[0].notes).toBe('Batch notes');
  expect(settled).toBe(false);
  await act(async () => { held.resolve(); expect(await pending).toMatchObject({ ok: true, applied: true }); });
  expect(settled).toBe(true);
  expect(changes).toHaveBeenCalledTimes(1);
  await act(async () => {
    expect(await api.applyEdits({ expectVersion: value.owner.state.version!, steps: [] }))
      .toMatchObject({ ok: true, applied: false });
  });
  expect(changes).toHaveBeenCalledTimes(1);
  await act(async () => {
    expect(await api.applyEdits({ expectVersion: 'stale', steps: [] }))
      .toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  });
  expect(changes).toHaveBeenCalledTimes(1);
  expect(value.methods).toEqual(['applyEdits', 'applyEdits', 'applyEdits']);
  expect(flush).toHaveBeenCalledTimes(3);
});

test('commands_plugins_notes_and_keyboard_share_replay', async () => {
  const value = session();
  open(value);
  let api!: PptxWorkerEditorApi;
  let context!: PptxPluginContext<null>;
  const plugin = definePptxPlugin({ id: 'writer', createState: () => null, initialize(ctx) { context = ctx; } });
  const view = render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen plugins={[plugin]}
    pluginGrants={{ writer: { document: 'write', editBatches: true } }} onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).toBeDefined());
  await waitFor(() => expect(context).toBeDefined());
  await act(async () => { api.selectText({ slide: 1, shapeId: 'shape', storyId: 'story', start: 0, end: 1 }); });
  await act(async () => { await api.commands.execute('bold', null); });
  await act(async () => { await context.run(async (fresh) => { await fresh.edits!.applyEdits({ expectVersion: value.owner.state.version!, steps: [] }); }); });
  fireEvent.change(view.getByTestId('pptx-notes-textarea'), { target: { value: 'notes' } });
  await act(async () => { api.selectText({ slide: 1, shapeId: 'shape', storyId: 'story', start: 1, end: 1 }); });
  await act(async () => { fireEvent.keyDown(view.getByRole('application'), { key: 'a' }); });
  expect(value.methods.filter((method) => method !== 'addUndoBoundary')).toEqual(['formatText', 'applyEdits', 'setSlideNotes', 'insertText']);
});

test('paint_and_interaction_frames_are_separate', async () => {
  const value = session();
  open(value);
  let api!: PptxWorkerEditorApi;
  let run!: EditablePresentation;
  const didPaint = EditablePresentation.prototype.didPaint;
  const spy = spyOn(EditablePresentation.prototype, 'didPaint').mockImplementation(function (
    this: EditablePresentation, frame, navigation
  ) { run = this; didPaint.call(this, frame, navigation); });
  restorers.push(() => spy.mockRestore());
  render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).toBeDefined());
  expect(value.access.layoutSlide).toHaveBeenCalled();
  expect(painted.length).toBeGreaterThan(0);
  expect(painted).not.toContain(value.peerFrame);
  expect(painted.every((list) => value.frames.some((frame) => frame.displayList === list))).toBe(true);
  const access = await act(() => api.handleAsync());
  await act(async () => {
    access.setSlideNotes('s', 'new');
    expect(run.overlays).toBe(false);
  });
  await waitFor(() => expect(value.frames.some((frame) => frame.sequence === value.owner.state.sequence)).toBe(true));
  expect(painted).not.toContain(value.peerFrame);
  const frame = run.frame(0)!;
  expect(run.currentPaint(frame, run.navigation + 1)).toBe(false);
  await waitFor(() => expect(run.overlays).toBe(true));
});

test('thumbnail_retains_evicted_images_until_paint_settles', async () => {
  const value = session(true, 2);
  open(value);
  const ids = Array.from({ length: 26 }, (_, index) => `image-${index}`);
  const bitmaps = ids.map(() => ({ width: 1, height: 1, close: mock(() => {}) }));
  const decodes = ids.map(() => deferred<ImageBitmap>());
  const firstDecoded = deferred<void>();
  const released = deferred<void>();
  const release = mock(() => {});
  const closedAtDraw: number[] = [];
  const draw = mock((_image: CanvasImageSource) => { closedAtDraw.push(bitmaps[0].close.mock.calls.length); });
  let thumbnail!: SlideDisplayList;
  let resolverCount = 0;
  const frame = value.owner.frame;
  value.owner.frame = async (slideId) => {
    const result = await frame(slideId);
    if (slideId === 's2') {
      thumbnail = result.displayList;
      result.media = new Map(ids.map((id, index) => [id, new Uint8Array([index])]));
    }
    return result;
  };
  const decode = spyOn(pptx, 'decodePresentationImage').mockImplementation((bytes) => decodes[bytes[0]].promise);
  const thumbnailImages = EditablePresentation.prototype.thumbnailImages;
  const resolvers = spyOn(EditablePresentation.prototype, 'thumbnailImages').mockImplementation(function (
    this: EditablePresentation, list
  ) {
    const images = thumbnailImages.call(this, list);
    if (list === thumbnail && images) {
      resolverCount += 1;
      const releaseImages = images.release;
      images.release = () => { releaseImages(); release(); released.resolve(); };
    }
    return images;
  });
  const paint = spyOn(pptx, 'paintSlide').mockImplementation(async (ctx, list, _dpr, _scale, options) => {
    if (list !== thumbnail || !ctx.canvas.closest('aside')) {
      ctx.fillRect(0, 0, 1, 1);
      painted.push(list);
      return;
    }
    const pending = ids.map((id) => options!.resolveImage!(id));
    void Promise.resolve(pending[0]).then(() => firstDecoded.resolve());
    const images = await Promise.all(pending);
    ctx.drawImage = draw;
    ctx.drawImage(images[0]!, 0, 0);
  });
  restorers.push(() => decode.mockRestore(), () => resolvers.mockRestore(),
    () => paint.mockRestore());
  render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen />);
  await waitFor(() => expect(decode).toHaveBeenCalledTimes(ids.length));
  expect(resolverCount).toBe(1);
  expect(decode).toHaveBeenCalledTimes(26);
  expect(draw).not.toHaveBeenCalled();
  expect(release).not.toHaveBeenCalled();
  await act(async () => { decodes[0].resolve(bitmaps[0] as unknown as ImageBitmap); await firstDecoded.promise; });
  expect(bitmaps[0].close).not.toHaveBeenCalled();
  expect(release).not.toHaveBeenCalled();
  await act(async () => {
    decodes.slice(1).forEach((decode, index) => decode.resolve(bitmaps[index + 1] as unknown as ImageBitmap));
    await released.promise;
  });
  expect(draw).toHaveBeenCalledTimes(1);
  expect(draw.mock.calls[0][0]).toBe(bitmaps[0]);
  expect(closedAtDraw).toEqual([0]);
  expect(release).toHaveBeenCalledTimes(1);
  expect(bitmaps[0].close).toHaveBeenCalledTimes(1);
  expect(bitmaps.slice(1).every((bitmap) => bitmap.close.mock.calls.length === 0)).toBe(true);
});

test('late_media_cannot_overpaint_new_generation', async () => {
  const old = session();
  const next = session();
  const spy = open(old).mockReturnValueOnce(old.owner).mockReturnValue(next.owner);
  const held = deferred<void>();
  let started = 0;
  delayPaint = () => { started += 1; return held.promise; };
  const view = render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen />);
  await waitFor(() => expect(started).toBeGreaterThan(0));
  delayPaint = undefined;
  view.rerender(<PptxEditor fonts={[]} file={new Uint8Array([2])} experimentalWorkerOpen />);
  await waitFor(() => expect(painted.some((list) => next.frames.some((frame) => frame.displayList === list))).toBe(true));
  await act(async () => { held.resolve(); });
  expect(spy).toHaveBeenCalledTimes(2);
  expect(painted.some((list) => old.frames.some((frame) => frame.displayList === list))).toBe(false);
});

test('replacement_preserves_old_recovery', async () => {
  const old = session();
  const next = session();
  open(old).mockReturnValueOnce(old.owner).mockReturnValue(next.owner);
  let api!: PptxWorkerEditorApi;
  const view = render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).toBeDefined());
  const retired = api;
  const access = await act(() => api.handleAsync());
  await act(async () => { access.insertText('story', 1, 'old'); });
  view.rerender(<PptxEditor fonts={[]} file={new Uint8Array([2])} experimentalWorkerOpen onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).not.toBe(retired));
  expect(await act(() => retired.recoverySave())).toEqual({ bytes: new TextEncoder().encode('xold'), recovery: true });
  await act(() => expect(retired.saveAsync()).rejects.toBeInstanceOf(pptx.PptxWorkerEditorDisposedError));
  view.unmount();
  expect((await act(() => retired.recoverySave())).bytes).toEqual(new TextEncoder().encode('xold'));
});

test('strictmode_disposes_partial_hydration', async () => {
  const owners: ReturnType<typeof session>[] = [];
  const spy = spyOn(workerEditorSessionOpener, 'open').mockImplementation(() => {
    const value = session(false); owners.push(value); return value.owner;
  });
  restorers.push(() => spy.mockRestore());
  const ready = mock(() => {});
  const view = render(<StrictMode><PptxEditor fonts={[]} file={file} experimentalWorkerOpen onReady={ready} /></StrictMode>);
  await waitFor(() => expect(owners.length).toBe(2));
  expect(owners[0].owner.dispose).toHaveBeenCalledTimes(1);
  await act(async () => { owners[0].hydrate(); });
  expect(ready).not.toHaveBeenCalled();
  view.unmount();
  expect(owners[1].owner.dispose).toHaveBeenCalledTimes(1);
});

test('proposal_visuals_and_slideshow_refuse_with_reason', async () => {
  const value = session();
  open(value);
  let api!: PptxWorkerEditorApi;
  const view = render(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).toBeDefined());
  for (const id of ['exportPng', 'slideshow', 'proposalDiff', 'proposalSelect'] as const) {
    const state = api.commands.getState(id);
    expect(state.enabled).toBe(false);
    if (!state.enabled) expect(state.disabledReason.message).toBe('This feature is unavailable in this editing mode.');
  }
  expect(view.queryByTestId('pptx-canvas-review-toolbar')).toBeNull();
  expect(value.access).not.toHaveProperty('exportMarkdown');
  expect(value.access).not.toHaveProperty('exportStructured');
});
