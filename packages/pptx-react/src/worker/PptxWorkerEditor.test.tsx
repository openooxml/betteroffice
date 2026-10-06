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

function session(ready = true) {
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
  const peerFrame: SlideDisplayList = { contractVersion: 1, width: 960, height: 540, primitives: [] };
  const state: PptxWorkerEditorState = { stage: ready ? 'ready' : 'hydrating', sequence: 0,
    acknowledgedSequence: 0, version: 'v0', projection: { format: 'pptx', stage: 'ready', version: 0,
      dirty: false, size: { width: snapshot.widthEmu, height: snapshot.heightEmu },
      slides: [{ id: 's', index: 0, name: null, layoutPartPath: null }] } };
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
      const sequence = current.sequence + 1;
      current = { ...current, sequence, acknowledgedSequence: sequence, version: `v${sequence}` };
      publish();
      return { sequence, consumed: true, revision: sequence, version: current.version!, engineVersion: current.version!,
        outcome: { result: op.method === 'applyEdits' ? { ok: true, version: current.version, applied: 1 } : true,
          applied: true, changedTargets: [], canUndo: true, canRedo: false } };
    }),
    flush: mock(async () => {}), saveAsync: mock(async () => new TextEncoder().encode(text)),
    frame: mock(async () => {
      const frame = { slideId: 's', slideIndex: 0, sequence: current.sequence, version: current.version!,
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
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return { canvas: this, clearRect() {}, setTransform() {}, fillRect() {} } as unknown as CanvasRenderingContext2D;
  } as HTMLCanvasElement['getContext']);
  const paint = spyOn(pptx, 'paintSlide').mockImplementation(async (ctx, frame) => {
    if (delayPaint) await delayPaint();
    ctx.fillRect(0, 0, 1, 1);
    painted.push(frame);
  });
  const observer = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class { observe() {} disconnect() {} } as unknown as typeof ResizeObserver;
  restorers.push(() => context.mockRestore(), () => paint.mockRestore(), () => { globalThis.ResizeObserver = observer; });
});
afterEach(() => { cleanup(); for (const restore of restorers.reverse()) restore(); restorers.length = 0; });
afterAll(() => { if (ownsDom) GlobalRegistrator.unregister(); });

test('dispatches_only_flagged_editable_props', async () => {
  const value = session();
  const worker = open(value);
  const viewer = spyOn(presentationSessionOpener, 'open').mockRejectedValue(new Error('viewer sentinel'));
  const main = spyOn(pptx, 'initWasm').mockRejectedValue(new Error('main sentinel'));
  restorers.push(() => viewer.mockRestore(), () => main.mockRestore());
  const view = render(<PptxEditor fonts={[]} file={file} clientId={42} experimentalWorkerOpen />);
  await waitFor(() => expect(worker).toHaveBeenCalledTimes(1));
  expect(main).not.toHaveBeenCalled();
  expect(worker.mock.calls[0][1]?.clientId).toBe(42);
  view.rerender(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen readOnly />);
  await waitFor(() => expect(viewer).toHaveBeenCalledTimes(1));
  view.rerender(<PptxEditor fonts={[]} file={file} />);
  await waitFor(() => expect(main).toHaveBeenCalledTimes(1));
  expect(worker).toHaveBeenCalledTimes(1);
  const error = new pptx.PptxWorkerEditorCollaborationError();
  worker.mockImplementationOnce(() => { throw error; });
  const failed = mock(() => {});
  const ready = mock(() => {});
  view.rerender(<PptxEditor fonts={[]} file={file} experimentalWorkerOpen collaboration={{ clientId: 1 }}
    onError={failed} onReady={ready} />);
  await waitFor(() => expect(failed).toHaveBeenCalledWith(error));
  expect(ready).not.toHaveBeenCalled();
  expect(main).toHaveBeenCalledTimes(1);
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
  expect(view.getByTestId('pptx-worker-status').textContent).toContain('Preparing editor');
  await act(async () => { value.hydrate(); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
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
  fireEvent.keyDown(view.getByRole('application'), { key: 'a' });
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
  const access = await api.handleAsync();
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
  const access = await api.handleAsync();
  await act(async () => { access.insertText('story', 1, 'old'); });
  view.rerender(<PptxEditor fonts={[]} file={new Uint8Array([2])} experimentalWorkerOpen onReady={(ready) => { api = ready; }} />);
  await waitFor(() => expect(api).not.toBe(retired));
  expect(await retired.recoverySave()).toEqual({ bytes: new TextEncoder().encode('xold'), recovery: true });
  await expect(retired.saveAsync()).rejects.toBeInstanceOf(pptx.PptxWorkerEditorDisposedError);
  view.unmount();
  expect((await retired.recoverySave()).bytes).toEqual(new TextEncoder().encode('xold'));
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
    if (!state.enabled) expect(state.disabledReason.message).toContain('worker editor');
  }
  expect(view.queryByTestId('pptx-canvas-review-toolbar')).toBeNull();
  expect(value.access).not.toHaveProperty('exportMarkdown');
  expect(value.access).not.toHaveProperty('exportStructured');
});
