import { deck, open } from './peerlessTestHarness';
import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import * as pptx from '@betteroffice/pptx';
import type { SlidePrimitive, TextBoxPrimitive } from '@betteroffice/pptx';
import { decodeCommandDisplayList, transformedRect } from '@betteroffice/pptx/worker';
import { PptxEditor, type PptxAsyncEditorProps, type PptxEditorApi, type PptxInThreadEditorProps,
  type PptxWorkerViewerProps } from '../PptxEditor';
import { PPTX_COMMAND_IDS } from '../commands/descriptors';
import { definePptxPlugin } from '../plugins/definePptxPlugin';
import type { PptxGeometryTarget, PptxPluginContext, PptxPluginGeometry,
  PptxPluginSelection } from '../plugins/types';
import { peerlessPainterFactories } from './usePeerlessPresentation';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import { createCommandFramePainter } from './commandFramePainter';

const { act, fireEvent, render, waitFor } = await import('@testing-library/react');
const restores: (() => void)[] = [];
afterEach(() => { for (const restore of restores.splice(0).reverse()) restore(); });

type Engine = 'IN' | 'PL';
type TextTarget = { slideId: string; shapeId: string; storyId: string };
type ReadContent = Extract<Awaited<ReturnType<PptxEditorApi['readContent']>>, { ok: true }>;
type SelectionKind = NonNullable<PptxPluginSelection>['target']['kind'];
type ViewerHasPlugins = Exclude<NonNullable<PptxWorkerViewerProps['plugins']>[number], never>;
const viewerHasNoPluginSlot: [ViewerHasPlugins] extends [never] ? true : false = true;
const selectionKinds: Record<SelectionKind, true> = { slide: true, shape: true, text: true };

function deferred() {
  let resolve!: () => void;
  const promise = new Promise<void>(done => { resolve = done; });
  return { promise, resolve };
}

function paintGate(engine: Engine) {
  let held: ReturnType<typeof deferred> | undefined;
  let entered = false;
  if (engine === 'IN') {
    const paint = pptx.paintSlide;
    const spy = spyOn(pptx, 'paintSlide').mockImplementation(async (...args) => {
      const gate = held;
      if (gate && args[0].canvas.dataset.testid === 'pptx-slide-canvas') { entered = true; await gate.promise; }
      return paint(...args);
    });
    restores.push(() => spy.mockRestore());
  } else {
    const spy = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options = {}) =>
      createCommandFramePainter(canvas, { ...options,
        async beforeCommit(current) {
          await options.beforeCommit?.(current);
          const gate = held;
          if (gate && options.onCommit) { entered = true; await gate.promise; }
        },
      }));
    restores.push(() => spy.mockRestore());
  }
  const release = () => { held?.resolve(); held = undefined; };
  restores.push(release);
  return { block() { entered = false; held = deferred(); }, release,
    async pending() { await waitFor(() => expect(entered).toBe(true)); } };
}

function anchorProbe(id = 'anchor-matrix') {
  const contexts: PptxPluginContext<null>[] = [];
  const plugin = definePptxPlugin({ id, createState: () => null,
    onEvent(context) { contexts.push(context); },
    overlay({ context }) { contexts.push(context); return <span data-testid={id} />; },
  });
  return { plugin, contexts, context: () => contexts[contexts.length - 1]! };
}

async function mountAnchors(engine: Engine, options: Partial<PptxAsyncEditorProps> = {}) {
  const bytes = options.file instanceof Uint8Array ? options.file : await deck();
  const bounds = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    if (this.tagName === 'CANVAS') return new DOMRect(30, 40,
      Number.parseFloat(this.style.width) || 960, Number.parseFloat(this.style.height) || 540);
    return new DOMRect(10, 20, 1000, 580);
  });
  restores.push(() => bounds.mockRestore());
  const settings = { ...options, file: bytes };
  if (engine === 'PL') {
    const value = await open(settings);
    const canvasBounds = spyOn(value.canvas, 'getBoundingClientRect').mockImplementation(() =>
      new DOMRect(30, 40, Number.parseFloat(value.canvas.style.width) || 960,
        Number.parseFloat(value.canvas.style.height) || 540));
    restores.push(() => canvasBounds.mockRestore());
    let currentApi: PptxAsyncEditorApi = value.api;
    settings.onReady = next => { currentApi = next; };
    return { ...value, api: () => currentApi, bytes,
      rerender(patch: Partial<PptxAsyncEditorProps>) {
        Object.assign(settings, patch);
        value.view.rerender(value.ui(settings.readOnly));
      } };
  }
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }];
  let api!: PptxEditorApi;
  const ui = () => <PptxEditor legacySyncApi experimentalWorkerOpen={false} fonts={fonts}
    {...settings as Partial<PptxInThreadEditorProps>} onReady={value => { api = value; }} />;
  const view = render(ui());
  await waitFor(() => expect(api).toBeDefined(), { timeout: 10_000 });
  return { view, api: () => api, bytes, client: undefined,
    canvas: view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement,
    stage: view.getByRole('application'),
    rerender(patch: Partial<PptxAsyncEditorProps>) { Object.assign(settings, patch); view.rerender(ui()); } };
}

async function content(value: Awaited<ReturnType<typeof mountAnchors>>): Promise<ReadContent> {
  const read = await value.api().readContent();
  if (!read.ok) throw new Error(read.failure.message);
  return read;
}

function range(read: ReadContent, target: TextTarget, anchor = 1, focus = 4): PptxGeometryTarget {
  return { kind: 'range', version: read.version, slideId: target.slideId, shapeId: target.shapeId,
    storyId: target.storyId, anchor, focus };
}

function targets(read: ReadContent, target: TextTarget): PptxGeometryTarget[] {
  const value = range(read, target);
  if (value.kind !== 'range') throw new Error('Expected range');
  return [{ kind: 'selection' }, value, { ...value, anchor: 4, focus: 1 },
    { ...value, anchor: 2, focus: 2 }, { ...value, version: 'stale', shapeId: 'missing' },
    { ...value, shapeId: 'missing' }, { ...value, storyId: 'missing' },
    { ...value, anchor: -1 }, { ...value, slideId: read.slides[1]?.id ?? 'not-presented' }];
}

async function parity(geometry: PptxPluginGeometry, values: PptxGeometryTarget[], codes: string[]) {
  const answers = values.map(target => geometry.getAnchorGeometry(target));
  expect(answers.map((answer): string => answer.ok ? 'ok' : answer.failure.code)).toEqual(codes);
  expect(await geometry.readAnchorGeometries(values)).toEqual(answers);
  for (const answer of answers) if (answer.ok) {
    expect(answer.layoutId).toBe(geometry.layout.id);
    expect(answer.version).toBe(geometry.layout.version);
    expect(answer.anchor.height).toBeGreaterThan(0);
  }
  return answers;
}

const normalCodes = ['ok', 'ok', 'ok', 'ok', 'stale-version', 'missing-target', 'missing-target', 'missing-target', 'layout-unavailable'];

async function selectText(probe: ReturnType<typeof anchorProbe>, read: ReadContent, target: TextTarget, anchor = 1, focus = 4) {
  await act(async () => {
    expect(await probe.context().navigation.selectText({ slideId: target.slideId, shapeId: target.shapeId, storyId: target.storyId, anchor, focus }, { expectVersion: read.version })).toEqual({ ok: true });
  });
  await waitFor(() => expect(probe.context().snapshot.selection?.target).toMatchObject({
    kind: 'text', shapeId: target.shapeId, storyId: target.storyId, anchor, focus }));
  await waitFor(() => expect(probe.context().geometry).toBeTruthy());
}

async function groupedDeck() {
  const zip = await JSZip.loadAsync(await deck());
  const path = 'ppt/slides/slide1.xml';
  const xml = await zip.file(path)!.async('string');
  const shape = xml.match(/<p:sp>.*?<\/p:sp>/)![0];
  const child = shape.replace('id="2" name="Text"', 'id="4" name="Child"')
    .replace('x="100000" y="100000"', 'x="4000000" y="2600000"')
    .replace('cx="6000000" cy="2400000"', 'cx="2000000" cy="1200000"');
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="5143500"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;
  zip.file(path, xml.replace('</p:spTree>', `${group}</p:spTree>`));
  return zip.generateAsync({ type: 'uint8array' });
}

for (const engine of ['IN', 'PL'] as const) describe(`plugin anchor matrix (${engine})`, () => {
  test('plugin anchor target matrix at non-unit zoom', async () => {
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const read = await content(value);
    const story = read.stories[0]!;
    await selectText(probe, read, story);
    for (const scale of [0.5, 1.5, 2]) {
      await act(async () => { expect(await value.api().commands.execute('zoom', { scale })).toMatchObject({ ok: true }); });
      await waitFor(() => expect(probe.context().geometry?.layout.zoom).toBe(scale));
      await act(async () => { await value.api().clearSelection(); });
      await waitFor(() => expect(probe.context().snapshot.selection?.target).toEqual({ kind: 'slide' }));
      await waitFor(() => expect(probe.context().geometry).toBeTruthy());
      const slide = probe.context().geometry!.getAnchorGeometry({ kind: 'selection' });
      expect(slide).toMatchObject({ ok: true });
      if (!slide.ok) throw new Error(slide.failure.message);
      expect(slide.rects).toEqual([slide.pageRect]);
      expect(slide.anchor).toEqual(slide.pageRect);
      await act(async () => { expect(await probe.context().navigation.selectShape({ slideId: story.slideId,
        shapeId: story.shapeId }, { expectVersion: read.version })).toEqual({ ok: true }); });
      await waitFor(() => expect(probe.context().snapshot.selection?.target).toEqual({ kind: 'shape', shapeId: story.shapeId }));
      await waitFor(() => expect(probe.context().geometry).toBeTruthy());
      const geometry = probe.context().geometry!;
      const shapeRect = geometry.getShapeRect(story.shapeId);
      expect(shapeRect).not.toBeNull();
      expect(geometry.getAnchorGeometry({ kind: 'selection' })).toMatchObject({ ok: true, rects: [shapeRect], anchor: shapeRect });
      for (const [anchor, focus] of [[1, 4], [4, 1], [2, 2]]) {
        await selectText(probe, read, story, anchor, focus);
        const geometry = probe.context().geometry!;
        const values = targets(read, story);
        const answers = await parity(geometry, values, normalCodes);
        expect(answers[0]).toEqual(geometry.getAnchorGeometry(range(read, story, anchor, focus)));
        if (answers[0].ok) expect(answers[0].rects.length).toBe(anchor === focus ? 0 : 1);
      }
    }
  });

  test('ranges on another slide refuse before resolving shape and story', async () => {
    const gate = paintGate(engine);
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const read = await content(value);
    const first = read.stories.find(story => story.slideId === read.slides[0]!.id)!;
    const second = read.stories.find(story => story.slideId === read.slides[1]!.id)!;
    const retained = probe.context().geometry!;
    const other = range(read, second);
    if (other.kind !== 'range') throw new Error('Expected range');
    const invalid = { ...other, shapeId: 'missing', storyId: 'missing', anchor: -1 };
    await parity(retained, [other, invalid, range(read, first)], ['layout-unavailable', 'layout-unavailable', 'ok']);
    gate.block();
    let navigation!: Promise<unknown>;
    act(() => { navigation = probe.context().navigation.goToSlide({ slideId: second.slideId }, { expectVersion: read.version }); });
    await gate.pending();
    await waitFor(() => expect(probe.context().geometry).toBeNull());
    await parity(retained, [other, invalid, range(read, first)], ['layout-unavailable', 'layout-unavailable', 'layout-unavailable']);
    await act(async () => { gate.release(); await navigation; });
    await waitFor(() => expect(probe.context().geometry?.layout.slideId).toBe(second.slideId));
    await parity(probe.context().geometry!, [other, invalid, range(read, first)], ['ok', 'missing-target', 'layout-unavailable']);
  });

  test('notes selection preserves slide geometry and rejects notes ranges', async () => {
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const beforeNotes = await content(value);
    await act(async () => {
      expect(await value.api().applyEdits({ expectVersion: beforeNotes.version,
        steps: [{ op: 'setSlideNotes', target: { slideId: beforeNotes.slides[0]!.id }, text: 'Speaker notes' }] })).toMatchObject({ ok: true });
    });
    await waitFor(() => expect(probe.context().geometry?.layout.version).toBe(probe.context().snapshot.version));
    const read = await content(value);
    const story = read.stories[0]!;
    await selectText(probe, read, story);
    const geometry = probe.context().geometry!;
    const selected = probe.context().snapshot.selection;
    const answer = geometry.getAnchorGeometry({ kind: 'selection' });
    const notes = value.view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement;
    await act(async () => { notes.focus(); notes.setSelectionRange(1, 4); fireEvent.select(notes); });
    expect(probe.context().snapshot.selection).toEqual(selected);
    expect(geometry.getAnchorGeometry({ kind: 'selection' })).toEqual(answer);
    await parity(geometry, [range(read, { ...story, storyId: `${story.slideId}:notes` }, 0, 0)], ['missing-target']);
    expect(read.stories.some(story => story.storyId.includes(':notes'))).toBe(false);
    expect(PPTX_COMMAND_IDS.some(id => /^(masterView|layoutView)$/.test(id))).toBe(false);
    expect(value.view.queryByRole('tab', { name: /^(Slide master|Master view|Layout view)$/i })).toBeNull();
  });

  test('proposal review gates every anchor target and restores matching geometry on exit', async () => {
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const first = await content(value);
    const story = first.stories[0]!;
    let proposalId = '';
    await act(async () => {
      const api = value.api();
      const edits = [{ type: 'replaceText' as const, storyId: story.storyId, start: 0, end: 1, text: 'Proposed' }];
      if ('handleAsync' in api) proposalId = (await (await api.handleAsync()).propose('Reviewer', null, edits)).id;
      else { proposalId = api.handle.propose('Reviewer', null, edits).id; api.refreshProposals(); }
      await api.commands.execute('proposalSelect', { proposalId });
      await api.commands.execute('proposalDiff', { enabled: false });
    });
    await waitFor(() => expect(probe.context().geometry).toBeTruthy());
    const read = await content(value);
    const retained: { geometry: PptxPluginGeometry; values: PptxGeometryTarget[] }[] = [];
    for (const mode of ['slide', 'shape', 'forward', 'backward', 'caret'] as const) {
      await act(async () => {
        if (mode === 'slide') await value.api().clearSelection();
        else if (mode === 'shape') await probe.context().navigation.selectShape(story, { expectVersion: read.version });
      });
      if (mode !== 'slide' && mode !== 'shape') await selectText(probe, read, story,
        mode === 'backward' ? 4 : mode === 'caret' ? 2 : 1, mode === 'backward' ? 1 : mode === 'caret' ? 2 : 4);
      await waitFor(() => expect(probe.context().geometry).toBeTruthy());
      retained.push({ geometry: probe.context().geometry!, values: targets(read, story) });
    }
    await act(async () => { expect(await value.api().commands.execute('proposalDiff', { enabled: true })).toMatchObject({ ok: true }); });
    await waitFor(() => expect(probe.context().geometry).toBeNull());
    for (const { geometry, values } of retained) await parity(geometry, values, values.map(() => 'layout-unavailable'));
    await act(async () => { expect(await value.api().commands.execute('proposalDiff', { enabled: false })).toMatchObject({ ok: true }); });
    await waitFor(() => expect(probe.context().geometry).toBeTruthy());
    await parity(probe.context().geometry!, targets(await content(value), story), normalCodes);
  });

  test("a removed plugin's retained anchors refuse while another plugin stays", async () => {
    const owner = anchorProbe('removed-anchors');
    const survivor = anchorProbe('surviving-anchors');
    const value = await mountAnchors(engine, { plugins: [owner.plugin, survivor.plugin] });
    await waitFor(() => expect(owner.context()?.geometry && survivor.context()?.geometry).toBeTruthy());
    const read = await content(value);
    const story = read.stories[0]!;
    await selectText(owner, read, story);
    await waitFor(() => expect(survivor.context().snapshot.selection?.target.kind).toBe('text'));
    const context = owner.context();
    const retained = context.geometry!;
    const values = targets(read, story);
    await parity(retained, values, normalCodes);
    await act(async () => { value.rerender({ plugins: [survivor.plugin] }); });
    expect(context.lifetimeSignal.aborted).toBe(true);
    expect(value.view.queryByTestId('removed-anchors')).toBeNull();
    expect(value.view.queryByTestId('surviving-anchors')).not.toBeNull();
    await parity(survivor.context().geometry!, values, normalCodes);
    await parity(retained, values, values.map(() => 'layout-unavailable'));
  });

  test('anchor target matrix follows version and presentation transitions', async () => {
    const gate = paintGate(engine);
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    let read = await content(value);
    const story = read.stories[0]!;
    await selectText(probe, read, story);
    const remote = pptx.openPresentation(value.bytes, { clientId: 9003 });
    const updates: Uint8Array[] = [];
    const off = remote.onUpdate(update => updates.push(update.slice()));
    try { remote.setSlideNotes(story.slideId, 'Remote notes'); }
    finally { off(); remote.dispose(); }
    for (const transition of ['remote', 'edit', 'undo', 'redo', 'replacement'] as const) {
      const retained = probe.context().geometry!;
      const retainedContext = probe.context();
      const values = targets(read, story);
      const generation = probe.context().snapshot.generation;
      const previousApi = value.api();
      gate.block();
      let operation!: Promise<unknown>;
      act(() => {
        const api = value.api();
        if (transition === 'remote') {
          if (value.client) operation = value.client.dispatch({ kind: 'applyRemoteUpdate', payload: {
            bytes: updates.map(update => update.slice().buffer), inputTime: 0, undoGroup: 'remote-anchors' } });
          else { for (const update of updates) (api as PptxEditorApi).handle.applyUpdate(update); (api as PptxEditorApi).refresh(); operation = Promise.resolve(); }
        } else if (transition === 'edit') operation = api.applyEdits({ expectVersion: read.version,
          steps: [{ op: 'setSlideNotes', target: { slideId: story.slideId }, text: 'Local notes' }] });
        else if (transition === 'replacement') { value.rerender({ file: value.bytes.slice() }); operation = Promise.resolve(); }
        else operation = api.commands.execute(transition, null);
      });
      await gate.pending();
      if (transition === 'replacement') {
        expect(retainedContext.lifetimeSignal.aborted).toBe(true);
        expect(value.view.queryByTestId('anchor-matrix')).toBeNull();
      } else await waitFor(() => expect(probe.context().geometry).toBeNull());
      await parity(retained, values, values.map(() => 'layout-unavailable'));
      await act(async () => { gate.release(); await operation; });
      if (transition === 'replacement') {
        await waitFor(() => expect(probe.context().snapshot.generation).not.toBe(generation));
        await waitFor(() => expect(value.api()).not.toBe(previousApi));
      }
      await waitFor(() => expect(probe.context().geometry).toBeTruthy());
      const next = await content(value);
      if (transition !== 'replacement') {
        expect(next.version).not.toBe(read.version);
        const old = values[1]!;
        expect(probe.context().geometry!.getAnchorGeometry(old)).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      } else await selectText(probe, next, next.stories[0]!);
      await parity(probe.context().geometry!, targets(next, next.stories[0]!), normalCodes);
      read = next;
    }
  });

  test('multi-shape selection with a grouped child is unreachable', async () => {
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { file: await groupedDeck(), plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const read = await content(value);
    const api = value.api();
    const snapshot = 'handleAsync' in api ? await (await api.handleAsync()).snapshot() : api.handle.snapshot();
    const slide = snapshot.slides[0]!;
    const group = slide.shapes.find(shape => shape.name === 'Group')!;
    const child = group.children[0]!;
    for (const shapeId of [slide.shapes[0]!.id, child.id]) {
      await act(async () => { expect(await probe.context().navigation.selectShape({ slideId: slide.id, shapeId },
        { expectVersion: read.version })).toEqual({ ok: true }); });
      await waitFor(() => expect(probe.context().snapshot.selection?.target).toEqual({ kind: 'shape', shapeId }));
      await waitFor(() => expect(probe.context().geometry).toBeTruthy());
      const geometry = probe.context().geometry!;
      const answer = geometry.getAnchorGeometry({ kind: 'selection' });
      expect(answer).toMatchObject({ ok: true, rects: [geometry.getShapeRect(shapeId)], anchor: geometry.getShapeRect(shapeId) });
      expect(await geometry.readAnchorGeometries([{ kind: 'selection' }])).toEqual([answer]);
    }
    expect(Object.keys(selectionKinds).sort()).toEqual(['shape', 'slide', 'text']);
    expect('selectShapes' in probe.context().navigation).toBe(false);
    expect('shapeIds' in probe.context().snapshot.selection!.target).toBe(false);
  });

  test('table cell text anchors resolve caret and range while cell-range selection is unreachable', async () => {
    const bytes = new Uint8Array(await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url)));
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { file: bytes, plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const read = await content(value);
    const api = value.api();
    const snapshot = 'handleAsync' in api ? await (await api.handleAsync()).snapshot() : api.handle.snapshot();
    const table = snapshot.slides[0]!.shapes.find(shape => (shape.graphic as { type?: string } | null)?.type === 'table')!;
    expect(table).toBeDefined();
    const stories = read.stories.filter(story => story.shapeId === table.id && story.text.length > 0);
    expect(stories.length).toBeGreaterThan(1);
    const frame = value.client ? decodeCommandDisplayList(value.client.state.frame!) : (api as PptxEditorApi).handle.layoutSlide(0);
    const cells = (primitives: SlidePrimitive[], parents: SlidePrimitive[] = []):
      { box: TextBoxPrimitive; transforms: SlidePrimitive[] }[] => primitives.flatMap(primitive => {
        const transforms = [...parents, primitive];
        return primitive.kind === 'textBox' ? [{ box: primitive, transforms }]
          : primitive.kind === 'table' || primitive.kind === 'chart' ? cells(primitive.primitives, transforms) : [];
      });
    for (const story of stories.slice(0, 2)) {
      const cell = cells(frame.primitives).find(({ box }) => box.shapeId === table.id && box.storyId === story.storyId)!;
      expect(cell).toBeDefined();
      const { box, transforms } = cell;
      expect(box).toBeDefined();
      const line = box.lines[0]!;
      const left = line.caretStops.find(stop => stop.position === 0)!.x;
      const right = line.caretStops.find(stop => stop.position === 1)!.x;
      for (const [anchor, focus] of [[0, 0], [0, 1], [1, 0]]) {
        await selectText(probe, read, story, anchor, focus);
        const geometry = probe.context().geometry!;
        const answer = geometry.getAnchorGeometry(range(read, story, anchor, focus));
        expect(answer.ok).toBe(true);
        if (!answer.ok) throw new Error(answer.failure.message);
        expect(answer.anchor.height).toBeGreaterThan(0);
        expect(answer.anchor.width).toBeGreaterThan(0);
        expect(answer.rects.length).toBe(anchor === focus ? 0 : 1);
        const expected = geometry.toOverlayRect({ space: 'slide-px', rect: transformedRect(transforms, {
          x: Math.min(left, right), y: line.y,
          width: anchor === focus ? 1 : Math.abs(right - left), height: line.height }) });
        expect(expected).not.toBeNull();
        expect(answer.anchor).toEqual(expected!);
        if (anchor !== focus) expect(answer.rects).toEqual([expected!]);
        expect(geometry.getAnchorGeometry({ kind: 'selection' })).toEqual(answer);
        expect(await geometry.readAnchorGeometries([{ kind: 'selection' }, range(read, story, anchor, focus)])).toEqual([answer, answer]);
      }
    }
    expect(probe.context().snapshot.selection?.target).toMatchObject({ kind: 'text', storyId: stories[1]!.storyId });
    expect('selectCells' in probe.context().navigation).toBe(false);
    expect(Object.keys(selectionKinds).sort()).toEqual(['shape', 'slide', 'text']);
    await parity(probe.context().geometry!, [range(read, { ...stories[0]!, storyId: 'missing-cell' }, 0, 0)], ['missing-target']);
  });

  test('empty text box caret uses paragraph line metrics', async () => {
    const zip = await JSZip.loadAsync(await deck());
    const path = 'ppt/slides/slide1.xml';
    zip.file(path, (await zip.file(path)!.async('string')).replace(/<a:p>.*?<\/a:p>/g,
      '<a:p><a:endParaRPr sz="2400"><a:latin typeface="Liberation Sans"/></a:endParaRPr></a:p>')
      .replace(/<a:p>.*?<\/a:p>(?=<a:p>)/g, ''));
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { file: await zip.generateAsync({ type: 'uint8array' }), plugins: [probe.plugin] });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const read = await content(value);
    const story = read.stories[0]!;
    expect(story.text).toBe('');
    await selectText(probe, read, story, 0, 0);
    const api = value.api();
    const frame = value.client ? decodeCommandDisplayList(value.client.state.frame!) : (api as PptxEditorApi).handle.layoutSlide(0);
    const text = frame.primitives.find(primitive => primitive.kind === 'textBox' && primitive.storyId === story.storyId) as TextBoxPrimitive;
    expect(text.lines.length).toBeGreaterThan(0);
    const line = text.lines[0]!;
    expect(line.height).toBeGreaterThan(0);
    expect(line.x).toBeGreaterThan(0);
    expect(line.y).toBeGreaterThan(0);
    const geometry = probe.context().geometry!;
    const expected = geometry.toOverlayRect({ space: 'slide-px', rect: transformedRect([text], {
      x: line.caretStops[0]?.x ?? line.x, y: line.y, width: 1, height: line.height }) });
    expect(expected).not.toBeNull();
    const answer = geometry.getAnchorGeometry({ kind: 'selection' });
    expect(answer).toMatchObject({ ok: true, rects: [], anchor: expected });
    expect(await geometry.readAnchorGeometries([{ kind: 'selection' }, range(read, story, 0, 0)])).toEqual([answer, answer]);
  });

  test('read-only and slideshow retain editor anchors while the viewer has no plugin host', async () => {
    const probe = anchorProbe();
    const value = await mountAnchors(engine, { plugins: [probe.plugin], readOnly: true });
    await waitFor(() => expect(probe.context()?.geometry).toBeTruthy());
    const read = await content(value);
    await selectText(probe, read, read.stories[0]!);
    expect(probe.context().snapshot.readOnly).toBe(true);
    await parity(probe.context().geometry!, targets(read, read.stories[0]!), normalCodes);
    await act(async () => { expect(await value.api().commands.execute('slideshow', null)).toMatchObject({ ok: true }); });
    await waitFor(() => expect(value.view.queryByRole('dialog')).not.toBeNull());
    const dialog = value.view.getByRole('dialog');
    expect(dialog.querySelector('[data-testid="plugin-overlays"]')).toBeNull();
    expect(dialog.querySelector('[data-testid="pptx-slide-canvas"]')).toBeNull();
    await parity(probe.context().geometry!, targets(read, read.stories[0]!), normalCodes);
    await act(async () => { fireEvent.keyDown(dialog, { key: 'ArrowRight' }); });
    await parity(probe.context().geometry!, targets(read, read.stories[0]!), normalCodes);
    expect(probe.context().geometry!.layout.slide).toBe(1);
    expect(viewerHasNoPluginSlot).toBe(true);
    await act(async () => { fireEvent.keyDown(dialog, { key: 'Escape' }); });
    await waitFor(() => expect(value.view.queryByRole('dialog')).toBeNull());
    await parity(probe.context().geometry!, targets(read, read.stories[0]!), normalCodes);
  });
});
