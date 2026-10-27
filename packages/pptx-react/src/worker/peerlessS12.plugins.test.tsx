import { deck, open, replyControl } from './peerlessTestHarness';
import JSZip from 'jszip';
import { expect, test } from 'bun:test';
import type { PptxPluginContext, PptxPluginRefusal, PptxPluginGrant } from '../plugins/types';
import { definePptxPlugin } from '../plugins/definePptxPlugin';
import type { PptxAsyncEditorProps } from '../PptxEditor';
const { act, fireEvent, waitFor } = await import('@testing-library/react');

test('plugin_grants', async () => {
  let context: PptxPluginContext<null> | undefined;
  const plugin = definePptxPlugin({ id: 'review', createState: () => null,
    onEvent(next) { context = next; },
    panel: { title: 'Review', placement: 'right', render: () => <div>Plugin panel</div> },
    commands: [{ id: 'check', label: 'Check', mutatesDocument: false, execute: () => ({ ok: true, status: 'executed' }) }],
  });
  const plugins = [plugin];
  const grants: Record<string, PptxPluginGrant> = { review: { document: 'write', editBatches: true } };
  const value = await open({ plugins, pluginGrants: grants });
  await waitFor(() => expect(context).toBeDefined());
  expect(value.view.getByText('Plugin panel')).toBeDefined();
  expect(await value.api.commands!.execute('plugin:review/check', null)).toMatchObject({ ok: true });
  const version = await value.api.version();
  replyControl.hold();
  let pending!: Promise<boolean>;
  act(() => { pending = value.api.goToSlide(2); });
  await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
  let outcome: unknown;
  const invocation = context!;
  const work = invocation.edits!.applyEdits({ expectVersion: version, steps: [{ op: 'setSlideNotes', target: { slideId: value.client.state.slides[0]!.id }, text: 'Denied' }] }).then((result) => { outcome = result; });
  grants.review = {};
  await act(async () => { value.view.rerender(value.ui()); });
  await act(async () => { replyControl.release(); await Promise.allSettled([pending, work]); });
  expect((outcome as PptxPluginRefusal).ok).toBe(false);
  expect(['aborted', 'document-replaced', 'plugin-unavailable', 'input-failed', 'permission-denied']).toContain((outcome as PptxPluginRefusal).failure.code);
});


test('plugin_shape_navigation_and_group_geometry', async () => {
  const zip = await JSZip.loadAsync(await deck());
  const path = 'ppt/slides/slide1.xml';
  const original = await zip.file(path)!.async('string');
  const shape = original.match(/<p:sp>.*?<\/p:sp>/)![0];
  const child = shape.replace('id="2" name="Text"', 'id="4" name="Child"');
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="5143500"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;
  zip.file(path, original.replace('</p:spTree>', `${group}${shape.replace('id="2" name="Text"', 'id="5" name="Front"')}</p:spTree>`));
  let context: PptxPluginContext<null> | undefined;
  const plugin = definePptxPlugin({ id: 'geometry', createState: () => null, onEvent(next) { context = next; } });
  const value = await open({ file: await zip.generateAsync({ type: 'uint8array' }), plugins: [plugin], pluginGrants: { geometry: { document: 'read' } } });
  const snapshot = await (await value.api.handleAsync()).snapshot();
  const slide = snapshot.slides[0]!;
  const back = slide.shapes.find((shape) => shape.name === 'Text')!;
  const grouped = slide.shapes.find((shape) => shape.name === 'Group')!;
  await waitFor(() => expect(context?.geometry?.getShapeRect(grouped.id)).toBeTruthy());
  expect(context!.geometry!.getShapeRect(grouped.id)).toEqual(context!.geometry!.getShapeRect(grouped.children[0]!.id));
  for (const shapeId of [back.id, grouped.id]) {
    await act(async () => {
      expect(await context!.navigation.selectShape({ slideId: slide.id, shapeId }, { expectVersion: await value.api.version() })).toMatchObject({ ok: true });
    });
    expect(value.client.state.selection).toMatchObject({ kind: 'object', shapeId });
    await waitFor(() => expect(context?.geometry?.getShapeRect(shapeId)).toBeTruthy());
  }
  const oldGeometry = context!.geometry!;
  await act(async () => { await value.api.goToSlide(2); });
  expect(oldGeometry.getShapeRect(back.id)).toBeNull();
});

test('plugin_lifecycle_state_commands_and_ordered_edits', async () => {
  let context: PptxPluginContext<{ count: number }> | undefined;
  const events: string[] = [];
  const cleanup: string[] = [];
  const plugin = definePptxPlugin({
    id: 'notes', createState: () => ({ count: 0 }),
    onEvent(next, event) {
      context = next;
      events.push(event.type);
      if (event.type === 'load') next.onCleanup((reason) => { cleanup.push(reason); });
    },
    panel: { title: 'Notes plugin', placement: 'bottom', render: ({ context }) => <div>Count {context.state.count}</div> },
    commands: [{ id: 'mark', label: 'Mark', mutatesDocument: true, shortcuts: ['Mod+Shift+K'],
      async execute(action) {
        const read = await action.read.readContent();
        if (!read.ok) return read;
        const result = await action.edits!.applyEdits({ expectVersion: read.version,
          steps: [{ op: 'setSlideNotes', target: { slideId: read.slides[0]!.id }, text: 'Plugin note' }] });
        if (!result.ok) return result;
        action.setState({ count: action.state.count + 1 }, result.version);
        return { ok: true, status: 'executed' };
      } }], toolbar: ['mark'],
  });
  const options = { plugins: [plugin], pluginGrants: { notes: { document: 'write' as const, editBatches: true as const } } };
  const value = await open(options);
  await waitFor(() => expect(context).toBeDefined());
  await act(async () => { expect(await value.api.commands!.execute('plugin:notes/mark', null)).toMatchObject({ ok: true }); });
  await waitFor(() => expect(value.view.getByText('Count 1')).toBeDefined());
  expect(events).toContain('document-change');
  expect(value.commands.filter((command) => command.kind === 'applyEdits')).toHaveLength(1);
  const read = await context!.read.readContent();
  expect(read.ok && read.slides[0]!.notes).toBe('Plugin note');
  act(() => { value.stage.focus(); fireEvent.keyDown(value.stage, { key: 'k', ctrlKey: true, shiftKey: true }); });
  await waitFor(() => expect(value.view.getByText('Count 2')).toBeDefined());
  const old = context!;
  options.plugins = [];
  await act(async () => { value.view.rerender(value.ui()); });
  expect(cleanup).toEqual(['removed']);
  expect(old.lifetimeSignal.aborted).toBe(true);
  expect(await old.read.readContent()).toMatchObject({ ok: false, failure: { code: 'plugin-unavailable' } });
  expect(await value.api.commands!.execute('plugin:notes/mark', null)).toMatchObject({ ok: false });
});

test('plugin_read_only_and_document_replacement_refuse_pending_edits', async () => {
  let context: PptxPluginContext<null> | undefined;
  const cleanups: string[] = [];
  const plugin = definePptxPlugin({ id: 'writer', createState: () => null,
    onEvent(next, event) {
      context = next;
      if (event.type === 'load') next.onCleanup((reason) => { cleanups.push(reason); });
    } });
  const options = { plugins: [plugin], pluginGrants: { writer: { document: 'write' as const, editBatches: true as const } } };
  const value = await open(options);
  await waitFor(() => expect(context).toBeDefined());
  const request = { expectVersion: await value.api.version(), steps: [{ op: 'setSlideNotes' as const,
    target: { slideId: value.client.state.slides[0]!.id }, text: 'Blocked' }] };
  await act(async () => { value.view.rerender(value.ui(true)); });
  await value.api.whenHydrated();
  expect(await context!.edits!.applyEdits(request)).toMatchObject({ ok: false, failure: { code: 'read-only' } });
  expect(await context!.read.validateEdits(request)).toMatchObject({ ok: false, failure: { code: 'read-only' } });
  const invocation = context!;
  replyControl.hold();
  let pending!: Promise<boolean>;
  act(() => { pending = value.api.goToSlide(2); });
  await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
  const work = invocation.edits!.applyEdits(request);
  work.catch(() => {});
  await act(async () => { value.view.rerender(<div>Closed</div>); replyControl.release(); await Promise.allSettled([pending]); });
  expect(await work).toMatchObject({ ok: false });
  expect(cleanups).toEqual(['unmounted']);
});

test('plugin_document_change_hook_can_commit_and_publish_state', async () => {
  const signals: boolean[] = [];
  const stateWrites: boolean[] = [];
  let hooks = 0;
  const plugin = definePptxPlugin({ id: 'hook', createState: () => ({ count: 0 }),
    async onEvent(context, event) {
      if (event.type !== 'document-change') return;
      const read = await context.read.readContent();
      if (!read.ok) return;
      if (read.slides[0]!.notes === 'Hook note') return;
      hooks++;
      const result = await context.edits!.applyEdits({ expectVersion: read.version,
        steps: [{ op: 'setSlideNotes', target: { slideId: read.slides[0]!.id }, text: 'Hook note' }] });
      if (!result.ok) return;
      signals.push(context.signal.aborted);
      stateWrites.push(context.setState({ count: 1 }, result.version));
    },
    panel: { title: 'Hook', placement: 'right', render: ({ context }) => <div>Hook count {context.state.count}</div> },
  });
  const value = await open({ plugins: [plugin], pluginGrants: { hook: { document: 'write', editBatches: true } } });
  const read = await value.api.readContent();
  if (!read.ok) throw new Error(read.failure.message);
  await act(async () => { await value.api.applyEdits({ expectVersion: read.version,
    steps: [{ op: 'setSlideNotes', target: { slideId: read.slides[0]!.id }, text: 'Host note' }] }); });
  await waitFor(() => expect(signals.length).toBeGreaterThan(0));
  expect(signals).toEqual([false]);
  expect(stateWrites).toEqual([true]);
  expect(hooks).toBe(1);
  expect(value.view.getByText('Hook count 1')).toBeDefined();
});

test('plugin_removal_refuses_before_pending_input_finishes', async () => {
  let context: PptxPluginContext<null> | undefined;
  const plugin = definePptxPlugin({ id: 'pending', createState: () => null, onEvent(next) { context = next; } });
  const options = { plugins: [plugin], pluginGrants: { pending: { document: 'write' as const, editBatches: true as const } } };
  const value = await open(options);
  await waitFor(() => expect(context).toBeDefined());
  const request = { expectVersion: await value.api.version(), steps: [{ op: 'setSlideNotes' as const,
    target: { slideId: value.client.state.slides[0]!.id }, text: 'Removed' }] };
  replyControl.hold();
  let navigation!: Promise<boolean>;
  act(() => { navigation = value.api.goToSlide(2); });
  await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
  let result: unknown;
  const work = context!.edits!.applyEdits(request).then((value) => { result = value; });
  options.plugins = [];
  await act(async () => { value.view.rerender(value.ui()); });
  await waitFor(() => expect(result).toMatchObject({ ok: false, failure: { code: 'plugin-unavailable' } }));
  expect(value.commands.filter((command) => command.kind === 'applyEdits')).toHaveLength(0);
  await act(async () => { replyControl.release(); await Promise.allSettled([navigation, work]); });
});

test('plugin_document_replacement_starts_a_new_generation', async () => {
  let context: PptxPluginContext<null> | undefined;
  const cleanups: string[] = [];
  const plugin = definePptxPlugin({ id: 'generation', createState: () => null,
    onEvent(next, event) {
      context = next;
      if (event.type === 'load') next.onCleanup((reason) => { cleanups.push(reason); });
    } });
  const options: Partial<PptxAsyncEditorProps> = { plugins: [plugin] };
  const value = await open(options);
  await waitFor(() => expect(context).toBeDefined());
  const previous = context!;
  options.file = await deck();
  await act(async () => { value.view.rerender(value.ui()); });
  await waitFor(() => expect(context!.snapshot.generation).not.toBe(previous.snapshot.generation));
  expect(previous.lifetimeSignal.aborted).toBe(true);
  expect(await previous.read.readContent()).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(cleanups).toEqual(['document-replaced']);
  await act(async () => { value.view.unmount(); });
  expect(cleanups).toEqual(['document-replaced', 'unmounted']);
});

test('peerless_plugin_reads_preserve_engine_facts', async () => {
  const { readFactsDeck } = await import('../plugins/readFacts.testSupport');
  let context: PptxPluginContext<null> | undefined;
  const plugin = definePptxPlugin({ id: 'facts', createState: () => null, onEvent(next) { context = next; } });
  const value = await open({ file: await readFactsDeck(), plugins: [plugin] });
  await waitFor(() => expect(context).toBeDefined());
  const result = await context!.read.readContent();
  expect(result).toEqual(await value.api.readContent());
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error('Read refused');
  expect(result.slides[0]).toMatchObject({ title: 'Heading\nNext', hidden: true, layout: { name: 'Heading layout', type: 'title' } });
  expect(result.stories[0]!.paragraphs.map((paragraph) => paragraph.list)).toEqual([
    { kind: 'number', scheme: 'alphaLcParenR', startAt: 1, value: 1, marker: 'a)' },
    { kind: 'number', scheme: 'alphaLcParenR', startAt: 1, value: 2, marker: 'b)' },
  ]);
});
