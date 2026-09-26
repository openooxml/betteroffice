import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { StrictMode, useState } from 'react';
import { createPortal } from 'react-dom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  initWasm,
  openPresentation,
  type PptxEditRequest,
  type PptxEditResult,
} from '@betteroffice/pptx';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import * as publicApi from '../index';
import {
  EditorToolbar,
  PptxEditor,
  PptxPluginToolbar,
  ToolbarCommandButton,
  definePptxPlugin,
  usePptxCommand,
  usePptxCommands,
  type PptxCommandStore,
  type PptxEditorApi,
  type PptxEditorProps,
  type PptxPlugin,
  type PptxPluginContext,
  type PptxPluginDefinition,
  type PptxPluginError,
  type PptxPluginEvent,
  type PptxPluginGeometry,
  type PptxPluginRefusal,
} from '../index';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

const root = resolve(import.meta.dir, '../../../..');
const quiet = { error: console.error, warn: console.warn };
let fixture: Uint8Array;
let faces: { family: string; bytes: Uint8Array }[];

beforeAll(async () => {
  const [wasm, deck, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = new Uint8Array(deck);
  faces = [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }];
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  await new Promise((done) => setTimeout(done, 50));
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

function last<T>(values: readonly T[]): T {
  return values[values.length - 1];
}

async function settle(ms = 20) {
  await act(async () => {
    await new Promise((done) => setTimeout(done, ms));
  });
}

async function until(done: () => boolean) {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) await settle(10);
  expect(done()).toBe(true);
}

/** A 2D context that accepts every call, so frames paint and present without pixels. */
function fakeContexts() {
  const contexts = new WeakMap<HTMLCanvasElement, { context: unknown; calls: string[] }>();
  const gradient = { addColorStop() {} };
  const contextOf = (canvas: HTMLCanvasElement) => {
    let entry = contexts.get(canvas);
    if (entry) return entry;
    const calls: string[] = [];
    const values: Record<string, unknown> = {
      canvas,
      measureText: (text: string) => ({
        width: String(text).length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      }),
      createLinearGradient: () => gradient,
      createRadialGradient: () => gradient,
      createPattern: () => null,
      getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
      getImageData: (_x: number, _y: number, width: number, height: number) => ({
        data: new Uint8ClampedArray(Math.max(4, width * height * 4)),
        width,
        height,
      }),
    };
    const context = new Proxy(values, {
      get(target, key) {
        if (typeof key !== 'string') return undefined;
        if (key in target) return target[key];
        return () => {
          calls.push(key);
        };
      },
      set(target, key, value) {
        target[key as string] = value;
        return true;
      },
    });
    entry = { context, calls };
    contexts.set(canvas, entry);
    return entry;
  };
  const spy = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (
    this: HTMLCanvasElement
  ) {
    return contextOf(this).context as never;
  });
  return {
    calls: (canvas: HTMLCanvasElement) => contextOf(canvas).calls,
    restore: () => spy.mockRestore(),
  };
}

type State = { count: number };

function describeEvent(event: PptxPluginEvent): string | null {
  switch (event.type) {
    case 'load':
      return `load:${event.reason}`;
    case 'document-change':
      return `document-change:${event.version}`;
    case 'mode-change':
      return `mode-change:${event.readOnly}`;
    case 'grants-change':
      return 'grants-change';
    default:
      return null;
  }
}

function recorder(id = 'acme.review', extra: Partial<PptxPluginDefinition<State>> = {}) {
  const log: string[] = [];
  const events: PptxPluginEvent[] = [];
  const contexts: PptxPluginContext<State>[] = [];
  const plugin = definePptxPlugin<State>({
    id,
    createState: () => ({ count: 0 }),
    initialize(context) {
      contexts.push(context);
      log.push('initialize');
      context.onCleanup((reason) => {
        log.push(`cleanup:${reason}`);
      });
    },
    onEvent(context, event) {
      contexts.push(context);
      events.push(event);
      const entry = describeEvent(event);
      if (entry) log.push(entry);
    },
    ...extra,
  });
  return { plugin, log, events, contexts };
}

async function mount(props: Partial<PptxEditorProps> = {}, strict = false) {
  const ready: PptxEditorApi[] = [];
  const element = (next: Partial<PptxEditorProps>) => {
    const editor = (
      <PptxEditor file={fixture} fonts={faces} onReady={(api) => ready.push(api)} {...next} />
    );
    return strict ? <StrictMode>{editor}</StrictMode> : editor;
  };
  const view = render(element(props));
  await until(() => ready.length > 0);
  return {
    api: () => last(ready),
    ready,
    view,
    rerender: (next: Partial<PptxEditorProps>) => view.rerender(element(next)),
  };
}

function notesRequest(version: string, slideId: string, text: string): PptxEditRequest {
  return { expectVersion: version, steps: [{ op: 'setSlideNotes', target: { slideId }, text }] };
}

async function read(api: PptxEditorApi) {
  const result = await api.readContent();
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

const WRITE = { 'acme.review': { document: 'write', editBatches: true } } as const;

/** Image decoding for inserted pictures that finishes only when the test says so. */
function pauseInsertedImages() {
  const originalImage = globalThis.Image;
  const pending: Array<() => void> = [];
  class PausedImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    naturalWidth = 400;
    naturalHeight = 200;
    set src(value: string) {
      if (value.startsWith('data:')) pending.push(() => this.onload?.());
      else queueMicrotask(() => this.onload?.());
    }
  }
  globalThis.Image = PausedImage as unknown as typeof Image;
  return {
    pending,
    finish: () => {
      for (const finish of pending.splice(0)) finish();
    },
    restore: () => {
      globalThis.Image = originalImage;
    },
  };
}

function insertImage(view: ReturnType<typeof render>) {
  const file = new File([Uint8Array.from([0x89, 0x50, 0x4e, 0x47])], 'queued.png', {
    type: 'image/png',
  });
  fireEvent.change(view.getByTestId('pptx-insert-image-input'), { target: { files: [file] } });
}

async function storyText(api: PptxEditorApi, storyId: string): Promise<string> {
  return (await read(api)).stories.find((story) => story.storyId === storyId)!.text;
}

describe('PptxEditor plugins', () => {
  test('initialize, load, one change per committed change, and cleanup on removal', async () => {
    const { plugin, log, contexts } = recorder();
    const { api, rerender } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    expect(log.slice(0, 2)).toEqual(['initialize', 'load:loaded']);
    const content = await read(api());
    expect(contexts[1].snapshot.version).toBe(content.version);
    const slideId = content.slides[0].id;
    const changes = () => log.filter((entry) => entry.startsWith('document-change'));

    let applied!: Awaited<ReturnType<PptxEditorApi['applyEdits']>>;
    await act(async () => {
      applied = await api().applyEdits(notesRequest(content.version, slideId, 'Plugin notes'));
    });
    expect(applied).toMatchObject({ ok: true, applied: true });
    await settle(50);
    expect(changes()).toEqual([`document-change:${applied.ok ? applied.version : ''}`]);

    await act(async () => {
      const stale = await api().applyEdits(notesRequest(content.version, slideId, 'Late'));
      expect(stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    });
    const current = await read(api());
    await act(async () => {
      const noop = await api().applyEdits(notesRequest(current.version, slideId, 'Plugin notes'));
      expect(noop).toMatchObject({ ok: true, applied: false });
    });
    await settle(50);
    expect(changes()).toHaveLength(1);

    for (const step of ['undo', 'redo'] as const) {
      await act(async () => {
        expect(await api().commands.execute(step, null)).toMatchObject({ ok: true });
      });
      await settle(50);
      expect(last(changes())).toBe(`document-change:${api().handle.version()}`);
    }
    expect(changes()).toHaveLength(3);

    const replica = openPresentation(fixture, {
      clientId: 7171,
      initialUpdate: api().handle.encodeStateAsUpdate(),
    });
    try {
      replica.setSlideNotes(slideId, 'Remote notes');
      await act(async () => {
        api().handle.applyUpdate(replica.encodeDiff(api().handle.encodeStateVector()));
      });
      await settle(50);
    } finally {
      replica.dispose();
    }
    expect(changes()).toHaveLength(4);
    expect(last(changes())).toBe(`document-change:${api().handle.version()}`);

    rerender({ plugins: [] });
    await settle();
    expect(log.filter((entry) => entry.startsWith('cleanup'))).toEqual(['cleanup:removed']);
    const removed = contexts[0];
    expect(removed.lifetimeSignal.aborted).toBe(true);
    expect(removed.setState({ count: 1 })).toBe(false);
    expect(await removed.read.version()).toMatchObject({
      ok: false,
      failure: { code: 'plugin-unavailable' },
    });

    rerender({ plugins: [plugin] });
    await until(() => log.includes('load:attached'));
  });

  test('a default plugin reads and navigates but cannot mutate through any path', async () => {
    const calls: string[] = [];
    function Probe() {
      const commands = usePptxCommands();
      const bold = commands.getState('bold');
      return (
        <>
          <button
            type="button"
            data-testid="probe"
            data-bold={bold.enabled ? 'enabled' : bold.disabledReason.code}
            onClick={() =>
              void commands.execute('bold', null).then((result) => {
                calls.push(result.ok ? 'ok' : result.failure.code);
              })
            }
          />
          <ToolbarCommandButton id="italic" />
        </>
      );
    }
    const { plugin, contexts, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'right', render: Probe },
    });
    const { api, view } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    const content = await read(api());
    const story = content.stories[0];
    await act(async () => {
      api().selectText({
        slide: 1,
        shapeId: story.shapeId,
        storyId: story.storyId,
        start: 0,
        end: 2,
      });
    });
    await settle();
    const probe = within(view.container).getByTestId('probe');
    expect(probe.getAttribute('data-bold')).toBe('permission-denied');
    fireEvent.click(probe);
    await settle();
    expect(calls).toEqual(['permission-denied']);
    const italic = within(view.getByTestId('plugin-dock-right')).getByRole('button', {
      name: 'Italic',
    });
    expect(italic.getAttribute('aria-disabled')).toBe('true');
    expect(document.getElementById(italic.getAttribute('aria-describedby')!)?.textContent).toBe(
      'The host application has not allowed this plugin to use this command.'
    );
    expect(api().commands.getState('italic').enabled).toBe(true);

    const context = last(contexts);
    expect(context.edits).toBeNull();
    expect(await context.commands.execute('bold', null)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
    const reread = await context.read.readContent();
    expect(reread.ok).toBe(true);
    if (!reread.ok) return;
    expect(
      await context.read.findText({
        text: story.text.slice(0, 3),
        within: { slideId: story.slideId },
      })
    ).toMatchObject({ ok: true, version: reread.version });
    const request = notesRequest(reread.version, story.slideId, 'Checked');
    expect(await context.read.validateEdits(request)).toMatchObject({ ok: true, wouldApply: true });
    expect(
      await context.navigation.goToSlide(
        { slideId: 'slide:missing' },
        { expectVersion: reread.version }
      )
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    expect(
      await context.navigation.goToSlide({ slideId: story.slideId }, { expectVersion: 'stale' })
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect((await read(api())).version).toBe(reread.version);
  });

  test('granted built-in mutations refuse without a policy path; granted views run', async () => {
    function Probe() {
      const bold = usePptxCommands().getState('bold');
      return (
        <output data-testid="probe">{bold.enabled ? 'enabled' : bold.disabledReason.code}</output>
      );
    }
    const { plugin, contexts, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'left', render: Probe },
    });
    const { api, view } = await mount({
      plugins: [plugin],
      pluginGrants: { 'acme.review': { document: 'write', commands: ['bold', 'zoom'] } },
    });
    await until(() => log.includes('load:loaded'));
    const story = (await read(api())).stories[0];
    await act(async () => {
      api().selectText({
        slide: 1,
        shapeId: story.shapeId,
        storyId: story.storyId,
        start: 0,
        end: 2,
      });
    });
    await settle();
    expect(within(view.container).getByTestId('probe').textContent).toBe('unsupported-policy');
    const context = last(contexts);
    expect(await context.commands.execute('bold', null)).toMatchObject({
      ok: false,
      failure: { code: 'unsupported-policy' },
    });
    expect(await act(() => context.commands.execute('zoom', { scale: 2 }))).toEqual({
      ok: true,
      status: 'executed',
    });
    expect(api().commands.getState('zoom').value).toBe(2);
    expect(api().commands.getState('bold').enabled).toBe(true);
  });

  test('edit batches need their grant, wait for pending input and respect readOnly', async () => {
    const originalImage = globalThis.Image;
    const pending: Array<() => void> = [];
    class PausedImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      naturalWidth = 400;
      naturalHeight = 200;
      set src(value: string) {
        if (value.startsWith('data:')) pending.push(() => this.onload?.());
        else queueMicrotask(() => this.onload?.());
      }
    }
    globalThis.Image = PausedImage as unknown as typeof Image;
    try {
      const { plugin, contexts, log } = recorder();
      const { api, view, rerender } = await mount({ plugins: [plugin], pluginGrants: WRITE });
      await until(() => log.includes('load:loaded'));
      const edits = last(contexts).edits!;
      expect(edits).not.toBeNull();
      const before = await read(api());
      const slideId = before.slides[1].id;

      const file = new File([Uint8Array.from([0x89, 0x50, 0x4e, 0x47])], 'queued.png', {
        type: 'image/png',
      });
      fireEvent.change(view.getByTestId('pptx-insert-image-input'), { target: { files: [file] } });
      await until(() => pending.length > 0);
      let settled = false;
      const late = edits
        .applyEdits(notesRequest(before.version, slideId, 'Too late'))
        .finally(() => {
          settled = true;
        });
      await settle();
      expect(settled).toBe(false);
      await act(async () => {
        for (const finish of pending.splice(0)) finish();
      });
      expect(await late).toMatchObject({ ok: false, failure: { code: 'stale-version' } });

      const current = await read(api());
      const applied = await act(() =>
        edits.applyEdits(notesRequest(current.version, slideId, 'Plugin notes'))
      );
      expect(applied).toMatchObject({ ok: true, applied: true });
      expect((await read(api())).slides[1].notes).toBe('Plugin notes');
      await act(async () => {
        await api().commands.execute('undo', null);
      });
      expect((await read(api())).slides[1].notes ?? '').not.toBe('Plugin notes');
      const version = (await read(api())).version;
      expect(
        await edits.applyEdits({ ...notesRequest(version, slideId, 'Hidden'), history: 'none' })
      ).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });

      rerender({ plugins: [plugin], pluginGrants: WRITE, readOnly: true });
      await until(() => log.includes('mode-change:true'));
      expect(await edits.applyEdits(notesRequest(version, slideId, 'Read-only'))).toMatchObject({
        ok: false,
        version,
        failure: { code: 'read-only' },
      });
      expect(
        await last(contexts).read.validateEdits(notesRequest(version, slideId, 'x'))
      ).toMatchObject({
        ok: false,
        failure: { code: 'read-only' },
      });

      rerender({ plugins: [plugin], pluginGrants: {}, readOnly: false });
      await until(() => log.includes('grants-change'));
      expect(await edits.applyEdits(notesRequest(version, slideId, 'Revoked'))).toMatchObject({
        ok: false,
        failure: { code: 'permission-denied' },
      });
      expect((await read(api())).version).toBe(version);

      const grant: { document: 'write'; editBatches?: true } = {
        document: 'write',
        editBatches: true,
      };
      const grants = { 'acme.review': grant };
      const changes = () => log.filter((entry) => entry === 'grants-change').length;
      rerender({ plugins: [plugin], pluginGrants: grants });
      await until(() => changes() === 2);
      delete grant.editBatches;
      rerender({ plugins: [plugin], pluginGrants: grants });
      await until(() => changes() === 3);
      expect(await edits.applyEdits(notesRequest(version, slideId, 'In place'))).toMatchObject({
        ok: false,
        failure: { code: 'permission-denied' },
      });
    } finally {
      globalThis.Image = originalImage;
    }
  });

  test('contributed commands run with their plugin, from the toolbar, API and shortcuts', async () => {
    const calls: string[] = [];
    const errors: PptxPluginError[] = [];
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+Shift+M', 'Mod+B', 'Mod+V'],
          getState: (context) => ({ enabled: true, active: context.state.count > 0 }),
          execute(context) {
            calls.push(context.pluginId);
            context.setState((previous) => ({ count: previous.count + 1 }));
            return { ok: true, status: 'executed' };
          },
        },
      ],
      toolbar: ['mark'],
    });
    const other = recorder('other');
    const { api, view } = await mount({
      plugins: [owner.plugin, other.plugin],
      onPluginError: (error) => errors.push(error),
    });
    await until(() => owner.log.includes('load:loaded') && other.log.includes('load:loaded'));
    await settle();
    const id = 'plugin:acme.review/mark' as const;
    expect(api().commands.getDescriptor(id)).toMatchObject({ label: 'Mark' });
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([
      ['acme.review', 'definition'],
      ['acme.review', 'definition'],
    ]);
    expect(String(errors[1].error)).toContain('already used by text editing');

    const more = view.getByTestId('pptx-toolbar-more');
    act(() => more.focus());
    fireEvent.keyDown(more, { key: 'ArrowDown' });
    const menu = view.getByRole('menu', { name: 'More' });
    const entry = within(menu).getByRole('menuitemcheckbox', { name: /^Mark/ });
    expect(entry.getAttribute('aria-checked')).toBe('false');
    fireEvent.click(entry);
    await settle();
    expect(calls).toEqual(['acme.review']);
    expect(api().commands.getState(id).active).toBe(true);

    expect(await act(() => api().commands.execute(id, null))).toMatchObject({ ok: true });
    fireEvent.keyDown(view.getByRole('application'), {
      key: 'M',
      shiftKey: true,
      ctrlKey: true,
      metaKey: true,
    });
    await settle();
    expect(calls).toEqual(['acme.review', 'acme.review', 'acme.review']);

    expect(await last(other.contexts).commands.execute(id, null)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });

    cleanup();
    const custom = await mount({
      plugins: [owner.plugin],
      toolbar: (
        <EditorToolbar mode="commands">
          <PptxPluginToolbar />
        </EditorToolbar>
      ),
    });
    await until(
      () => within(custom.view.container).queryByRole('button', { name: 'Mark' }) !== null
    );
  });

  test('a failing contribution is isolated and the others keep working', async () => {
    const errors: PptxPluginError[] = [];
    const broken = definePptxPlugin({
      id: 'broken',
      createState: () => null,
      panel: {
        title: 'Broken',
        placement: 'left',
        render: () => {
          throw new Error('render failed');
        },
      },
      commands: [
        {
          id: 'go',
          label: 'Go',
          mutatesDocument: false,
          execute: () => ({ ok: true, status: 'executed' }),
        },
      ],
      toolbar: ['go'],
    });
    const cleanups: string[] = [];
    const throwing = definePptxPlugin({
      id: 'throws',
      createState: () => null,
      initialize(context) {
        context.onCleanup((reason) => {
          cleanups.push(reason);
          throw new Error('cleanup failed');
        });
      },
      onEvent(_context, event) {
        if (event.type === 'load') throw new Error('load failed');
      },
    });
    const rejectsAction = definePptxPlugin({
      id: 'rejects-action',
      createState: () => null,
      panel: {
        title: 'Action',
        placement: 'bottom',
        render: ({ context }) => (
          <button
            type="button"
            data-testid="failing-action"
            onClick={() =>
              void context.run(async () => {
                throw new Error('action rejected');
              })
            }
          />
        ),
      },
    });
    const healthy = recorder('healthy', {
      panel: { title: 'Healthy', placement: 'left', render: () => <p data-testid="healthy">ok</p> },
    });
    const { api, view } = await mount({
      plugins: [broken, throwing, rejectsAction, healthy.plugin],
      onPluginError: (error) => {
        errors.push(error);
        throw new Error('reporter failed');
      },
    });
    await until(() =>
      errors.some((error) => error.pluginId === 'broken' && error.phase === 'render')
    );
    await until(() =>
      errors.some((error) => error.pluginId === 'throws' && error.phase === 'cleanup')
    );
    expect(errors.find((error) => error.pluginId === 'throws')?.phase).toBe('event');
    expect(cleanups).toEqual(['failed']);
    await until(() => within(view.container).queryByTestId('failing-action') !== null);
    fireEvent.click(within(view.container).getByTestId('failing-action'));
    await until(() =>
      errors.some((error) => error.pluginId === 'rejects-action' && error.phase === 'action')
    );
    await settle();
    expect(within(view.container).queryByTestId('failing-action')).toBeNull();
    expect(within(view.container).getByTestId('healthy').textContent).toBe('ok');
    expect(api().commands.getDescriptor('plugin:broken/go')).toBeNull();
    expect(api().commands.getState('save').enabled).toBe(true);
  });

  test('replacing the presentation ends the activation and loads a fresh one', async () => {
    const { plugin, log, contexts } = recorder();
    const { rerender } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    const first = contexts[0];
    rerender({ plugins: [plugin], file: fixture.slice() });
    await until(() => log.includes('load:replaced'));
    expect(log.indexOf('cleanup:document-replaced')).toBeLessThan(log.indexOf('load:replaced'));
    expect(first.lifetimeSignal.aborted).toBe(true);
    expect(await first.read.version()).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    const fresh = last(contexts);
    expect(fresh.snapshot.generation).not.toBe(first.snapshot.generation);
    expect(await fresh.read.version()).toMatchObject({ ok: true });
  });

  test('a presentation replaced as it opens activates plugins once, over the new one', async () => {
    const { plugin, log, contexts } = recorder();
    const plugins = [plugin];
    const errors: PptxPluginError[] = [];
    let replace = true;
    function Host() {
      const [file, setFile] = useState(fixture);
      return (
        <PptxEditor
          file={file}
          fonts={faces}
          plugins={plugins}
          onPluginError={(error) => errors.push(error)}
          onReady={() => {
            if (replace) setFile(fixture.slice());
            replace = false;
          }}
        />
      );
    }
    render(<Host />);
    await until(() => log.includes('load:loaded'));
    await settle(150);
    expect(log).toEqual(['initialize', 'load:loaded']);
    expect(await last(contexts).read.version()).toMatchObject({ ok: true });
    expect(errors).toEqual([]);
  });

  test('StrictMode setup, cleanup and setup leaves one live activation', async () => {
    const { plugin, log, contexts } = recorder();
    await mount({ plugins: [plugin] }, true);
    await until(() => log.includes('load:loaded'));
    await settle();
    const initialized = log.filter((entry) => entry === 'initialize').length;
    const cleaned = log.filter((entry) => entry.startsWith('cleanup')).length;
    expect(initialized - cleaned).toBe(1);
    const live = new Set(
      contexts
        .filter((context) => !context.lifetimeSignal.aborted)
        .map((context) => context.lifetimeSignal)
    );
    expect(live.size).toBe(1);
  });

  test('navigation resolves session ids after input, keeps focus and reports selections', async () => {
    const { plugin, events, contexts, log } = recorder('acme.review', {
      panel: {
        title: 'Review',
        placement: 'right',
        render: () => <input data-testid="panel-input" aria-label="Note" />,
      },
    });
    const { api, view } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    const context = () => last(contexts);
    const content = await read(api());
    const second = content.slides[1];
    const stage = view.getByRole('application');
    const input = view.getByTestId('panel-input');
    input.focus();
    const options = { expectVersion: content.version };

    expect(
      await act(() => context().navigation.goToSlide({ slideId: second.id }, options))
    ).toEqual({
      ok: true,
    });
    const thumbnails = within(view.getByRole('complementary', { name: /slides/i })).getAllByRole(
      'button'
    );
    expect(thumbnails[1].getAttribute('aria-current')).toBe('page');
    expect(document.activeElement).toBe(input);
    await until(() =>
      events.some(
        (event) =>
          event.type === 'selection-change' &&
          event.selection?.slideId === second.id &&
          event.selection.target.kind === 'slide'
      )
    );

    const shape = second.shapes.find((candidate) => candidate.textStories.length > 0)!;
    expect(
      await act(() =>
        context().navigation.selectShape({ slideId: second.id, shapeId: shape.id }, options)
      )
    ).toEqual({ ok: true });
    await until(() => context().snapshot.selection?.target.kind === 'shape');

    const story = content.stories.find((candidate) => candidate.shapeId === shape.id)!;
    expect(
      await act(() =>
        context().navigation.selectText(
          { slideId: second.id, shapeId: shape.id, storyId: story.storyId, anchor: 4, focus: 1 },
          { ...options, focus: true }
        )
      )
    ).toEqual({ ok: true });
    await until(() => context().snapshot.selection?.target.kind === 'text');
    expect(context().snapshot.selection).toEqual({
      slideId: second.id,
      slide: 2,
      target: { kind: 'text', shapeId: shape.id, storyId: story.storyId, anchor: 4, focus: 1 },
    });
    expect(document.activeElement).toBe(stage);

    expect(
      await context().navigation.selectShape({ slideId: second.id, shapeId: 'gone' }, options)
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    expect(
      await context().navigation.selectText(
        { slideId: second.id, shapeId: shape.id, storyId: story.storyId, anchor: 0, focus: 1e6 },
        options
      )
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    expect(
      await context().navigation.selectShape(
        { slideId: content.slides[0].id, shapeId: shape.id },
        options
      )
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });

    await act(async () => {
      await api().applyEdits(notesRequest(content.version, second.id, 'Moved on'));
    });
    expect(
      await context().navigation.goToSlide({ slideId: content.slides[0].id }, options)
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(thumbnails[1].getAttribute('aria-current')).toBe('page');
  });

  test('a batch the load hook applies reports its receipt, and load repeats at that version', async () => {
    const outcomes: (PptxEditResult | PptxPluginRefusal)[] = [];
    const seen: string[] = [];
    const plugin = definePptxPlugin<null>({
      id: 'acme.review',
      createState: () => null,
      async onEvent(context, event) {
        if (event.type === 'load' || event.type === 'document-change') {
          seen.push(`${event.type}:${event.version}`);
        }
        if (event.type !== 'load' || !context.edits) return;
        const content = await context.read.readContent();
        if (!content.ok || content.slides[0].notes === 'From load') return;
        outcomes.push(
          await context.edits.applyEdits(
            notesRequest(content.version, content.slides[0].id, 'From load')
          )
        );
      },
    });
    const { api } = await mount({ plugins: [plugin], pluginGrants: WRITE });
    await until(() => outcomes.length > 0);
    const after = await read(api());
    expect(outcomes).toEqual([
      expect.objectContaining({ ok: true, applied: true, version: after.version }),
    ]);
    await until(() => seen.includes(`load:${after.version}`));
    await settle(50);
    const { baseVersion } = outcomes[0] as { baseVersion: string };
    expect(seen).toEqual([`load:${baseVersion}`, `load:${after.version}`]);
  });

  test('layouts follow the presented frame, its version, slide and zoom', async () => {
    const canvases = fakeContexts();
    try {
      const layouts: (string | null)[] = [];
      const contexts: PptxPluginContext<null>[] = [];
      const plugin = definePptxPlugin<null>({
        id: 'acme.review',
        createState: () => null,
        onEvent(context, event) {
          contexts.push(context);
          if (event.type === 'layout-change') layouts.push(event.layout?.version ?? null);
        },
        overlay: ({ context, geometry }) => (
          <div
            data-testid="layout-marker"
            data-version={geometry.layout.version}
            data-snapshot={context.snapshot.version}
            data-slide={geometry.layout.slide}
            data-zoom={geometry.layout.zoom}
          />
        ),
      });
      const { api, view } = await mount({ plugins: [plugin] });
      const marker = () =>
        view.container.querySelector<HTMLElement>('[data-testid="layout-marker"]');
      await until(() => marker() !== null);
      const content = await read(api());
      expect(marker()!.dataset).toMatchObject({
        version: content.version,
        snapshot: content.version,
        slide: '1',
      });
      const layer = view.getByTestId('plugin-overlays');
      expect(layer.parentElement).toBe(view.getByTestId('pptx-slide-canvas').parentElement);
      expect(layer.style.pointerEvents).toBe('none');

      let next = '';
      await act(async () => {
        const applied = await api().applyEdits(
          notesRequest(content.version, content.slides[0].id, 'Changed')
        );
        if (applied.ok) next = applied.version;
      });
      await until(() => marker()?.dataset.version === next);
      const afterEdit = layouts.slice(layouts.lastIndexOf(content.version) + 1);
      expect(afterEdit[0]).toBeNull();
      expect(last(afterEdit)).toBe(next);

      await act(async () => {
        await api().commands.execute('zoom', { scale: 1.5 });
      });
      await until(() => marker()?.dataset.zoom === '1.5');
      const context = last(contexts);
      expect(context.geometry?.layout).toMatchObject({ version: next, zoom: 1.5, width: 1280 });
      expect(context.geometry?.getShapeRect('slide:0:256:shape:8:missing')).toBeNull();

      await act(async () => {
        await context.navigation.goToSlide(
          { slideId: content.slides[2].id },
          { expectVersion: next }
        );
      });
      await until(() => marker()?.dataset.slide === '3');
      expect(last(contexts).snapshot.layout?.slideId).toBe(content.slides[2].id);

      const story = content.stories.find(
        (candidate) => candidate.slideId === content.slides[2].id
      )!;
      await act(async () => {
        api().handle.propose('Agent', null, [
          { type: 'replaceText', storyId: story.storyId, start: 0, end: 0, text: 'Proposed ' },
        ]);
        api().refreshProposals();
      });
      await until(() => marker() === null);
      expect(last(contexts).geometry).toBeNull();
      expect(last(layouts)).toBeNull();
      await act(async () => {
        await api().commands.execute('proposalDiff', { enabled: false });
      });
      await until(() => marker()?.dataset.slide === '3');
    } finally {
      canvases.restore();
    }
  });

  for (const outcome of ['decodes', 'fails'] as const) {
    test(`a superseded paint never draws over the newer slide when its image ${outcome}`, async () => {
      const canvases = fakeContexts();
      const originalBitmap = globalThis.createImageBitmap;
      const pending: Array<() => void> = [];
      globalThis.createImageBitmap = (() =>
        new Promise((done, fail) =>
          pending.push(() =>
            outcome === 'decodes'
              ? done({ width: 4, height: 4 } as ImageBitmap)
              : fail(new Error('decode failed'))
          )
        )) as never;
      try {
        const contexts: PptxPluginContext<null>[] = [];
        const plugin = definePptxPlugin<null>({
          id: 'acme.review',
          createState: () => null,
          onEvent(context) {
            contexts.push(context);
          },
          overlay: ({ geometry }) => (
            <div data-testid="layout-marker" data-slide={geometry.layout.slide} />
          ),
        });
        const { api, view } = await mount({ plugins: [plugin] });
        await until(() => pending.length > 0 && contexts.length > 0);
        const marker = () =>
          view.container.querySelector<HTMLElement>('[data-testid="layout-marker"]');
        expect(marker()).toBeNull();
        const content = await read(api());
        await act(async () => {
          await last(contexts).navigation.goToSlide(
            { slideId: content.slides[1].id },
            { expectVersion: content.version }
          );
        });
        await until(() => marker()?.dataset.slide === '2');
        const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
        const drawn = canvases.calls(canvas).length;
        await act(async () => {
          for (const finish of pending.splice(0)) finish();
        });
        await settle(50);
        expect(canvases.calls(canvas).length).toBe(drawn);
        expect(marker()?.dataset.slide).toBe('2');
      } finally {
        globalThis.createImageBitmap = originalBitmap;
        canvases.restore();
      }
    });
  }

  test('public presenters and hooks bind contributed commands', async () => {
    const calls: string[] = [];
    function Bound() {
      const mark = usePptxCommand('plugin:acme.review/mark');
      return (
        <output data-testid="bound">{`${mark.label}:${mark.state.enabled}:${mark.shortcut}`}</output>
      );
    }
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+Shift+M'],
          execute(context) {
            calls.push(context.pluginId);
            return { ok: true, status: 'executed' };
          },
        },
      ],
    });
    const { view } = await mount({
      plugins: [owner.plugin],
      toolbar: (
        <EditorToolbar mode="commands">
          <ToolbarCommandButton id="plugin:acme.review/mark" />
          <ToolbarCommandButton id="plugin:acme.review/absent" label="Absent" />
          <Bound />
        </EditorToolbar>
      ),
    });
    const body = within(view.container);
    await until(() => body.queryByRole('button', { name: 'Mark' }) !== null);
    expect(body.queryByRole('button', { name: 'Absent' })).toBeNull();
    expect(body.getByTestId('bound').textContent).toMatch(/^Mark:true:.+M$/);
    fireEvent.click(body.getByRole('button', { name: 'Mark' }));
    await settle();
    expect(calls).toEqual(['acme.review']);
  });

  test('two editors keep separate activations and events', async () => {
    const { plugin, contexts, log } = recorder();
    const first = await mount({ plugins: [plugin] });
    const second = await mount({ plugins: [plugin] });
    await until(() => log.filter((entry) => entry === 'load:loaded').length === 2);
    const generations = new Set(contexts.map((context) => context.lifetimeSignal));
    expect(generations.size).toBe(2);
    const content = await read(first.api());
    await act(async () => {
      await first.api().applyEdits(notesRequest(content.version, content.slides[0].id, 'One'));
    });
    await settle(50);
    const changed = contexts.filter(
      (context) => context.snapshot.version === first.api().handle.version()
    );
    expect(new Set(changed.map((context) => context.lifetimeSignal)).size).toBe(1);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toHaveLength(1);
    expect(second.api().handle.version()).not.toBe(first.api().handle.version());
  });

  test('the demo review plugin works through the documented exports', async () => {
    mock.module('@betteroffice/pptx-react', () => publicApi);
    const demo = '../../../../apps/demo/app/pptx/ReviewPlugin';
    const { reviewPlugin } = (await import(demo)) as { reviewPlugin: PptxPlugin };
    const { api, view, rerender } = await mount({ plugins: [reviewPlugin] });
    const body = within(view.container);
    const action = () => body.getByRole('button', { name: 'Mark current slide reviewed' });
    await until(() => body.queryByText(/slides, version/) !== null);
    expect(action().hasAttribute('disabled')).toBe(true);
    const mark = 'plugin:demo.review/mark-reviewed' as const;
    expect(api().commands.getState(mark).disabledReason?.code).toBe('write-not-granted');
    expect(api().commands.getDescriptor(mark)?.label).toBe('Mark reviewed');

    const grants = { 'demo.review': { document: 'write', editBatches: true } } as const;
    rerender({ plugins: [reviewPlugin], pluginGrants: grants });
    await until(() => !action().hasAttribute('disabled'));
    fireEvent.click(action());
    await until(() => body.queryByText('Marked the current slide as reviewed.') !== null);
    const notes = (await read(api())).slides[0].notes ?? '';
    expect(notes.endsWith('Reviewed ✓')).toBe(true);

    rerender({ plugins: [reviewPlugin], pluginGrants: grants, readOnly: true });
    await until(() => action().hasAttribute('disabled'));

    rerender({ plugins: [], pluginGrants: grants });
    await until(() => body.queryByRole('button', { name: 'Mark current slide reviewed' }) === null);
  });

  test('plugins that are not defined through definePptxPlugin are reported, not run', async () => {
    const errors: PptxPluginError[] = [];
    await mount({
      plugins: [{ id: 'raw' } as unknown as PptxPlugin],
      onPluginError: (error) => errors.push(error),
    });
    await settle();
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([['raw', 'definition']]);
  });
});

describe('PptxEditor plugin boundaries and races', () => {
  /** An editor with an overlay input and a panel whose button and field also render in a portal. */
  async function mountKeyboardProbe(marks: string[]) {
    function Panel() {
      return (
        <>
          <button type="button" data-testid="panel-button">
            Panel
          </button>
          {createPortal(
            <>
              <input data-testid="portal-input" aria-label="Portal" />
              <button type="button" data-testid="portal-button">
                Portal
              </button>
            </>,
            document.body
          )}
        </>
      );
    }
    const command = (id: string) => ({
      id,
      label: 'Mark',
      mutatesDocument: false,
      shortcuts: ['Mod+Shift+K'],
      execute(context: PptxPluginContext<null>) {
        marks.push(context.pluginId);
        return { ok: true, status: 'executed' } as const;
      },
    });
    const owner = definePptxPlugin<null>({
      id: 'acme.review',
      createState: () => null,
      overlay: () => (
        <input data-testid="overlay-input" aria-label="Note" style={{ pointerEvents: 'auto' }} />
      ),
      commands: [command('mark')],
      toolbar: ['mark'],
    });
    const other = definePptxPlugin<null>({
      id: 'acme.other',
      createState: () => null,
      panel: { title: 'Other', placement: 'right', render: Panel },
    });
    const mounted = await mount({ plugins: [owner, other] });
    await until(
      () =>
        mounted.view.queryByTestId('overlay-input') !== null &&
        mounted.view.queryByTestId('panel-button') !== null
    );
    const story = (await read(mounted.api())).stories[0];
    const original = await storyText(mounted.api(), story.storyId);
    await act(async () => {
      mounted
        .api()
        .selectText({ slide: 1, shapeId: story.shapeId, storyId: story.storyId, start: 0, end: 0 });
    });
    fireEvent.keyDown(mounted.view.getByRole('application'), { key: 'q' });
    await settle();
    const typed = await storyText(mounted.api(), story.storyId);
    expect(typed).toBe(`q${original}`);
    return { ...mounted, story, original, typed };
  }

  const chord = { ctrlKey: true, metaKey: true };

  test('text fields in plugin chrome keep their editing keys and never edit the slide', async () => {
    const canvases = fakeContexts();
    try {
      const marks: string[] = [];
      const { api, view, story, typed } = await mountKeyboardProbe(marks);
      const fields = [
        view.getByTestId('overlay-input'),
        document.body.querySelector<HTMLElement>('[data-testid="portal-input"]')!,
      ];
      for (const field of fields) {
        for (const key of ['x', 'Backspace', 'Delete', 'Enter']) fireEvent.keyDown(field, { key });
        fireEvent.keyDown(field, { key: 'z', ...chord });
        fireEvent.keyDown(field, { key: 'z', shiftKey: true, ...chord });
        fireEvent.keyDown(field, { key: 'K', shiftKey: true, ...chord });
        await settle();
        expect(await storyText(api(), story.storyId)).toBe(typed);
      }
      expect(marks).toEqual([]);
    } finally {
      canvases.restore();
    }
  });

  test('built-in shortcuts act on the editor from plugin chrome; contributed ones run for their owner', async () => {
    const canvases = fakeContexts();
    try {
      const marks: string[] = [];
      const { api, view, story, original, typed } = await mountKeyboardProbe(marks);
      const text = () => storyText(api(), story.storyId);
      const toolbarMark = view.container.querySelector<HTMLElement>('button[aria-label="Mark"]')!;
      const steps: [HTMLElement, boolean, string][] = [
        [view.getByTestId('panel-button'), false, original],
        [
          within(view.getByTestId('plugin-dock-right')).getByRole('tab', { name: 'Other' }),
          true,
          typed,
        ],
        [
          document.body.querySelector<HTMLElement>('[data-testid="portal-button"]')!,
          false,
          original,
        ],
        [toolbarMark, true, typed],
      ];
      for (const [target, shiftKey, expected] of steps) {
        fireEvent.keyDown(target, { key: 'z', shiftKey, ...chord });
        await settle();
        expect(await text()).toBe(expected);
      }

      await act(async () => {
        api().selectText({
          slide: 1,
          shapeId: story.shapeId,
          storyId: story.storyId,
          start: 1,
          end: 3,
        });
      });
      const bold = api().commands.getState('bold').active;
      fireEvent.keyDown(view.getByTestId('panel-button'), { key: 'b', ...chord });
      await settle();
      expect(api().commands.getState('bold').active).not.toBe(bold);

      fireEvent.keyDown(view.getByTestId('panel-button'), { key: 'K', shiftKey: true, ...chord });
      fireEvent.keyDown(
        document.body.querySelector<HTMLElement>('[data-testid="portal-button"]')!,
        {
          key: 'K',
          shiftKey: true,
          ...chord,
        }
      );
      await settle();
      expect(marks).toEqual(['acme.review', 'acme.review']);
    } finally {
      canvases.restore();
    }
  });

  test('retained geometry refuses in the same step as a committed change', async () => {
    const canvases = fakeContexts();
    const rects = spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(
      () =>
        ({
          left: 0,
          top: 0,
          width: 640,
          height: 360,
          x: 0,
          y: 0,
          right: 640,
          bottom: 360,
        } as DOMRect)
    );
    try {
      let shapeId = '';
      let retained: PptxPluginGeometry | null = null;
      const duringChange: unknown[] = [];
      const contexts: PptxPluginContext<null>[] = [];
      const plugin = definePptxPlugin<null>({
        id: 'acme.review',
        createState: () => null,
        onEvent(context, event) {
          contexts.push(context);
          if (event.type === 'document-change' && retained) {
            duringChange.push(retained.getShapeRect(shapeId));
          }
        },
        overlay: () => <div data-testid="layout-marker" />,
      });
      const { api } = await mount({ plugins: [plugin] });
      const content = await read(api());
      shapeId = content.stories[0].shapeId;
      await until(() => contexts.some((context) => context.geometry !== null));
      retained = contexts.find((context) => context.geometry !== null)!.geometry!;
      expect(retained.getShapeRect(shapeId)).not.toBeNull();

      api().handle.setSlideNotes(content.slides[0].id, 'Direct');
      expect(retained.getShapeRect(shapeId)).toBeNull();
      expect(
        retained.toOverlayRect({ space: 'slide-px', rect: { x: 0, y: 0, width: 1, height: 1 } })
      ).toBeNull();
      expect(retained.getPositionAtPoint(10, 10)).toBeNull();
      await until(() => duringChange.length > 0);
      expect(duringChange).toEqual([null]);

      await act(async () => api().refresh());
      await until(() => last(contexts).geometry?.getShapeRect(shapeId) != null);
      expect(last(contexts).geometry!.layout.version).toBe(api().handle.version());
    } finally {
      rects.mockRestore();
      canvases.restore();
    }
  });

  test('replacement during initialization, a queued read or an action refuses', async () => {
    const images = pauseInsertedImages();
    try {
      let release!: () => void;
      const gate = new Promise<void>((done) => (release = done));
      const log: string[] = [];
      const errors: PptxPluginError[] = [];
      const contexts: PptxPluginContext<null>[] = [];
      let first: PptxPluginContext<null> | null = null;
      const plugin = definePptxPlugin<null>({
        id: 'acme.review',
        createState: () => null,
        async initialize(context) {
          context.onCleanup((reason) => {
            log.push(`cleanup:${reason}`);
          });
          if (first) return;
          first = context;
          await gate;
          const result = await context.read.version();
          log.push(result.ok ? 'read:ok' : `read:${result.failure.code}`);
        },
        onEvent(context, event) {
          contexts.push(context);
          if (event.type === 'load') log.push(`load:${event.reason}`);
        },
      });
      const { view, rerender } = await mount({
        plugins: [plugin],
        onPluginError: (error) => errors.push(error),
      });
      await until(() => first !== null);
      rerender({
        plugins: [plugin],
        onPluginError: (error) => errors.push(error),
        file: fixture.slice(),
      });
      await until(() => log.includes('load:replaced'));
      release();
      await until(() => log.some((entry) => entry.startsWith('read:')));
      expect(log).toEqual(['cleanup:document-replaced', 'load:replaced', 'read:document-replaced']);

      const second = last(contexts);
      insertImage(view);
      await until(() => images.pending.length > 0);
      const reading = second.read.readContent();
      let acted: unknown = null;
      let proceed!: () => void;
      const waiting = new Promise<void>((done) => (proceed = done));
      const action = second.run(async (context) => {
        await waiting;
        acted = await context.read.version();
      });
      rerender({
        plugins: [plugin],
        onPluginError: (error) => errors.push(error),
        file: fixture.slice(),
      });
      expect(await reading).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
      proceed();
      await action;
      expect(acted).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
      await until(() => log.filter((entry) => entry === 'load:replaced').length === 2);
      images.finish();
      await settle();
      expect(errors).toEqual([]);
    } finally {
      images.restore();
    }
  });

  test('queued writes check the grant and readOnly when they run, not when they were issued', async () => {
    const images = pauseInsertedImages();
    try {
      const { plugin, contexts, log } = recorder();
      const { api, view, rerender } = await mount({ plugins: [plugin], pluginGrants: WRITE });
      await until(() => log.includes('load:loaded'));
      const edits = last(contexts).edits!;
      const before = await read(api());
      const slideId = before.slides[1].id;

      insertImage(view);
      await until(() => images.pending.length > 0);
      const revoked = edits.applyEdits(notesRequest(before.version, slideId, 'Revoked'));
      rerender({ plugins: [plugin], pluginGrants: {} });
      await until(() => log.includes('grants-change'));
      await act(async () => images.finish());
      expect(await revoked).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });

      rerender({ plugins: [plugin], pluginGrants: WRITE });
      await until(() => log.filter((entry) => entry === 'grants-change').length === 2);
      const current = await read(api());
      insertImage(view);
      await until(() => images.pending.length > 0);
      const readOnly = last(contexts).edits!.applyEdits(
        notesRequest(current.version, slideId, 'Read-only')
      );
      rerender({ plugins: [plugin], pluginGrants: WRITE, readOnly: true });
      await until(() => log.includes('mode-change:true'));
      await act(async () => images.finish());
      expect(await readOnly).toMatchObject({ ok: false, failure: { code: 'read-only' } });
      expect((await read(api())).slides[1].notes ?? '').toBe(before.slides[1].notes ?? '');
    } finally {
      images.restore();
    }
  });

  test('a deferred command completes with the grant and presentation current then', async () => {
    let store: PptxCommandStore | null = null;
    function Panel() {
      store = usePptxCommands();
      return null;
    }
    const grants = { 'acme.review': { commands: ['zoom'] } } as const;
    const { plugin, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'left', render: Panel },
    });
    const { api, rerender } = await mount({ plugins: [plugin], pluginGrants: grants });
    await until(() => store !== null && log.includes('load:loaded'));
    const revoked = pptxCommandController(store!)!.prepare('zoom');
    rerender({ plugins: [plugin], pluginGrants: {} });
    await until(() => log.includes('grants-change'));
    expect(await revoked.execute({ scale: 2 })).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });

    rerender({ plugins: [plugin], pluginGrants: grants });
    await until(() => log.filter((entry) => entry === 'grants-change').length === 2);
    const replaced = pptxCommandController(store!)!.prepare('zoom');
    rerender({ plugins: [plugin], pluginGrants: grants, file: fixture.slice() });
    await until(() => log.includes('load:replaced'));
    expect(await replaced.execute({ scale: 2 })).toMatchObject({ ok: false });
    expect(api().commands.getState('zoom').value).toBe('fit');
  });

  test('a deletion-only batch publishes exactly one change', async () => {
    const { plugin, log } = recorder();
    const { api } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    const content = await read(api());
    const story = content.stories.find((candidate) => candidate.text.length > 2)!;
    const { slideId, shapeId, storyId } = story;
    let applied!: Awaited<ReturnType<PptxEditorApi['applyEdits']>>;
    await act(async () => {
      applied = await api().applyEdits({
        expectVersion: content.version,
        steps: [
          {
            op: 'deleteText',
            target: { kind: 'range', slideId, shapeId, storyId, start: 0, end: 1 },
          },
        ],
      });
    });
    expect(applied).toMatchObject({ ok: true, applied: true });
    await settle(50);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toEqual([
      `document-change:${applied.ok ? applied.version : ''}`,
    ]);
  });

  test('registrations follow id and revision; duplicates and late cleanups are handled', async () => {
    const errors: PptxPluginError[] = [];
    const a = recorder('acme.a');
    const b = recorder('acme.b');
    const onPluginError = (error: PptxPluginError) => errors.push(error);
    const { rerender } = await mount({ plugins: [a.plugin, b.plugin], onPluginError });
    await until(() => a.log.includes('load:loaded') && b.log.includes('load:loaded'));
    const firstA = a.contexts[0];

    rerender({ plugins: [b.plugin, a.plugin], onPluginError });
    const sameRevision = recorder('acme.a');
    rerender({ plugins: [b.plugin, sameRevision.plugin], onPluginError });
    await settle();
    expect([...a.log, ...b.log].filter((entry) => entry.startsWith('cleanup'))).toEqual([]);
    expect(firstA.lifetimeSignal.aborted).toBe(false);
    expect(sameRevision.log).toEqual([]);

    const revised = recorder('acme.a', { revision: 2 });
    rerender({ plugins: [b.plugin, revised.plugin], onPluginError });
    await until(() => revised.log.includes('load:attached'));
    expect(a.log).toContain('cleanup:definition-replaced');
    expect(firstA.lifetimeSignal.aborted).toBe(true);

    const late: string[] = [];
    firstA.onCleanup((reason) => {
      late.push(reason);
    });
    expect(late).toEqual(['definition-replaced']);

    const twin = recorder('acme.b');
    rerender({ plugins: [b.plugin, twin.plugin, revised.plugin], onPluginError });
    await until(() => b.log.includes('cleanup:removed'));
    expect(twin.log).toEqual([]);
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([
      ['acme.b', 'definition'],
    ]);
  });
});
