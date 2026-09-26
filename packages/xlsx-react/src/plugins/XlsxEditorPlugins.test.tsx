import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { StrictMode, useState } from 'react';
import { createPortal } from 'react-dom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  cellRect,
  initWasm,
  openWorkbook,
  selectionAt,
  type XlsxEditRequest,
} from '@betteroffice/xlsx';
import { xlsxCommandController } from '../commands/createXlsxCommandStore';
import { pluginDefinition } from './defineXlsxPlugin';
import * as publicApi from '../index';
import {
  EditorToolbar,
  ToolbarCommandButton,
  XlsxEditor,
  XlsxPluginToolbar,
  defineXlsxPlugin,
  useXlsxCommand,
  useXlsxCommands,
  type XlsxCommandResult,
  type XlsxCommandStore,
  type XlsxEditorApi,
  type XlsxEditorProps,
  type XlsxPlugin,
  type XlsxPluginCommandResult,
  type XlsxPluginContext,
  type XlsxPluginDefinition,
  type XlsxPluginError,
  type XlsxPluginEvent,
  type XlsxPluginGeometry,
} from '../index';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

const root = resolve(import.meta.dir, '../../../..');
const VIEWPORT = { width: 800, height: 600 };
const LAYOUT = [
  ['clientWidth', VIEWPORT.width],
  ['clientHeight', VIEWPORT.height],
] as const;
const quiet = { error: console.error, warn: console.warn };
const originalGetContext = HTMLCanvasElement.prototype.getContext;
const originalLayout = LAYOUT.map(
  ([property]) =>
    [property, Object.getOwnPropertyDescriptor(HTMLElement.prototype, property)] as const
);
let fixture: Uint8Array;
let charted: Uint8Array;

function stubContext(): CanvasRenderingContext2D {
  return new Proxy({} as Record<string, unknown>, {
    get(target, key) {
      if (key === 'measureText') return () => ({ width: 0 });
      if (typeof key === 'string' && key in target) return target[key];
      return () => {};
    },
    set(target, key, value) {
      target[key as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

beforeAll(async () => {
  HTMLCanvasElement.prototype.getContext = (() =>
    stubContext()) as unknown as HTMLCanvasElement['getContext'];
  for (const [property, value] of LAYOUT) {
    Object.defineProperty(HTMLElement.prototype, property, {
      configurable: true,
      get: () => value,
    });
  }
  await initWasm(
    new Uint8Array(
      readFileSync(resolve(root, 'packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm'))
    )
  );
  fixture = new Uint8Array(readFileSync(resolve(root, 'packages/xlsx/test-fixtures/sample.xlsx')));
  charted = new Uint8Array(readFileSync(resolve(root, 'packages/xlsx/test-fixtures/charts.xlsx')));
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  await new Promise((done) => setTimeout(done, 50));
  console.error = quiet.error;
  console.warn = quiet.warn;
  HTMLCanvasElement.prototype.getContext = originalGetContext;
  for (const [property, descriptor] of originalLayout) {
    if (descriptor) Object.defineProperty(HTMLElement.prototype, property, descriptor);
  }
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

type State = { count: number };

function describeEvent(event: XlsxPluginEvent): string | null {
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

function recorder(id = 'acme.review', extra: Partial<XlsxPluginDefinition<State>> = {}) {
  const log: string[] = [];
  const events: XlsxPluginEvent[] = [];
  const contexts: XlsxPluginContext<State>[] = [];
  const plugin = defineXlsxPlugin<State>({
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

async function mount(props: Partial<XlsxEditorProps> = {}, strict = false) {
  const ready: XlsxEditorApi[] = [];
  const element = (next: Partial<XlsxEditorProps>) => {
    const editor = <XlsxEditor file={fixture} onReady={(api) => void ready.push(api)} {...next} />;
    return strict ? <StrictMode>{editor}</StrictMode> : editor;
  };
  const view = render(element(props));
  await until(() => ready.length > 0);
  return {
    api: () => last(ready),
    ready,
    view,
    rerender: (next: Partial<XlsxEditorProps>) => view.rerender(element(next)),
  };
}

function setCell(
  expectVersion: string,
  a1: string,
  value: string,
  sheetId = 'sheet:0'
): XlsxEditRequest {
  return {
    expectVersion,
    steps: [
      { op: 'setCellInputs', target: { sheetId, range: { kind: 'a1', a1 } }, inputs: [[value]] },
    ],
  };
}

const input = (api: XlsxEditorApi, row: number, col: number, sheet = 0) =>
  api.handle.cell(sheet, row, col).input;

const WRITE = { 'acme.review': { document: 'write', editBatches: true } } as const;

/** A clipboard whose reads finish only when the test resolves them, in order. */
function holdClipboard() {
  const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const reads: ((text: string) => void)[] = [];
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: { readText: () => new Promise<string>((resolve) => reads.push(resolve)) },
  });
  return {
    reads,
    resolve: (text: string) => reads.shift()!(text),
    restore: () => {
      if (original) Object.defineProperty(navigator, 'clipboard', original);
      else Reflect.deleteProperty(navigator, 'clipboard');
    },
  };
}

/** Queues a paste into B3 that lands once the clipboard read resolves. */
async function queuePaste(view: ReturnType<typeof render>, api: XlsxEditorApi) {
  await act(async () => {
    api.selectCells(0, selectionAt({ row: 2, col: 1 }));
  });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
}

const chord = { ctrlKey: true, metaKey: true };

const sheetTabs = (view: ReturnType<typeof render>) =>
  within(view.getByTestId('xlsx-sheet-tabs')).getAllByRole('tab');

describe('XlsxEditor plugins', () => {
  test('initialize, load, one change per committed change, and cleanup on removal', async () => {
    const { plugin, log, contexts } = recorder();
    const { api, rerender } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    expect(log.slice(0, 2)).toEqual(['initialize', 'load:loaded']);
    const start = api().handle.version();
    expect(contexts[1].snapshot.version).toBe(start);
    const changes = () => log.filter((entry) => entry.startsWith('document-change'));

    let applied!: Awaited<ReturnType<XlsxEditorApi['applyEdits']>>;
    await act(async () => {
      applied = await api().applyEdits(setCell(start, 'B3', '321'));
    });
    expect(applied).toMatchObject({ ok: true, applied: true });
    await settle(30);
    expect(changes()).toEqual([`document-change:${applied.ok ? applied.version : ''}`]);

    await act(async () => {
      const stale = await api().applyEdits(setCell(start, 'B3', 'late'));
      expect(stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      const noop = await api().applyEdits(setCell(api().handle.version(), 'B3', '321'));
      expect(noop).toMatchObject({ ok: true, applied: false });
    });
    await settle(30);
    expect(changes()).toHaveLength(1);

    for (const step of ['bold', 'undo', 'redo'] as const) {
      await act(async () => {
        expect(await api().commands.execute(step, null)).toMatchObject({ ok: true });
      });
      await settle(30);
      expect(last(changes())).toBe(`document-change:${api().handle.version()}`);
    }
    expect(changes()).toHaveLength(4);

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

  test('document changes carry the settled version; remote updates publish too', async () => {
    const recalculated: string[] = [];
    const { plugin, log } = recorder('acme.review', {
      async onEvent(context, event) {
        if (event.type !== 'document-change') return;
        const read = await context.read.readCells({
          ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'D3' } }],
        });
        if (read.ok)
          recalculated.push(
            `${read.version === event.version}:${read.ranges[0].cells[0][0].displayText}`
          );
        log.push('document-change');
      },
    });
    const { api } = await mount({ plugins: [plugin], collaboration: {} });
    await until(() => api().handle !== undefined);
    await settle(30);
    await act(async () => {
      await api().applyEdits(setCell(api().handle.version(), 'B3', '4242'));
    });
    await until(() => recalculated.length === 1);
    expect(recalculated).toEqual(['true:4299']);

    const replica = openWorkbook(fixture, { collaborative: true, clientId: 7171 });
    try {
      replica.applyUpdate(api().handle.encodeStateAsUpdate());
      replica.editCell(0, 2, 1, '43');
      await act(async () => {
        api().handle.applyUpdate(replica.encodeStateAsUpdate(api().handle.encodeStateVector()));
      });
    } finally {
      replica.dispose();
    }
    await until(() => recalculated.length === 2);
    expect(recalculated[1]).toBe('true:100');
    expect(log.filter((entry) => entry === 'document-change')).toHaveLength(2);
  });

  test('a default plugin reads and navigates but cannot mutate through any path', async () => {
    const calls: string[] = [];
    function Probe() {
      const commands = useXlsxCommands();
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
    const read = await context.read.readCells({
      ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3:B4' } }],
    });
    expect(read.ok).toBe(true);
    if (!read.ok) return;
    expect(read.ranges[0].cells.map((row) => row[0].displayText)).toEqual(['100', '200']);
    expect(await context.read.findText({ text: 'Quarterly' })).toMatchObject({
      ok: true,
      version: read.version,
    });
    expect(await context.read.validateEdits(setCell(read.version, 'B3', '1'))).toMatchObject({
      ok: true,
      wouldApply: true,
    });
    const origin = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
    expect(
      await context.navigation.selectCells(
        { sheetId: 'sheet:missing', selection: origin },
        { expectVersion: read.version }
      )
    ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    expect(
      await context.navigation.selectCells(
        { sheetId: 'sheet:1', selection: origin },
        { expectVersion: 'stale' }
      )
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(api().handle.version()).toBe(read.version);
    expect(input(api(), 2, 1)).toBe('100');
  });

  test('granted built-in mutations refuse without a policy path; granted views run', async () => {
    function Probe() {
      const bold = useXlsxCommands().getState('bold');
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
    const clipboard = holdClipboard();
    try {
      const { plugin, contexts, log } = recorder();
      const { api, view, rerender } = await mount({ plugins: [plugin], pluginGrants: WRITE });
      await until(() => log.includes('load:loaded'));
      const edits = last(contexts).edits!;
      expect(edits).not.toBeNull();
      const before = api().handle.version();

      await queuePaste(view, api());
      let settled = false;
      const late = edits.applyEdits(setCell(before, 'B4', 'Too late')).finally(() => {
        settled = true;
      });
      await settle();
      expect(settled).toBe(false);
      await act(async () => clipboard.resolve('Pasted'));
      expect(await late).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      expect(input(api(), 2, 1)).toBe('Pasted');

      const applied = await act(() =>
        edits.applyEdits(setCell(api().handle.version(), 'B4', 'Plugin'))
      );
      expect(applied).toMatchObject({ ok: true, applied: true });
      expect(input(api(), 3, 1)).toBe('Plugin');
      await act(async () => {
        await api().commands.execute('undo', null);
      });
      expect(input(api(), 3, 1)).toBe('200');
      const version = api().handle.version();
      expect(
        await edits.applyEdits({ ...setCell(version, 'B4', 'Hidden'), history: 'none' })
      ).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });

      rerender({ plugins: [plugin], pluginGrants: WRITE, readOnly: true });
      await until(() => log.includes('mode-change:true'));
      expect(await edits.applyEdits(setCell(version, 'B4', 'Read-only'))).toMatchObject({
        ok: false,
        version,
        failure: { code: 'read-only' },
      });
      expect(await last(contexts).read.validateEdits(setCell(version, 'B4', 'x'))).toMatchObject({
        ok: false,
        failure: { code: 'read-only' },
      });

      rerender({ plugins: [plugin], pluginGrants: {}, readOnly: false });
      await until(() => log.includes('grants-change'));
      expect(await edits.applyEdits(setCell(version, 'B4', 'Revoked'))).toMatchObject({
        ok: false,
        failure: { code: 'permission-denied' },
      });

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
      expect(await edits.applyEdits(setCell(version, 'B4', 'Revoked in place'))).toMatchObject({
        ok: false,
        failure: { code: 'permission-denied' },
      });
      expect(api().handle.version()).toBe(version);
    } finally {
      clipboard.restore();
    }
  });

  test('contributed commands run with their plugin, from the toolbar, API and shortcuts', async () => {
    const calls: string[] = [];
    const errors: XlsxPluginError[] = [];
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+Shift+M', 'Mod+B'],
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
    ]);

    const more = view.getByTestId('xlsx-toolbar-more');
    act(() => more.focus());
    fireEvent.keyDown(more, { key: 'ArrowDown' });
    const menu = view.getByRole('menu', { name: 'More toolbar items' });
    const entry = within(menu).getByRole('menuitemcheckbox', { name: /^Mark/ });
    expect(entry.getAttribute('aria-checked')).toBe('false');
    fireEvent.click(entry);
    await settle();
    expect(calls).toEqual(['acme.review']);
    expect(api().commands.getState(id).active).toBe(true);

    expect(await act(() => api().commands.execute(id, null))).toMatchObject({ ok: true });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'M', shiftKey: true, ...chord });
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
          <EditorToolbar.Toolbar>
            <XlsxPluginToolbar />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      ),
    });
    await until(() => custom.view.container.querySelector('button[aria-label="Mark"]') !== null);
  });

  test('contributed commands return their own failures and batch refusals unchanged', async () => {
    const returned: XlsxPluginCommandResult[] = [];
    const paused = { code: 'paused', message: 'Paused' };
    const owner = recorder('acme.review', {
      commands: [
        {
          id: 'quota',
          label: 'Quota',
          mutatesDocument: false,
          getState: (context) =>
            context.state.count > 0
              ? { enabled: false, disabledReason: paused }
              : { enabled: true },
          execute(context) {
            context.setState({ count: 1 });
            returned.push({ ok: false, failure: { code: 'quota-exceeded', message: 'Try later' } });
            return last(returned);
          },
        },
        {
          id: 'stale',
          label: 'Stale',
          mutatesDocument: true,
          async execute(context) {
            const result = await context.edits!.applyEdits(setCell('stale', 'B4', 'Late'));
            returned.push(result.ok ? { ok: true, status: 'executed' } : result);
            return last(returned);
          },
        },
      ],
    });
    const { api } = await mount({ plugins: [owner.plugin], pluginGrants: WRITE });
    await until(() => owner.log.includes('load:loaded'));
    const commands = api().commands;

    expect(await act(() => commands.execute('plugin:acme.review/quota', null))).toBe(returned[0]);
    expect(await act(() => commands.execute('plugin:acme.review/quota', null))).toEqual({
      ok: false,
      failure: paused,
    });
    const stale = await act(() => commands.execute('plugin:acme.review/stale', null));
    expect(stale).toBe(returned[1]);
    if (stale.ok || !('version' in stale)) throw new Error('expected a batch refusal');
    expect([stale.version, stale.failure.code]).toEqual([api().handle.version(), 'stale-version']);

    const builtIn: XlsxCommandResult = await act(() => commands.execute('zoom', { scale: 2 }));
    expect(builtIn.ok).toBe(true);
  });

  test('a failing contribution is isolated and the others keep working', async () => {
    const errors: XlsxPluginError[] = [];
    const broken = defineXlsxPlugin({
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
    const throwing = defineXlsxPlugin({
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
    const rejectsAction = defineXlsxPlugin({
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

  test('replacing the workbook ends the activation and loads a fresh one', async () => {
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

  test('a workbook replaced from onReady never reaches plugins; its successor loads once', async () => {
    const { plugin, log, contexts } = recorder();
    const plugins = [plugin];
    const ready: XlsxEditorApi[] = [];
    function Host() {
      const [file, setFile] = useState(fixture);
      return (
        <XlsxEditor
          file={file}
          plugins={plugins}
          onReady={(api) => {
            ready.push(api);
            if (ready.length === 1) setFile(charted);
          }}
        />
      );
    }
    render(<Host />);
    await until(() => ready.length === 2 && log.includes('load:loaded'));
    await settle(50);
    expect(log).toEqual(['initialize', 'load:loaded']);
    expect(await last(contexts).read.version()).toMatchObject({
      ok: true,
      version: last(ready).handle.version(),
    });
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

  test('navigation resolves sheet ids after input, keeps focus and reports selections', async () => {
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
    const version = api().handle.version();
    const scroll = view.getByTestId('xlsx-scroll');
    const field = view.getByTestId('panel-input');
    act(() => field.focus());
    const reversed = { anchor: { row: 3, col: 2 }, focus: { row: 1, col: 0 } };
    expect(
      await act(() =>
        context().navigation.selectCells(
          { sheetId: 'sheet:1', selection: reversed },
          { expectVersion: version }
        )
      )
    ).toEqual({ ok: true });
    expect(sheetTabs(view)[1].getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(field);
    await until(() =>
      events.some(
        (event) =>
          event.type === 'selection-change' &&
          event.selection?.sheetId === 'sheet:1' &&
          event.selection.cells?.anchor.row === 3
      )
    );
    expect(context().snapshot.selection).toEqual({
      sheetId: 'sheet:1',
      sheetIndex: 1,
      cells: reversed,
      chartId: null,
    });

    expect(
      await act(() =>
        context().navigation.selectCells(
          { sheetId: 'sheet:1', selection: selectionAt({ row: 0, col: 1 }) },
          { expectVersion: version, focus: true }
        )
      )
    ).toEqual({ ok: true });
    expect(document.activeElement).toBe(scroll);

    const position = api().handle.cellPosition(1, 40, 0);
    expect(
      await act(() =>
        context().navigation.scrollToCell(
          { sheetId: 'sheet:1', row: 40, col: 0 },
          { expectVersion: version, align: 'start' }
        )
      )
    ).toEqual({ ok: true });
    await until(() => scroll.scrollTop === position.y);
    expect(context().snapshot.selection?.cells).toEqual(selectionAt({ row: 0, col: 1 }));

    expect(
      await act(() =>
        context().navigation.scrollToCell(
          { sheetId: 'sheet:0', row: 2, col: 1 },
          { expectVersion: version }
        )
      )
    ).toEqual({ ok: true });
    await until(() => context().snapshot.selection?.sheetId === 'sheet:0');
    expect(context().snapshot.selection?.cells).toBeNull();

    for (const target of [
      { sheetId: 'sheet:0', row: 2_000_000, col: 0 },
      { sheetId: 'sheet:0', row: -1, col: 0 },
      { sheetId: 'sheet:9', row: 0, col: 0 },
    ]) {
      expect(
        await context().navigation.scrollToCell(target, { expectVersion: version })
      ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    }
    await act(async () => {
      await api().applyEdits(setCell(version, 'B3', 'Moved on'));
    });
    expect(
      await context().navigation.selectCells(
        { sheetId: 'sheet:1', selection: selectionAt({ row: 0, col: 0 }) },
        { expectVersion: version }
      )
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(sheetTabs(view)[0].getAttribute('aria-selected')).toBe('true');
  });

  for (const change of ['replaced', 'removed'] as const) {
    const what = change === 'replaced' ? 'the workbook is replaced' : 'its plugin is removed';
    test(`a reveal deferred to the next frame is dropped once ${what}`, async () => {
      const { plugin, contexts, log } = recorder();
      const { api, view, rerender } = await mount({ plugins: [plugin] });
      await until(() => log.includes('load:loaded'));
      const scroll = view.getByTestId('xlsx-scroll');
      const navigated = await act(async () => {
        const result = await last(contexts).navigation.scrollToCell(
          { sheetId: 'sheet:0', row: 60, col: 0 },
          { expectVersion: api().handle.version(), align: 'start' }
        );
        rerender(
          change === 'replaced' ? { plugins: [plugin], file: fixture.slice() } : { plugins: [] }
        );
        return result;
      });
      expect(navigated).toEqual({ ok: true });
      await until(() => log.includes(change === 'replaced' ? 'load:replaced' : 'cleanup:removed'));
      await settle(50);
      expect(scroll.scrollTop).toBe(0);
    });
  }

  test('layouts follow the painted frame, its version, sheet, zoom and scroll', async () => {
    const layouts: (string | null)[] = [];
    const contexts: XlsxPluginContext<null>[] = [];
    const plugin = defineXlsxPlugin<null>({
      id: 'acme.review',
      createState: () => null,
      onEvent(context, event) {
        contexts.push(context);
        if (event.type === 'layout-change') layouts.push(event.layout?.version ?? null);
      },
      overlay: ({ context, geometry }) => (
        <div
          data-testid="layout-marker"
          data-layout={geometry.layout.id}
          data-version={geometry.layout.version}
          data-snapshot={context.snapshot.version}
          data-sheet={geometry.layout.sheetId}
          data-zoom={geometry.layout.zoom}
          data-y={geometry.layout.viewport.y}
        />
      ),
    });
    const { api, view } = await mount({ plugins: [plugin] });
    const marker = () => view.container.querySelector<HTMLElement>('[data-testid="layout-marker"]');
    await until(() => marker() !== null);
    const version = api().handle.version();
    expect(marker()!.dataset).toMatchObject({ version, snapshot: version, sheet: 'sheet:0' });
    const layer = view.getByTestId('plugin-overlays');
    const canvas = view.getByTestId('xlsx-scroll').querySelector('canvas')!;
    expect(layer.parentElement).toBe(canvas.parentElement);
    expect(
      layer.compareDocumentPosition(view.getByTestId('xlsx-overlay-host')) &
        Node.DOCUMENT_POSITION_FOLLOWING
    ).toBeTruthy();
    expect(layer.style.pointerEvents).toBe('none');

    let next = '';
    await act(async () => {
      const applied = await api().applyEdits(setCell(version, 'B3', 'Changed'));
      if (applied.ok) next = applied.version;
    });
    await until(() => marker()?.dataset.version === next);
    const afterEdit = layouts.slice(layouts.lastIndexOf(version) + 1);
    expect(afterEdit).not.toContain(version);
    expect(last(afterEdit)).toBe(next);

    await act(async () => {
      await api().commands.execute('zoom', { scale: 1.5 });
    });
    await until(() => marker()?.dataset.zoom === '1.5');

    await act(async () => {
      await last(contexts).navigation.selectCells(
        { sheetId: 'sheet:1', selection: selectionAt({ row: 0, col: 0 }) },
        { expectVersion: next }
      );
    });
    await until(() => marker()?.dataset.sheet === 'sheet:1');
    expect(last(contexts).snapshot.layout?.sheetId).toBe('sheet:1');

    const scroll = view.getByTestId('xlsx-scroll');
    const before = marker()!.dataset.layout;
    scroll.scrollTop = 150;
    fireEvent.scroll(scroll);
    await until(() => marker()?.dataset.y === String(150 / 1.5));
    expect(marker()!.dataset.layout).not.toBe(before);
  });

  test('frozen panes keep their cells in place while the body scrolls', async () => {
    const source = openWorkbook(fixture.slice());
    source.applyOps([
      {
        type: 'setFreezePane',
        sheet: 0,
        pane: { rows: 2, cols: 1, top_left: { row: 2, col: 1 } },
      },
    ]);
    const frozen = source.save();
    source.dispose();
    const contexts: XlsxPluginContext<null>[] = [];
    const plugin = defineXlsxPlugin<null>({
      id: 'acme.frozen',
      createState: () => null,
      onEvent(context) {
        contexts.push(context);
      },
    });
    const { api, view } = await mount({ plugins: [plugin], file: frozen });
    const geometry = () => last(contexts)?.geometry ?? null;
    await until(() => geometry() !== null);
    const scroll = view.getByTestId('xlsx-scroll');
    scroll.scrollTop = 200;
    fireEvent.scroll(scroll);
    await until(() => geometry()?.layout.viewport.y === 200);
    const frame = api().handle.displayList(geometry()!.layout.viewport).grid!;
    const cell = (row: number, col: number) =>
      geometry()!.getCellRect({ sheetId: 'sheet:0', row, col });
    expect(cell(0, 0)).toMatchObject({ x: 0, y: frame.rowOffsets[0] });
    expect(cell(1, 0)?.y).toBe(frame.rowOffsets[1]);
    const body = frame.rowIndices![2];
    expect(body).toBeGreaterThan(2);
    expect(cell(body, 1)?.y).toBe(frame.rowOffsets[2]);
    expect(cell(2, 1)).toBeNull();
  });

  test('public presenters and hooks bind contributed commands', async () => {
    const calls: string[] = [];
    function Bound() {
      const mark = useXlsxCommand('plugin:acme.review/mark');
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
          <EditorToolbar.Toolbar>
            <ToolbarCommandButton id="plugin:acme.review/mark" />
            <ToolbarCommandButton id="plugin:acme.review/absent" label="Absent" />
            <Bound />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      ),
    });
    const button = (name: string) =>
      view.container.querySelector<HTMLElement>(`button[aria-label="${name}"]`);
    await until(() => button('Mark') !== null);
    expect(button('Absent')).toBeNull();
    expect(view.getByTestId('bound').textContent).toMatch(/^Mark:true:.+M$/);
    fireEvent.click(button('Mark')!);
    await settle();
    expect(calls).toEqual(['acme.review']);
  });

  test('two editors keep separate activations and events', async () => {
    const { plugin, contexts, log } = recorder();
    const first = await mount({ plugins: [plugin] });
    const second = await mount({ plugins: [plugin], file: fixture.slice() });
    await until(() => log.filter((entry) => entry === 'load:loaded').length === 2);
    expect(new Set(contexts.map((context) => context.lifetimeSignal)).size).toBe(2);
    await act(async () => {
      await first.api().applyEdits(setCell(first.api().handle.version(), 'B3', 'One'));
    });
    await settle(30);
    const changed = contexts.filter(
      (context) => context.snapshot.version === first.api().handle.version()
    );
    expect(new Set(changed.map((context) => context.lifetimeSignal)).size).toBe(1);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toHaveLength(1);
    expect(input(second.api(), 2, 1)).toBe('100');
  });

  test('the demo review plugin works through the documented exports', async () => {
    mock.module('@betteroffice/xlsx-react', () => publicApi);
    const demo = '../../../../apps/demo/app/xlsx/ReviewPlugin';
    const { reviewPlugin } = (await import(demo)) as { reviewPlugin: XlsxPlugin };
    const { api, view, rerender } = await mount({ plugins: [reviewPlugin] });
    const body = within(view.container);
    const action = () => body.getByRole('button', { name: 'Mark selection reviewed' });
    await until(() => body.queryByText(/3 sheets, version/) !== null);
    await until(() => body.queryByText(/Selection A1: 1 cells/) !== null);
    expect(action().hasAttribute('disabled')).toBe(true);
    const mark = 'plugin:demo.review/mark-reviewed' as const;
    expect(api().commands.getState(mark).disabledReason?.code).toBe('write-not-granted');
    expect(api().commands.getDescriptor(mark)?.label).toBe('Mark reviewed');

    const grants = { 'demo.review': { document: 'write', editBatches: true } } as const;
    rerender({ plugins: [reviewPlugin], pluginGrants: grants });
    await until(() => !action().hasAttribute('disabled'));
    fireEvent.click(action());
    await until(() => body.queryByText('Marked the selection as reviewed.') !== null);
    expect(api().handle.selectionFormatting(0, 'A1:A1').fillColor).toBe('#d9ead3');

    fireEvent.click(body.getByRole('button', { name: 'Summary' }));
    await until(() => sheetTabs(view)[1].getAttribute('aria-selected') === 'true');

    rerender({ plugins: [reviewPlugin], pluginGrants: grants, readOnly: true });
    await until(() => action().hasAttribute('disabled'));

    rerender({ plugins: [], pluginGrants: grants });
    await until(() => body.queryByRole('button', { name: 'Mark selection reviewed' }) === null);
  });

  test('plugins that are not defined through defineXlsxPlugin are reported, not run', async () => {
    const errors: XlsxPluginError[] = [];
    await mount({
      plugins: [{ id: 'raw' } as unknown as XlsxPlugin],
      onPluginError: (error) => errors.push(error),
    });
    await settle();
    expect(errors.map((error) => [error.pluginId, error.phase])).toEqual([['raw', 'definition']]);
  });
});

describe('XlsxEditor plugin boundaries and races', () => {
  /** An editor with an overlay input and a panel whose button and field also render in a portal. */
  async function mountKeyboardProbe(marks: string[]) {
    function Panel() {
      return (
        <>
          <button type="button" data-testid="panel-button">
            Panel
          </button>
          <input data-testid="panel-input" aria-label="Panel note" />
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
    const owner = defineXlsxPlugin<null>({
      id: 'acme.review',
      createState: () => null,
      overlay: () => (
        <input data-testid="overlay-input" aria-label="Note" style={{ pointerEvents: 'auto' }} />
      ),
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+Shift+K'],
          execute(context) {
            marks.push(context.pluginId);
            return { ok: true, status: 'executed' };
          },
        },
      ],
      toolbar: ['mark'],
    });
    const other = defineXlsxPlugin<null>({
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
    await act(async () => {
      mounted.api().selectCells(0, selectionAt({ row: 2, col: 1 }));
    });
    const version = mounted.api().handle.version();
    await act(async () => {
      await mounted.api().applyEdits(setCell(version, 'B3', 'Edited'));
    });
    return mounted;
  }

  test('text fields and clicks in plugin chrome never reach the grid', async () => {
    const clipboard = holdClipboard();
    try {
      const marks: string[] = [];
      const { api, view } = await mountKeyboardProbe(marks);
      const version = api().handle.version();
      const nameBox = view.getByTestId('xlsx-name-box') as HTMLInputElement;
      const fields = [
        view.getByTestId('overlay-input'),
        view.getByTestId('panel-input'),
        document.body.querySelector<HTMLElement>('[data-testid="portal-input"]')!,
      ];
      for (const field of fields) {
        for (const key of ['x', 'Backspace', 'Delete', 'Enter', 'F2', 'ArrowDown']) {
          fireEvent.keyDown(field, { key });
        }
        for (const key of ['z', 'v', 'x']) fireEvent.keyDown(field, { key, ...chord });
        fireEvent.keyDown(field, { key: 'z', shiftKey: true, ...chord });
        fireEvent.keyDown(field, { key: 'K', shiftKey: true, ...chord });
        fireEvent.mouseDown(field);
        fireEvent.click(field);
        fireEvent.doubleClick(field);
        await settle();
        expect(api().handle.version()).toBe(version);
        expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
        expect(nameBox.value).toBe('B3');
      }
      expect(clipboard.reads).toHaveLength(0);
      expect(marks).toEqual([]);
    } finally {
      clipboard.restore();
    }
  });

  test('pointer movement over plugin chrome drives no grid gesture or hover', async () => {
    const opened = openWorkbook(charted.slice());
    const painted = opened.displayList({ x: 0, y: 0, ...VIEWPORT });
    const [chart] = painted.charts!;
    const cell = (row: number, col: number) => {
      const rect = cellRect(painted.grid!, row, col)!;
      return { clientX: rect.x + rect.w / 2, clientY: rect.y + rect.h / 2 };
    };
    const center = {
      clientX: chart.clip.x + chart.clip.w / 2,
      clientY: chart.clip.y + chart.clip.h / 2,
    };
    const { plugin } = recorder('acme.review', {
      overlay: () => (
        <input data-testid="overlay-input" aria-label="Note" style={{ pointerEvents: 'auto' }} />
      ),
    });
    const { view } = await mount({ plugins: [plugin], file: charted });
    await until(() => view.queryByTestId('overlay-input') !== null);
    const surface = view.getByTestId('xlsx-scroll');
    const overlay = view.getByTestId('overlay-input');
    const nameBox = view.getByTestId('xlsx-name-box') as HTMLInputElement;

    fireEvent.mouseMove(surface, center);
    expect(surface.style.cursor).toBe('move');
    fireEvent.mouseMove(overlay, center);
    expect(surface.style.cursor).toBe('default');

    fireEvent.mouseDown(surface, cell(0, 0));
    fireEvent.mouseMove(overlay, { ...cell(10, 0), buttons: 1 });
    fireEvent.mouseUp(window, cell(10, 0));
    await settle();
    expect(nameBox.value).toBe('A1');

    fireEvent.mouseDown(surface, center);
    await until(() => view.queryByTestId('xlsx-chart-selection') !== null);
    const left = view.getByTestId('xlsx-chart-selection').style.left;
    fireEvent.mouseMove(overlay, {
      clientX: center.clientX + 40,
      clientY: center.clientY + 24,
      buttons: 1,
    });
    expect(view.getByTestId('xlsx-chart-selection').style.left).toBe(left);
    fireEvent.mouseUp(window, center);
    opened.dispose();
  });

  test('built-in shortcuts act on the editor from plugin chrome; contributed ones run for their owner', async () => {
    const marks: string[] = [];
    const { api, view } = await mountKeyboardProbe(marks);
    const toolbarMark = view.container.querySelector<HTMLElement>('button[aria-label="Mark"]')!;
    const steps: [HTMLElement, boolean, string][] = [
      [view.getByTestId('panel-button'), false, '100'],
      [
        within(view.getByTestId('plugin-dock-right')).getByRole('tab', { name: 'Other' }),
        true,
        'Edited',
      ],
      [document.body.querySelector<HTMLElement>('[data-testid="portal-button"]')!, false, '100'],
      [toolbarMark, true, 'Edited'],
    ];
    for (const [target, shiftKey, expected] of steps) {
      fireEvent.keyDown(target, { key: 'z', shiftKey, ...chord });
      await settle();
      expect(input(api(), 2, 1)).toBe(expected);
    }

    const bold = api().handle.selectionFormatting(0, 'B3:B3').bold;
    fireEvent.keyDown(view.getByTestId('panel-button'), { key: 'b', ...chord });
    await settle();
    expect(api().handle.selectionFormatting(0, 'B3:B3').bold).toBe(!bold);

    fireEvent.keyDown(view.getByTestId('panel-button'), { key: 'K', shiftKey: true, ...chord });
    fireEvent.keyDown(document.body.querySelector<HTMLElement>('[data-testid="portal-button"]')!, {
      key: 'K',
      shiftKey: true,
      ...chord,
    });
    await settle();
    expect(marks).toEqual(['acme.review', 'acme.review']);
  });

  test('retained geometry refuses in the same step as a committed change or a new frame', async () => {
    let retained: XlsxPluginGeometry | null = null;
    const duringChange: unknown[] = [];
    const contexts: XlsxPluginContext<null>[] = [];
    const plugin = defineXlsxPlugin<null>({
      id: 'acme.review',
      createState: () => null,
      onEvent(context, event) {
        contexts.push(context);
        if (event.type === 'document-change' && retained) {
          duringChange.push(retained.getCellRect({ sheetId: 'sheet:0', row: 2, col: 1 }));
        }
      },
      overlay: () => <div data-testid="layout-marker" />,
    });
    const { api, view } = await mount({ plugins: [plugin] });
    const b3 = { sheetId: 'sheet:0', row: 2, col: 1 };
    await until(() => contexts.some((context) => context.geometry !== null));
    retained = contexts.find((context) => context.geometry !== null)!.geometry!;
    expect(retained.getCellRect(b3)).not.toBeNull();

    api().handle.editCell(0, 2, 1, 'Direct');
    expect(retained.getCellRect(b3)).toBeNull();
    expect(
      retained.getRangeRect({ sheetId: 'sheet:0', range: { top: 0, left: 0, bottom: 1, right: 1 } })
    ).toBeNull();
    await until(() => duringChange.length > 0);
    expect(duringChange).toEqual([null]);

    await until(() => last(contexts).geometry?.getCellRect(b3) != null);
    const painted = last(contexts).geometry!;
    expect(painted.layout.version).toBe(api().handle.version());
    const scroll = view.getByTestId('xlsx-scroll');
    scroll.scrollTop = 40;
    fireEvent.scroll(scroll);
    await until(() => last(contexts).geometry?.layout.viewport.y === 40);
    expect(painted.getCellRect(b3)).toBeNull();
  });

  test('replacement during initialization, a queued read or an action refuses', async () => {
    const clipboard = holdClipboard();
    try {
      let release!: () => void;
      const gate = new Promise<void>((done) => (release = done));
      const log: string[] = [];
      const errors: XlsxPluginError[] = [];
      const contexts: XlsxPluginContext<null>[] = [];
      let first: XlsxPluginContext<null> | null = null;
      const plugin = defineXlsxPlugin<null>({
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
      const onPluginError = (error: XlsxPluginError) => errors.push(error);
      const { api, view, rerender } = await mount({ plugins: [plugin], onPluginError });
      await until(() => first !== null);
      rerender({ plugins: [plugin], onPluginError, file: fixture.slice() });
      await until(() => log.includes('load:replaced'));
      release();
      await until(() => log.some((entry) => entry.startsWith('read:')));
      expect(log).toEqual(['cleanup:document-replaced', 'load:replaced', 'read:document-replaced']);

      const second = last(contexts);
      await queuePaste(view, api());
      const reading = second.read.readCells({ ranges: [] });
      let acted: unknown = null;
      let proceed!: () => void;
      const waiting = new Promise<void>((done) => (proceed = done));
      const action = second.run(async (context) => {
        await waiting;
        acted = await context.read.version();
      });
      rerender({ plugins: [plugin], onPluginError, file: fixture.slice() });
      expect(await reading).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
      proceed();
      await action;
      expect(acted).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
      await until(() => log.filter((entry) => entry === 'load:replaced').length === 2);
      await act(async () => clipboard.resolve('late'));
      await settle();
      expect(errors).toEqual([]);
    } finally {
      clipboard.restore();
    }
  });

  test('queued writes check the grant and readOnly when they run, not when they were issued', async () => {
    const clipboard = holdClipboard();
    try {
      const { plugin, contexts, log } = recorder();
      const { api, view, rerender } = await mount({ plugins: [plugin], pluginGrants: WRITE });
      await until(() => log.includes('load:loaded'));
      const edits = last(contexts).edits!;

      await queuePaste(view, api());
      const revoked = edits.applyEdits(setCell(api().handle.version(), 'B4', 'Revoked'));
      rerender({ plugins: [plugin], pluginGrants: {} });
      await until(() => log.includes('grants-change'));
      await act(async () => clipboard.resolve('First'));
      expect(await revoked).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });

      rerender({ plugins: [plugin], pluginGrants: WRITE });
      await until(() => log.filter((entry) => entry === 'grants-change').length === 2);
      await queuePaste(view, api());
      const readOnly = last(contexts).edits!.applyEdits(
        setCell(api().handle.version(), 'B4', 'Read-only')
      );
      rerender({ plugins: [plugin], pluginGrants: WRITE, readOnly: true });
      await until(() => log.includes('mode-change:true'));
      await act(async () => clipboard.resolve('Second'));
      expect(await readOnly).toMatchObject({ ok: false, failure: { code: 'read-only' } });
      expect(input(api(), 2, 1)).toBe('First');
      expect(input(api(), 3, 1)).toBe('200');
    } finally {
      clipboard.restore();
    }
  });

  test('a deferred command completes with the grant and workbook current then', async () => {
    let store: XlsxCommandStore | null = null;
    function Panel() {
      store = useXlsxCommands();
      return null;
    }
    const grants = { 'acme.review': { commands: ['zoom'] } } as const;
    const { plugin, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'left', render: Panel },
    });
    const { api, rerender } = await mount({ plugins: [plugin], pluginGrants: grants });
    await until(() => store !== null && log.includes('load:loaded'));
    const revoked = xlsxCommandController(store!)!.prepare('zoom');
    rerender({ plugins: [plugin], pluginGrants: {} });
    await until(() => log.includes('grants-change'));
    expect(await revoked.execute({ scale: 2 })).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });

    rerender({ plugins: [plugin], pluginGrants: grants });
    await until(() => log.filter((entry) => entry === 'grants-change').length === 2);
    const replaced = xlsxCommandController(store!)!.prepare('zoom');
    rerender({ plugins: [plugin], pluginGrants: grants, file: fixture.slice() });
    await until(() => log.includes('load:replaced'));
    expect(await replaced.execute({ scale: 2 })).toMatchObject({ ok: false });
    expect(api().commands.getState('zoom').value).toBe(1);
  });

  test('a clearing batch across sheets publishes exactly one change', async () => {
    const { plugin, log } = recorder();
    const { api } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    let applied!: Awaited<ReturnType<XlsxEditorApi['applyEdits']>>;
    await act(async () => {
      applied = await api().applyEdits({
        expectVersion: api().handle.version(),
        steps: [setCell('', 'B3', '').steps[0], { ...setCell('', 'A1', '', 'sheet:1').steps[0] }],
      });
    });
    expect(applied).toMatchObject({ ok: true, applied: true });
    await settle(30);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toEqual([
      `document-change:${applied.ok ? applied.version : ''}`,
    ]);
  });

  test('registrations follow id and revision; duplicates and late cleanups are handled', async () => {
    const errors: XlsxPluginError[] = [];
    const a = recorder('acme.a');
    const b = recorder('acme.b');
    const onPluginError = (error: XlsxPluginError) => errors.push(error);
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

/** Bytes of `fixture` after `ops`, applied through the engine's operation escape hatch. */
function withOps(bytes: Uint8Array, ops: unknown[]): Uint8Array {
  const workbook = openWorkbook(bytes.slice());
  try {
    workbook.applyOps(ops);
    return workbook.save();
  } finally {
    workbook.dispose();
  }
}

describe('XlsxEditor plugin adapters', () => {
  test('public parts that need editor internals render nothing in a contribution', async () => {
    function Panel() {
      const bar = (
        <EditorToolbar mode="commands">
          <EditorToolbar.FormulaBar />
        </EditorToolbar>
      );
      return (
        <>
          {bar}
          {createPortal(<div data-testid="portal-bar">{bar}</div>, document.body)}
        </>
      );
    }
    const { plugin, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'right', render: Panel },
      overlay: () => <EditorToolbar.FormulaBar />,
    });
    const { api, view } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    await until(() => view.queryByTestId('plugin-dock-right') !== null);
    await settle();
    const inputs = document.body.querySelectorAll('[data-testid="xlsx-formula-input"]');
    expect(inputs).toHaveLength(1);
    expect(view.getByTestId('xlsx-toolbar').contains(inputs[0])).toBe(true);
    await act(async () => {
      api().selectCells(0, selectionAt({ row: 2, col: 1 }));
    });
    const formula = inputs[0] as HTMLInputElement;
    fireEvent.change(formula, { target: { value: 'Typed' } });
    fireEvent.keyDown(formula, { key: 'Enter' });
    expect(input(api(), 2, 1)).toBe('Typed');
  });

  test('a load hook that applies a batch every run loads once, then sees its change', async () => {
    const receipts: unknown[] = [];
    const seen: string[] = [];
    const { plugin } = recorder('acme.review', {
      async onEvent(context, event) {
        if (event.type === 'load' || event.type === 'document-change') {
          seen.push(`${event.type}:${event.version}`);
        }
        if (event.type !== 'load' || !context.edits) return;
        const applied = await context.edits.applyEdits(
          setCell(event.version, 'B3', `From load ${seen.length}`)
        );
        receipts.push(applied);
        if (applied.ok) context.setState({ count: receipts.length }, applied.version);
      },
      panel: {
        title: 'Loaded',
        placement: 'left',
        render: ({ context }) => <output data-testid="loaded">{context.state.count}</output>,
      },
    });
    const { api, view } = await mount({ plugins: [plugin], pluginGrants: WRITE });
    await until(() => within(view.container).queryByTestId('loaded')?.textContent === '1');
    await settle(150);
    const version = api().handle.version();
    const [load] = seen;
    expect(receipts).toMatchObject([
      { ok: true, applied: true, baseVersion: load.slice(5), version },
    ]);
    expect(input(api(), 2, 1)).toBe('From load 1');
    expect(seen).toEqual([load, `document-change:${version}`]);
  });

  test('a change the load hook did not make delivers load again at the latest version', async () => {
    const loads: string[] = [];
    const changes: string[] = [];
    let release!: () => void;
    const gate = new Promise<void>((done) => (release = done));
    const { plugin } = recorder('acme.review', {
      async onEvent(_context, event) {
        if (event.type === 'document-change') changes.push(event.version);
        if (event.type !== 'load') return;
        loads.push(event.version);
        if (loads.length === 1) await gate;
      },
    });
    const { api } = await mount({ plugins: [plugin] });
    await until(() => loads.length === 1);
    const before = api().handle.version();
    await act(async () => {
      await api().applyEdits(setCell(before, 'B3', 'External'));
    });
    release();
    await until(() => loads.length === 2);
    await settle(50);
    expect(loads).toEqual([before, api().handle.version()]);
    expect(loads[1]).not.toBe(before);
    expect(changes).toEqual([]);
  });

  test('defineXlsxPlugin keeps the contributions it was given when they change in place', () => {
    const shortcuts = ['Mod+Shift+R'];
    const commands = [
      {
        id: 'run',
        label: 'Run',
        mutatesDocument: false,
        shortcuts,
        execute: () => ({ ok: true, status: 'executed' }) as const,
      },
    ];
    const toolbar = ['run'];
    const panel = { title: 'Review', placement: 'left' as const, render: () => null };
    const plugin = defineXlsxPlugin<null>({
      id: 'acme.keys',
      createState: () => null,
      commands,
      toolbar,
      panel,
    });
    shortcuts.push('Mod+Shift+B');
    commands.push({ ...commands[0], id: 'late' });
    toolbar.length = 0;
    panel.title = 'Changed';
    const definition = pluginDefinition(plugin)!;
    expect(definition.commands!.map((command) => [command.id, command.shortcuts])).toEqual([
      ['run', ['Mod+Shift+R']],
    ]);
    expect(definition.toolbar).toEqual(['run']);
    expect(definition.panel!.title).toBe('Review');
    expect(Object.isFrozen(definition.commands![0].shortcuts)).toBe(true);
  });

  test('a workbook whose onReady throws never reaches plugins', async () => {
    const { plugin, log } = recorder('acme.review', {
      panel: { title: 'Review', placement: 'right', render: () => <p>panel</p> },
    });
    const view = render(
      <XlsxEditor
        file={fixture}
        plugins={[plugin]}
        onReady={() => {
          throw new Error('host setup failed');
        }}
      />
    );
    await until(() => view.queryByTestId('xlsx-error') !== null);
    await settle(50);
    expect(log).toEqual([]);
    expect(view.queryByTestId('plugin-dock-right')).toBeNull();
  });

  test('plugin shortcuts without Mod, Alt or a function key are invalid definitions', async () => {
    const errors: XlsxPluginError[] = [];
    const definition = (id: string, chord: string) =>
      recorder(id, {
        commands: [
          {
            id: 'mark',
            label: 'Mark',
            mutatesDocument: false,
            shortcuts: [chord],
            execute: () => ({ ok: true, status: 'executed' }),
          },
        ],
      });
    const deleting = definition('acme.delete', 'Delete');
    const entering = definition('acme.enter', 'Enter');
    await mount({
      plugins: [deleting.plugin, entering.plugin],
      onPluginError: (error) => errors.push(error),
    });
    await settle(50);
    expect(errors.map((error) => [error.pluginId, error.phase, String(error.error)])).toEqual([
      ['acme.delete', 'definition', 'TypeError: Command "mark" has an invalid shortcut'],
      ['acme.enter', 'definition', 'TypeError: Command "mark" has an invalid shortcut'],
    ]);
    expect([...deleting.log, ...entering.log]).toEqual([]);
  });

  test('plugin shortcuts the editor handles are reported; the others dispatch', async () => {
    const errors: XlsxPluginError[] = [];
    const runs: string[] = [];
    const { plugin, log } = recorder('acme.review', {
      commands: [
        {
          id: 'mark',
          label: 'Mark',
          mutatesDocument: false,
          shortcuts: ['Mod+V', 'Alt+ArrowDown', 'F2', 'Mod+Shift+R', 'F7'],
          execute(context) {
            runs.push(context.pluginId);
            return { ok: true, status: 'executed' };
          },
        },
      ],
    });
    const { api, view } = await mount({
      plugins: [plugin],
      onPluginError: (error) => errors.push(error),
    });
    await until(() => log.includes('load:loaded'));
    await settle();
    const descriptor = api().commands.getDescriptor('plugin:acme.review/mark');
    expect(descriptor?.shortcuts.map((shortcut) => shortcut.chord)).toEqual(['Mod+Shift+R', 'F7']);
    expect(errors.map((error) => String(error.error))).toEqual([
      expect.stringMatching(/Mod\+V .* the grid/),
      expect.stringMatching(/Alt\+ArrowDown .* the grid/),
      expect.stringMatching(/F2 .* the grid/),
    ]);
    const surface = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(surface, { key: 'R', shiftKey: true, ...chord });
    fireEvent.keyDown(surface, { key: 'F7' });
    await settle();
    expect(runs).toEqual(['acme.review', 'acme.review']);
  });

  test('a reveal deferred to the next frame is dropped once the user switches sheets', async () => {
    const { plugin, contexts, log } = recorder();
    const { api, view } = await mount({ plugins: [plugin] });
    await until(() => log.includes('load:loaded'));
    const scroll = view.getByTestId('xlsx-scroll');
    const navigated = await act(async () => {
      const result = await last(contexts).navigation.scrollToCell(
        { sheetId: 'sheet:0', row: 60, col: 0 },
        { expectVersion: api().handle.version(), align: 'start' }
      );
      fireEvent.click(sheetTabs(view)[1]);
      return result;
    });
    expect(navigated).toEqual({ ok: true });
    await settle(50);
    expect(sheetTabs(view)[1].getAttribute('aria-selected')).toBe('true');
    expect(scroll.scrollTop).toBe(0);
  });

  test('plugin reads and batches land after cell and formula drafts and a composition', async () => {
    const { plugin, contexts, log } = recorder();
    const { api, view } = await mount({ plugins: [plugin], pluginGrants: WRITE });
    await until(() => log.includes('load:loaded'));
    const context = () => last(contexts);
    const b3 = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'B3' } };
    await act(async () => {
      api().selectCells(0, selectionAt({ row: 2, col: 1 }));
    });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'Cell draft' } });
    const read = await act(() => context().read.readCells({ ranges: [b3] }));
    expect(read).toMatchObject({ ok: true, version: api().handle.version() });
    if (read.ok) expect(read.ranges[0].cells[0][0].displayText).toBe('Cell draft');
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();

    const before = api().handle.version();
    fireEvent.change(view.getByTestId('xlsx-formula-input'), {
      target: { value: 'Formula draft' },
    });
    expect(
      await act(() => context().edits!.applyEdits(setCell(before, 'B3', 'Batch')))
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(input(api(), 2, 1)).toBe('Formula draft');

    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const cell = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    act(() => cell.focus());
    fireEvent.compositionStart(cell);
    fireEvent.change(cell, { target: { value: '日本' } });
    let settled = false;
    const composing = context()
      .read.readCells({ ranges: [b3] })
      .then((result) => {
        settled = true;
        return result;
      });
    await settle(50);
    expect(settled).toBe(false);
    fireEvent.change(cell, { target: { value: '日本語' } });
    const composed = await act(async () => {
      fireEvent.compositionEnd(cell);
      return composing;
    });
    expect(composed).toMatchObject({ ok: true, version: api().handle.version() });
    if (composed.ok) expect(composed.ranges[0].cells[0][0].displayText).toBe('日本語');
  });

  test('plugin reads wait for a chart nudge and refuse during a chart drag', async () => {
    const [chart] = openWorkbook(charted.slice()).displayList({ x: 0, y: 0, ...VIEWPORT }).charts!;
    const center = {
      clientX: chart.clip.x + chart.clip.w / 2,
      clientY: chart.clip.y + chart.clip.h / 2,
    };
    const { plugin, contexts, log } = recorder();
    const { api, view } = await mount({ plugins: [plugin], file: charted });
    await until(() => log.includes('load:loaded'));
    const surface = view.getByTestId('xlsx-scroll');
    fireEvent.mouseDown(surface, center);
    expect(await act(() => last(contexts).read.version())).toMatchObject({
      ok: false,
      failure: { code: 'input-failed' },
    });
    fireEvent.mouseUp(window, center);

    const before = api().handle.version();
    await act(async () => {
      fireEvent.keyDown(surface, { key: 'ArrowRight' });
    });
    expect(api().handle.version()).toBe(before);
    const read = await act(() => last(contexts).read.version());
    expect(read).toMatchObject({ ok: true });
    if (read.ok) expect(read.version).not.toBe(before);
    expect(read).toMatchObject({ version: api().handle.version() });
  });

  test('geometry follows variable and hidden tracks and addresses merged areas as ranges', async () => {
    const file = withOps(fixture, [
      { type: 'setColWidth', sheet: 0, col: 0, width: 8 },
      { type: 'setColWidth', sheet: 0, col: 1, width: 30 },
      { type: 'setRowHeight', sheet: 0, row: 3, height: 0 },
      { type: 'setRowHeight', sheet: 0, row: 4, height: 40 },
      {
        type: 'mergeCells',
        sheet: 0,
        range: { start: { row: 5, col: 1 }, end: { row: 6, col: 2 } },
      },
    ]);
    const contexts: XlsxPluginContext<null>[] = [];
    const plugin = defineXlsxPlugin<null>({
      id: 'acme.geometry',
      createState: () => null,
      onEvent(context) {
        contexts.push(context);
      },
    });
    const { api } = await mount({ plugins: [plugin], file });
    const geometry = () => last(contexts)?.geometry ?? null;
    await until(() => geometry()?.layout.zoom === 1);
    for (const zoom of [1, 2]) {
      if (zoom !== 1) {
        await act(async () => {
          await api().commands.execute('zoom', { scale: zoom });
        });
        await until(() => geometry()?.layout.zoom === zoom);
      }
      const cell = (row: number, col: number) =>
        geometry()!.getCellRect({ sheetId: 'sheet:0', row, col });
      const frame = api().handle.displayList(geometry()!.layout.viewport).grid!;
      const painted = cellRect(frame, 2, 1)!;
      expect(cell(2, 1)).toEqual({
        x: painted.x * zoom,
        y: painted.y * zoom,
        width: painted.w * zoom,
        height: painted.h * zoom,
      });
      expect(cell(2, 1)!.width).toBeGreaterThan(cell(2, 2)!.width * 1.5);
      expect(cell(3, 1)).toBeNull();
      expect(cell(4, 1)!.y).toBeCloseTo(cell(2, 1)!.y + cell(2, 1)!.height, 5);
      expect(cell(4, 1)!.height).toBeCloseTo(((40 * 4) / 3) * zoom, 1);
      const merged = geometry()!.getRangeRect({
        sheetId: 'sheet:0',
        range: { top: 5, left: 1, bottom: 6, right: 2 },
      })!;
      expect(merged).toEqual({
        x: cell(5, 1)!.x,
        y: cell(5, 1)!.y,
        width: cell(5, 1)!.width + cell(5, 2)!.width,
        height: cell(5, 1)!.height + cell(6, 1)!.height,
      });
      expect(cell(5, 1)!.width).toBeLessThan(merged.width);
    }
  });
});
