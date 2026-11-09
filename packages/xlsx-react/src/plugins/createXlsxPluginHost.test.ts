import { expect, test } from 'bun:test';
import type { WorkbookHandle } from '@betteroffice/xlsx';
import { createXlsxPluginHost } from './createXlsxPluginHost';
import { defineXlsxPlugin } from './defineXlsxPlugin';
import type { XlsxPluginEvent } from './types';

const settle = () => new Promise((done) => setTimeout(done, 0));
function setupSignals() {
  let version = 'v1';
  const updates = new Set<() => void>();
  const document = { version: () => version, onUpdate: (listener: () => void) => {
    updates.add(listener); return () => { updates.delete(listener); };
  } } as unknown as WorkbookHandle;
  const host = createXlsxPluginHost({
    handle: () => document,
    admit: async <T>(_handle: WorkbookHandle, operation: () => T) => ({ ok: true as const, value: operation() }),
    readOnlyRefusal: () => null,
    applyEdits: () => { throw new Error('Unavailable'); },
    commands: () => null,
    navigator: { selectCells: () => null, scrollToCell: () => null },
    geometry: () => null,
    translate: (key) => key,
  });
  return { host, open: () => host.open(document), change(next: string) {
    version = next;
    for (const listener of updates) listener();
  } };
}

test('focus and locale are deduplicated in snapshots and events', async () => {
  const { host, open } = setupSignals();
  const events: XlsxPluginEvent[] = [];
  host.localeChanged('de');
  host.setPlugins([defineXlsxPlugin({ id: 'signals', createState: () => null,
    onEvent(_context, event) {
      events.push(event);
    } })]);
  open();
  await settle();
  expect(host.activations()[0]!.context.snapshot).toMatchObject({ focused: false, locale: 'de' });
  expect(events.map((event) => event.type)).toEqual(['load']);
  host.focusChanged(true);
  host.focusChanged(true);
  host.localeChanged('fr');
  host.localeChanged('fr');
  await settle();
  expect(events.filter((event) => event.type === 'focus-change')).toEqual([
    { type: 'focus-change', generation: host.generation()!, focused: true },
  ]);
  expect(events.filter((event) => event.type === 'locale-change')).toEqual([
    { type: 'locale-change', generation: host.generation()!, locale: 'fr' },
  ]);
  expect(host.activations()[0]!.context.snapshot).toMatchObject({ focused: true, locale: 'fr' });
  host.close('unmounted');
});

test('a superseded generation receives no queued signals', async () => {
  const { host, open } = setupSignals();
  const events: XlsxPluginEvent[] = [];
  let release!: () => void;
  const pending = new Promise<void>((done) => { release = done; });
  let first = true;
  let signal: AbortSignal | undefined;
  host.setPlugins([defineXlsxPlugin({ id: 'superseded', createState: () => null,
    initialize(context) {
      if (!first) return;
      first = false;
      signal = context.lifetimeSignal;
      return pending;
    },
    onEvent(_context, event) { events.push(event); } })]);
  open();
  const superseded = host.generation();
  host.focusChanged(true);
  host.localeChanged('de');
  host.selectionUnsettled();
  host.selectionSettled('keyboard');
  open();
  release();
  await settle();
  expect(signal?.aborted).toBe(true);
  expect(events.some((event) => event.generation === superseded)).toBe(false);
  expect(host.activations()[0]!.context.snapshot).toMatchObject({ focused: true, locale: 'de' });
  host.focusChanged(false);
  await settle();
  expect(events[events.length - 1]).toEqual({ type: 'focus-change', generation: host.generation()!, focused: false });
  host.close('unmounted');
});


test('selection settling publishes the current snapshot once and ignores superseded generations', async () => {
  const { host, open } = setupSignals();
  const events: XlsxPluginEvent[] = [];
  host.setPlugins([defineXlsxPlugin({ id: 'selection-settle', createState: () => null,
    onEvent(_context, event) { events.push(event); } })]);
  open();
  await settle();
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  const selection = { sheetId: 'sheet:0', sheetIndex: 0, cells: { anchor: { row: 0, col: 0 }, focus: { row: 2, col: 2 } }, chartId: null };
  host.selectionUnsettled();
  host.selectionChanged(selection);
  await settle();
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(false);
  host.selectionSettled('pointer');
  host.selectionSettled('pointer');
  await settle();
  expect(events.filter((event) => event.type === 'selection-settle')).toEqual([
    { type: 'selection-settle', generation: host.generation()!, version: host.version(), selection, source: 'pointer' },
  ]);
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  host.selectionUnsettled();
  host.selectionSettled('keyboard');
  await settle();
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
  const previous = host.generation();
  open();
  host.selectionUnsettled(previous);
  host.selectionSettled('touch', previous);
  await settle();
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
  host.close('unmounted');
});


for (const changed of ['selection', 'version', 'selection-returned', 'selection-mutated', 'unchanged'] as const) {
  test(`queued selection settles are checked at dispatch (${changed})`, async () => {
    const { host, open, change } = setupSignals();
    let entered!: () => void;
    let release!: () => void;
    let finished!: () => void;
    const started = new Promise<void>((resolve) => { entered = resolve; });
    const held = new Promise<void>((resolve) => { release = resolve; });
    const drained = new Promise<void>((resolve) => { finished = resolve; });
    const events: XlsxPluginEvent[] = [];
    const delivered: Array<{ event: XlsxPluginEvent; version: string; selection: unknown }> = [];
    host.setPlugins([defineXlsxPlugin({ id: 'queued-settle', createState: () => null,
      async onEvent(context, event) {
        if (event.type === 'mode-change') { entered(); await held; }
        if (event.type === 'selection-settle') {
          delivered.push({ event, version: context.snapshot.version, selection: context.snapshot.selection });
        }
        events.push(event);
        if (event.type === 'locale-change') finished();
      } })]);
    open();
    await settle();
    const first = { sheetId: 'sheet:0', sheetIndex: 0, cells: { anchor: { row: 0, col: 0 }, focus: { row: 2, col: 2 } }, chartId: null };
    const second = { ...first, cells: { anchor: { row: 1, col: 1 }, focus: { row: 3, col: 3 } } };
    host.modeChanged(true);
    await started;
    try {
      host.selectionChanged(first);
      host.selectionUnsettled();
      host.selectionSettled('pointer');
      if (changed === 'selection' || changed === 'selection-returned') host.selectionChanged(second);
      if (changed === 'selection-returned') host.selectionChanged(first);
      if (changed === 'selection-mutated') first.cells.focus.row = 3;
      if (changed === 'version') change('v2');
      host.localeChanged('de');
      release();
      await drained;
      const settled = events.filter((event) => event.type === 'selection-settle');
      expect(settled).toHaveLength(changed === 'unchanged' ? 1 : 0);
      for (const item of delivered) if (item.event.type === 'selection-settle') {
        expect(item.event.version).toBe(item.version);
        expect(item.event.selection).toEqual(item.selection);
      }
      if (changed !== 'unchanged' && changed !== 'selection-returned') {
        host.selectionUnsettled();
        host.selectionSettled('keyboard');
        host.localeChanged('fr');
        await settle();
        expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
      }
    } finally { release(); host.close('unmounted'); }
  });
}
