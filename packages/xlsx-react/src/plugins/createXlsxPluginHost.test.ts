import { expect, test } from 'bun:test';
import type { WorkbookHandle } from '@betteroffice/xlsx';
import { createXlsxPluginHost } from './createXlsxPluginHost';
import { defineXlsxPlugin } from './defineXlsxPlugin';
import type { XlsxPluginEvent } from './types';

const settle = () => new Promise((done) => setTimeout(done, 0));
function setupSignals() {
  const document = { version: () => 'v1', onUpdate: () => () => {} } as unknown as WorkbookHandle;
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
  return { host, open: () => host.open(document) };
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
