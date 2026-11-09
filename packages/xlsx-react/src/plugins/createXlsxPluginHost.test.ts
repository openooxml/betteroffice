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
    admit: async (_handle, operation) => ({ ok: true, value: operation() }),
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
    { type: 'focus-change', generation: host.generation(), focused: true },
  ]);
  expect(events.filter((event) => event.type === 'locale-change')).toEqual([
    { type: 'locale-change', generation: host.generation(), locale: 'fr' },
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
  open();
  release();
  await settle();
  expect(signal?.aborted).toBe(true);
  expect(events.some((event) => event.generation === superseded)).toBe(false);
  expect(host.activations()[0]!.context.snapshot).toMatchObject({ focused: true, locale: 'de' });
  host.focusChanged(false);
  await settle();
  expect(events.at(-1)).toEqual({ type: 'focus-change', generation: host.generation(), focused: false });
  host.close('unmounted');
});
