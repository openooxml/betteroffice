import { expect, test } from 'bun:test';
import type { PresentationHandle } from '@betteroffice/pptx';
import { createPptxPluginHost } from './createPptxPluginHost';
import { definePptxPlugin } from './definePptxPlugin';
import type { PptxPluginEvent, PptxPluginGeometry, PptxPluginLayout } from './types';

const settle = () => new Promise((done) => setTimeout(done, 0));
function setupSignals() {
  const document = { version: () => 'v1', onUpdate: () => () => {} } as unknown as PresentationHandle;
  let geometry: PptxPluginGeometry | null = null;
  const host = createPptxPluginHost({
    handle: () => document,
    admit: async (_handle, operation) => ({ ok: true, value: operation() }),
    readOnlyRefusal: () => null,
    applyEdits: () => { throw new Error('Unavailable'); },
    commands: () => null,
    navigator: { goToSlide: () => null, selectShape: () => null, selectText: () => null },
    geometry: () => geometry,
    translate: (key) => key,
  });
  return { host, open: () => host.open(document), setGeometry: (next: PptxPluginGeometry | null) => { geometry = next; } };
}

test('focus and locale are deduplicated in snapshots and events', async () => {
  const { host, open } = setupSignals();
  const events: PptxPluginEvent[] = [];
  host.localeChanged('de');
  host.setPlugins([definePptxPlugin({ id: 'signals', createState: () => null,
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

test('viewport changes compare values and clear when layout disappears', async () => {
  let geometry: PptxPluginGeometry | null = null;
  const { host, open, setGeometry } = setupSignals();
  const events: PptxPluginEvent[] = [];
  host.setPlugins([definePptxPlugin({ id: 'viewport', createState: () => null,
    onEvent(_context, event) { events.push(event); } })]);
  open();
  await settle();
  const layout: PptxPluginLayout = { id: 'layout', version: 'v1', slideId: 'slide', slide: 1, zoom: 1, width: 1280, height: 720 };
  const rect = { x: 0, y: 0, width: 600, height: 400 };
  geometry = { layout, viewport: rect } as PptxPluginGeometry;
  setGeometry(geometry);
  host.layoutChanged(layout);
  host.geometryChanged();
  await settle();
  events.length = 0;
  host.geometryChanged();
  geometry.viewport = { ...rect };
  host.geometryChanged();
  await settle();
  expect(events).toEqual([]);
  geometry.viewport.y = 80;
  host.geometryChanged();
  host.geometryChanged();
  await settle();
  expect(events).toEqual([{ type: 'viewport-change', generation: host.generation(),
    viewport: { ...rect, y: 80 } }]);
  host.layoutChanged(null);
  host.geometryChanged();
  await settle();
  expect(events.at(-1)).toEqual({ type: 'viewport-change', generation: host.generation(), viewport: null });
  expect(events.filter((event) => event.type === 'viewport-change')).toHaveLength(2);
  host.close('unmounted');
});

test('a superseded generation receives no queued signals', async () => {
  const { host, open, setGeometry } = setupSignals();
  const events: PptxPluginEvent[] = [];
  let release!: () => void;
  const pending = new Promise<void>((done) => { release = done; });
  let first = true;
  let signal: AbortSignal | undefined;
  host.setPlugins([definePptxPlugin({ id: 'superseded', createState: () => null,
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
  const layout: PptxPluginLayout = { id: 'layout', version: 'v1', slideId: 'slide', slide: 1, zoom: 1, width: 1280, height: 720 };
  const geometry = { layout, viewport: { x: 0, y: 0, width: 600, height: 400 } } as PptxPluginGeometry;
  setGeometry(geometry);
  host.layoutChanged(layout);
  host.geometryChanged();
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
