import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { describe, expect, spyOn, test } from 'bun:test';
import * as pluginRuntime from '../../../../shared/plugin-host/runtime';
import type { ResidentProposalReply, YrsSession } from '@betteroffice/docx/yrs';
import { selectionClock } from '../../../../e2e/browser/plugin-host/selection-clock';
import { createSelectionSettle } from '../../../../shared/plugin-host/selectionSettle';
import { registerWorkerProposalAuthority } from '../components/DocxEditor/internals/workerProposalAuthority';
import { createDocxPluginHost, type DocxPluginHostAccess } from './createDocxPluginHost';
import { defineDocxPlugin } from './defineDocxPlugin';
import { docxSelectionIdentity } from './selectionIdentity';
import type { DocxProposalSnapshot } from './proposalPreview';
import type { DocxAnchorGeometryResult, DocxPluginEvent, DocxPluginGeometry, DocxPluginLayout, DocxPluginSelection } from './types';

function stubSession(previewVersion = 0) {
  const updates = new Set<() => void>();
  const proposals = new Set<(snapshot: DocxProposalSnapshot) => void>();
  let snapshot: DocxProposalSnapshot = { version: 'v1', previewVersion, proposals: [] };
  const session = {
    version: () => snapshot.version,
    getProposals: () => snapshot,
    onUpdate(listener: () => void) {
      updates.add(listener);
      return () => {
        updates.delete(listener);
      };
    },
    onProposalChange(listener: (snapshot: DocxProposalSnapshot) => void) {
      proposals.add(listener);
      return () => {
        proposals.delete(listener);
      };
    },
  } as unknown as YrsSession;
  return {
    session,
    updates,
    proposals,
    change(next: Partial<DocxProposalSnapshot>) {
      snapshot = { ...snapshot, ...next };
      for (const listener of proposals) listener(snapshot);
    },
  };
}

function stubGeometry() {
  const listeners = new Set<() => void>();
  let subscriptions = 0;
  return {
    listeners,
    subscriptions: () => subscriptions,
    subscribeGeometry(listener: () => void) {
      subscriptions += 1;
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    change() {
      for (const listener of listeners) listener();
    },
  };
}

function setup(
  geometry: () => DocxPluginGeometry | null = () => null,
  subscribeGeometry?: DocxPluginHostAccess['subscribeGeometry'],
  selectionPublished?: DocxPluginHostAccess['selectionPublished'],
  extra: Partial<DocxPluginHostAccess> = {}
) {
  const events: DocxPluginEvent[] = [];
  const host = createDocxPluginHost({
    pagedEditorRef: { current: null },
    writeMode: () => 'viewing',
    commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }),
    subscribeLayout: () => () => {},
    geometry,
    subscribeGeometry,
    selectionPublished,
    translate: (key) => key,
    ...extra,
  });
  const plugin = defineDocxPlugin({
    id: 'test.anchors',
    createState: () => null,
    onEvent(_context, event) {
      events.push(event);
    },
  });
  return { host, plugin, events };
}

const layout = (previewVersion: number): DocxPluginLayout => ({
  id: 'layout',
  version: 'v1',
  previewVersion,
  reservationsVersion: null,
  zoom: 1,
  pageCount: 1,
});
const settle = () => new Promise((done) => setTimeout(done, 0));

for (const mode of ['session', 'worker'] as const) {
  test(`${mode} contexts bind synchronous and pending geometry to their snapshot selection`, async () => {
    const frame = layout(0);
    let x = 1;
    let pending: Promise<void> | null = null;
    let release!: () => void;
    const answer = (): DocxAnchorGeometryResult => ({ ok: true, version: 'v1', previewVersion: 0,
      layoutId: frame.id, rects: [{ pageIndex: 0, x, y: 0, width: 3, height: 12 }],
      anchor: { pageIndex: 0, x: x + 3, y: 0, width: 0, height: 12 },
      pageRect: { x: 0, y: 0, width: 100, height: 200 } });
    const geometry = { layout: frame,
      getAnchorGeometry: answer,
      readAnchorGeometries: async (targets) => { if (pending) await pending; return targets.map(answer); },
    } as DocxPluginGeometry;
    const { host, plugin } = setup(() => geometry);
    host.setPlugins([plugin]);
    if (mode === 'session') host.open(stubSession().session);
    else host.openWorker('v1');
    const first = { formatting: null,
      displayRange: { story: 'body', from: 1, to: 4, layoutId: frame.id } };
    host.layoutChanged(frame);
    host.selectionChanged(first);
    await host.notificationsIdle();
    const retained = host.activations()[0]!.context;
    expect(retained.geometry!.getAnchorGeometry({ kind: 'selection' })).toMatchObject({ ok: true, rects: [{ x: 1 }] });
    pending = new Promise<void>((resolve) => { release = resolve; });
    const targets = [{ kind: 'selection' }, { kind: 'range', version: 'v1', range: {
      story: 'body', start: { paraId: 'p', offset: 0 }, end: { paraId: 'p', offset: 3 }, view: 'accepted',
    } }] as const;
    const read = retained.geometry!.readAnchorGeometries(targets);
    x = 8;
    host.selectionChanged({ ...first, displayRange: { ...first.displayRange, from: 8, to: 11 } });
    await host.notificationsIdle();
    const current = host.activations()[0]!.context;
    expect(current.snapshot.version).toBe(retained.snapshot.version);
    expect(current.snapshot.layout).toEqual(retained.snapshot.layout);
    expect(current.geometry!.getAnchorGeometry({ kind: 'selection' })).toMatchObject({ ok: true, rects: [{ x: 8 }] });
    expect(retained.snapshot.selection).toEqual(first);
    expect(retained.geometry!.getAnchorGeometry({ kind: 'selection' }))
      .toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
    expect(retained.geometry!.getAnchorGeometry(targets[1])).toMatchObject({ ok: true });
    release();
    expect(await read).toMatchObject([
      { ok: false, failure: { code: 'layout-unavailable' } }, { ok: true },
    ]);
    expect(await retained.geometry!.readAnchorGeometries([{ kind: 'selection' }]))
      .toMatchObject([{ ok: false, failure: { code: 'layout-unavailable' } }]);
    host.close('unmounted');
  });
}

for (const mode of ['session', 'worker'] as const) {
  describe(`plugin host ${mode} geometry subscriptions`, () => {
    function setupGeometryHost() {
      const geometry = stubGeometry();
      const { host } = setup(() => null, geometry.subscribeGeometry);
      const open = () => {
        if (mode === 'worker') host.openWorker('v1');
        else host.open(stubSession().session);
      };
      return { host, geometry, open };
    }

    test('detaches on unmount and ignores later geometry changes', async () => {
      const { host, geometry, open } = setupGeometryHost();
      let notifications = 0;
      host.subscribe(() => notifications += 1);
      open();
      expect(geometry.listeners.size).toBe(1);
      host.close('unmounted');
      await Promise.resolve();
      notifications = 0;
      geometry.change();
      await Promise.resolve();
      expect(notifications).toBe(0);
      expect(geometry.listeners.size).toBe(0);
    });

    test('reattaches once when the same host reopens after unmount', async () => {
      const { host, geometry, open } = setupGeometryHost();
      let notifications = 0;
      host.subscribe(() => notifications += 1);
      open();
      host.close('unmounted');
      expect(geometry.listeners.size).toBe(0);
      open();
      expect(geometry.listeners.size).toBe(1);
      expect(geometry.subscriptions()).toBe(2);
      await Promise.resolve();
      notifications = 0;
      geometry.change();
      await Promise.resolve();
      expect(notifications).toBe(1);
      host.close('unmounted');
      expect(geometry.listeners.size).toBe(0);
    });

    test('keeps one subscription across document replacement and repeated opens', async () => {
      const { host, geometry, open } = setupGeometryHost();
      let notifications = 0;
      host.subscribe(() => notifications += 1);
      open();
      host.close('document-replaced');
      expect(geometry.listeners.size).toBe(1);
      open();
      open();
      expect(geometry.listeners.size).toBe(1);
      expect(geometry.subscriptions()).toBe(1);
      await Promise.resolve();
      notifications = 0;
      geometry.change();
      await Promise.resolve();
      expect(notifications).toBe(1);
      host.close('unmounted');
      expect(geometry.listeners.size).toBe(0);
    });
  });
}

describe('plugin host proposal previews', () => {
  test('a proposal event handler can commit a worker read and observe records at the same preview', async () => {
    const registry = stubSession();
    Object.assign(registry.session, { encodeStateVector: () => new Uint8Array() });
    let snapshot: DocxProposalSnapshot = { version: 'v1', previewVersion: 0, proposals: [] };
    const authority = registerWorkerProposalAuthority(registry.session, {
      async proposal(): Promise<ResidentProposalReply> {
        return {
          mirror: { version: snapshot.version, proposals: { previewVersion: snapshot.previewVersion,
            entries: snapshot.proposals.map(record => ({ key: record.id, record,
              suggest: { author: 'Reviewer', date: '2026-10-01T00:00:00Z' } })) } },
          geometry: { version: snapshot.version, previewVersion: snapshot.previewVersion,
            proposals: '', targets: {}, hidden: [] },
          changedStories: [], updates: [], stateVector: new Uint8Array(),
        };
      },
      documentRead: async () => { throw new Error('Unexpected document read'); },
      handOver: async () => { throw new Error('Unexpected hand-over'); },
    }, { editorPeer: true, current: () => true, laidOut: async () => {},
      adopted: () => {}, relayout: () => {}, contentChanged: () => {} });
    await authority.initialize();
    const { host } = setup();
    let finish = (_committed: boolean) => {};
    const nextCommit = () => new Promise<boolean>(resolve => { finish = resolve; });
    let events = 0;
    let flushes = 0;
    const admission = { flush: async () => { flushes += 1; }, validate: () => {} };
    const main = async (): Promise<DocxProposalSnapshot> => { throw new Error('Unexpected replica read'); };
    const plugin = defineDocxPlugin<DocxProposalSnapshot | null>({
      id: 'test.worker-proposal-read', createState: () => null,
      async onEvent(context, event) {
        if (event.type !== 'proposal-change') return;
        events += 1;
        const current = await authority.getProposals(main, admission);
        finish(context.setState(current, current.version));
      },
    });
    try {
      host.setPlugins([plugin]);
      host.open(registry.session);
      await settle();
      snapshot = { ...snapshot, previewVersion: 1 };
      let committed = nextCommit();
      await authority.getProposals(main);
      expect(await committed).toBe(true);
      expect(host.activations()[0]!.context.state).toEqual(snapshot);
      expect(events).toBe(1);
      expect(flushes).toBe(1);
      snapshot = { ...snapshot, proposals: [{ id: 'proposal', state: 'accepted',
        paragraph: { kind: 'session', sessionId: 'session', story: 'body', paraId: 'paragraph' },
        revisionIds: ['revision'], changed: true }] };
      committed = nextCommit();
      await authority.getProposals(main);
      expect(await committed).toBe(true);
      expect(host.activations()[0]!.context.state).toEqual(snapshot);
      expect(events).toBe(2);
      expect(flushes).toBe(2);
    } finally {
      host.close('unmounted');
    }
  });

  test('publishes preview events and clears stale layout ids synchronously', async () => {
    const { host, plugin, events } = setup();
    const registry = stubSession();
    host.setPlugins([plugin]);
    host.open(registry.session);
    await settle();
    host.layoutChanged(layout(0));
    await settle();
    events.length = 0;
    registry.change({ previewVersion: 1 });
    expect(host.previewVersion()).toBe(1);
    expect(host.layoutId()).toBeNull();
    await settle();
    expect(events).toEqual([
      { type: 'proposal-change', generation: host.generation()!, version: 'v1', previewVersion: 1 },
      { type: 'layout-change', generation: host.generation()!, layout: null },
    ]);
    host.layoutChanged(layout(0));
    expect(host.layoutId()).toBeNull();
    host.layoutChanged(layout(1));
    expect(host.layoutId()).toBe('layout');
    await settle();
    expect(events.at(-1)).toMatchObject({ type: 'layout-change', layout: { previewVersion: 1 } });
    host.close('unmounted');
  });

  test('notifies registrations at the same preview without invalidating current pixels', async () => {
    const { host, plugin, events } = setup();
    const registry = stubSession(3);
    host.setPlugins([plugin]);
    host.open(registry.session);
    expect(host.previewVersion()).toBe(3);
    host.layoutChanged(layout(3));
    await settle();
    events.length = 0;
    registry.change({
      proposals: [
        {
          id: 'proposal',
          state: 'accepted',
          paragraph: { kind: 'session', sessionId: 'session', story: 'body', paraId: 'paragraph' },
          revisionIds: [],
          changed: false,
        },
      ],
    });
    expect(host.layoutId()).toBe('layout');
    await settle();
    expect(events).toEqual([
      {
        type: 'proposal-change',
        generation: host.generation()!,
        version: 'v1',
        previewVersion: 3,
      },
    ]);
    host.close('unmounted');
  });

  test('a worker proposal version change publishes document-change before proposal-change without an update', async () => {
    const { host, plugin, events } = setup();
    const registry = stubSession();
    host.setPlugins([plugin]);
    host.open(registry.session);
    host.layoutChanged(layout(0));
    await settle();
    events.length = 0;
    registry.change({ version: 'v2', previewVersion: 1 });
    expect(host.version()).toBe('v2');
    expect(host.previewVersion()).toBe(1);
    expect(host.layoutId()).toBeNull();
    await settle();
    expect(events).toEqual([
      { type: 'document-change', generation: host.generation()!, version: 'v2' },
      { type: 'layout-change', generation: host.generation()!, layout: null },
      { type: 'proposal-change', generation: host.generation()!, version: 'v2', previewVersion: 1 },
    ]);
    for (const listener of registry.updates) listener();
    await settle();
    expect(events.filter((event) => event.type === 'document-change')).toHaveLength(1);
    host.close('unmounted');
  });

  test('repeats the current layout once its pixels are presented, and nothing for another one', async () => {
    const { host, plugin, events } = setup();
    const registry = stubSession();
    host.setPlugins([plugin]);
    host.open(registry.session);
    await settle();
    host.layoutChanged(layout(0));
    await settle();
    events.length = 0;
    host.layoutPresented({ ...layout(0), id: 'older' });
    host.layoutPresented(layout(0));
    await settle();
    expect(events).toEqual([
      { type: 'layout-change', generation: host.generation()!, layout: layout(0) },
    ]);
    host.close('unmounted');
  });

  test('repeats the current layout once its geometry exists, unless a layout change carried it', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const { host, plugin, events } = setup(() => geometry);
    const registry = stubSession();
    host.setPlugins([plugin]);
    host.open(registry.session);
    await settle();
    host.layoutChanged(layout(0));
    await settle();
    events.length = 0;
    host.geometryPresented(layout(0));
    geometry = { layout: layout(0) } as DocxPluginGeometry;
    host.geometryPresented({ ...layout(0), id: 'older' });
    await settle();
    expect(events).toEqual([]);
    host.geometryPresented(layout(0));
    await settle();
    host.geometryPresented(layout(0));
    await settle();
    expect(events).toEqual([
      { type: 'layout-change', generation: host.generation()!, layout: layout(0) },
    ]);
    const next = { ...layout(0), id: 'next' };
    geometry = { layout: next } as DocxPluginGeometry;
    events.length = 0;
    host.layoutChanged(next);
    await settle();
    host.geometryPresented(next);
    await settle();
    expect(events).toEqual([
      { type: 'layout-change', generation: host.generation()!, layout: next },
    ]);
    geometry = null;
    host.layoutChanged(null);
    host.layoutChanged(next);
    await settle();
    events.length = 0;
    geometry = { layout: next } as DocxPluginGeometry;
    host.geometryPresented(next);
    await settle();
    expect(events).toEqual([
      { type: 'layout-change', generation: host.generation()!, layout: next },
    ]);
    host.close('unmounted');
  });

  test('subscribes only with installed plugins and detaches on removal, replacement and close', () => {
    const { host, plugin } = setup();
    const first = stubSession();
    const second = stubSession(7);
    host.open(first.session);
    expect(first.updates.size).toBe(0);
    expect(first.proposals.size).toBe(0);
    first.change({ previewVersion: 2 });
    host.setPlugins([plugin]);
    expect(first.updates.size).toBe(1);
    expect(first.proposals.size).toBe(1);
    expect(host.previewVersion()).toBe(2);
    host.setPlugins([plugin]);
    expect(first.proposals.size).toBe(1);
    host.setPlugins([]);
    expect(first.updates.size).toBe(0);
    expect(first.proposals.size).toBe(0);
    host.setPlugins([plugin]);
    host.open(second.session);
    expect(first.updates.size).toBe(0);
    expect(first.proposals.size).toBe(0);
    expect(second.updates.size).toBe(1);
    expect(second.proposals.size).toBe(1);
    expect(host.previewVersion()).toBe(7);
    first.change({ previewVersion: 99 });
    expect(host.previewVersion()).toBe(7);
    host.close('document-replaced');
    expect(second.updates.size).toBe(0);
    expect(second.proposals.size).toBe(0);
    host.open(second.session);
    host.close('unmounted');
    expect(second.proposals.size).toBe(0);
  });

  test('defaults to preview zero without a registry and still observes document changes', async () => {
    const { host, plugin, events } = setup();
    const registry = stubSession();
    Reflect.deleteProperty(registry.session, 'getProposals');
    Reflect.deleteProperty(registry.session, 'onProposalChange');
    host.setPlugins([plugin]);
    host.open(registry.session);
    expect(host.previewVersion()).toBe(0);
    host.layoutChanged(layout(0));
    await settle();
    events.length = 0;
    registry.session.version = () => 'v2';
    for (const listener of registry.updates) listener();
    expect(host.layoutId()).toBeNull();
    await settle();
    expect(events).toEqual([
      { type: 'document-change', generation: host.generation()!, version: 'v2' },
      { type: 'layout-change', generation: host.generation()!, layout: null },
    ]);
    host.close('unmounted');
  });
});


test('notification idle waits for running handlers and every queued event', async () => {
  const { host } = setup();
  const registry = stubSession();
  const delivered: string[] = [];
  let release!: () => void;
  let started!: () => void;
  const entering = new Promise<void>((resolve) => { started = resolve; });
  const held = new Promise<void>((resolve) => { release = resolve; });
  const plugin = defineDocxPlugin({
    id: 'test.drain', createState: () => null,
    async onEvent(_context, event) {
      if (event.type === 'mode-change') { started(); await held; }
      delivered.push(event.type);
    },
  });
  host.setPlugins([plugin]);
  host.open(registry.session);
  await host.notificationsIdle();
  delivered.length = 0;
  host.modeChanged('viewing', true);
  await entering;
  host.selectionChanged({ formatting: null, displayRange: { story: 'body', from: 1, to: 2, layoutId: 'layout' } });
  let idle = false;
  const drained = host.notificationsIdle().then(() => { idle = true; });
  await Promise.resolve();
  expect(idle).toBe(false);
  expect(delivered).toEqual([]);
  release();
  await drained;
  expect(delivered).toEqual(['mode-change', 'selection-change']);
  expect(idle).toBe(true);
  host.close('unmounted');
  await host.notificationsIdle();
});

test('focus and locale are deduplicated in snapshots and events', async () => {
  const { host } = setup();
  const open = () => host.openWorker('v1');
  const events: DocxPluginEvent[] = [];
  host.localeChanged('de');
  host.setPlugins([defineDocxPlugin({ id: 'signals', createState: () => null,
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

test('viewport changes compare values and clear when layout disappears', async () => {
  let geometry: DocxPluginGeometry | null = null;
  const { host } = setup(() => geometry);
  const open = () => host.openWorker('v1');
  const events: DocxPluginEvent[] = [];
  host.setPlugins([defineDocxPlugin({ id: 'viewport', createState: () => null,
    onEvent(_context, event) { events.push(event); } })]);
  open();
  await settle();
  const layout: DocxPluginLayout = { id: 'layout', version: 'v1', previewVersion: 0, zoom: 1, pageCount: 1 };
  const rect = { x: 0, y: 0, width: 600, height: 400 };
  geometry = { layout, viewport: rect } as DocxPluginGeometry;
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
  expect(events).toEqual([{ type: 'viewport-change', generation: host.generation()!,
    viewport: { ...rect, y: 80 } }]);
  host.layoutChanged(null);
  host.geometryChanged();
  await settle();
  expect(events.at(-1)).toEqual({ type: 'viewport-change', generation: host.generation()!, viewport: null });
  expect(events.filter((event) => event.type === 'viewport-change')).toHaveLength(2);
  host.close('unmounted');
});

test('a superseded generation receives no queued signals', async () => {
  let geometry: DocxPluginGeometry | null = null;
  const { host } = setup(() => geometry);
  const open = () => host.openWorker('v1');
  const events: DocxPluginEvent[] = [];
  let release!: () => void;
  const pending = new Promise<void>((done) => { release = done; });
  let first = true;
  let signal: AbortSignal | undefined;
  host.setPlugins([defineDocxPlugin({ id: 'superseded', createState: () => null,
    initialize(context) {
      if (!first) return;
      first = false;
      signal = context.lifetimeSignal;
      return pending;
    },
    onEvent(_context, event) { events.push(event); } })]);
  open();
  const superseded = host.generation();
  host.selectionUnsettled();
  host.selectionSettled('keyboard');
  host.focusChanged(true);
  host.localeChanged('de');
  const layout: DocxPluginLayout = { id: 'layout', version: 'v1', previewVersion: 0, zoom: 1, pageCount: 1 };
  geometry = { layout, viewport: { x: 0, y: 0, width: 600, height: 400 } } as DocxPluginGeometry;
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
  expect(events.at(-1)).toEqual({ type: 'focus-change', generation: host.generation()!, focused: false });
  host.close('unmounted');
});


test('only emitted resolved changes notify the settle machine once', async () => {
  const keys: string[] = [];
  const { host, plugin, events } = setup(undefined, undefined, key => keys.push(key));
  host.setPlugins([plugin]);
  host.openWorker('v1');
  await host.notificationsIdle();
  const empty = { formatting: null, displayRange: null };
  const publish = async (offset: number | null, cause?: 'remote' | 'seed') => {
    host.selectionChanged(empty, docxSelectionIdentity(0, offset === null ? null : ['sticky-a', 'sticky-b'],
      offset === null ? null : ['fn:1', offset, offset + 2], false), { owner: host, cause });
    await host.notificationsIdle();
    expect(keys).toHaveLength(events.filter(event => event.type === 'selection-change').length);
  };
  await publish(1);
  await publish(1);
  expect(keys).toHaveLength(1);
  await publish(4, 'remote');
  await publish(1);
  await publish(null);
  await publish(1);
  await publish(2, 'seed');
  await publish(3);
  expect(keys).toHaveLength(7);
  expect(keys[1]).toBe(keys[0]);
  expect(keys[2]).not.toBe(keys[1]);
  expect(keys[3]).not.toBe(keys[2]);
  expect(keys[4]).not.toBe(keys[3]);
  expect(keys[5]).toBe(keys[4]);
  expect(keys[6]).not.toBe(keys[5]);
  host.close('unmounted');
});

function queuedSelectionHost(sourceCause?: 'remote' | 'seed') {
  let offset = 1;
  let ready = true;
  let reads = 0;
  const causes: Array<'remote' | 'seed' | undefined> = [];
  const keys: string[] = [];
  const frame = layout(0);
  const answer = () => ({ ok: true as const, version: 'v1', previewVersion: 0, layoutId: frame.id,
    rects: [], anchor: null, pageRect: null });
  const geometry = { layout: frame, getAnchorGeometry: answer,
    readAnchorGeometries: async () => [answer()] } as unknown as DocxPluginGeometry;
  const owner = {};
  const h = setup(() => geometry, undefined, key => keys.push(key), {
    selectionCause: () => sourceCause,
    readSelection(cause) {
      reads += 1;
      causes.push(cause);
      return { selection: { formatting: null,
        displayRange: ready ? { story: 'body', from: offset, to: offset, layoutId: frame.id } : null },
      identity: docxSelectionIdentity(0, ['sticky', offset], ['body', offset, 'body', offset, 0], true), owner, ready };
    },
  });
  h.host.setPlugins([h.plugin]);
  h.host.openWorker('v1');
  h.host.layoutChanged(frame);
  return { ...h, keys, causes, reads: () => reads, mapped: (next: boolean) => { ready = next; },
    move: (next: number) => { offset = next; } };
}

for (const marks of [['remote', 'remote'], ['remote', undefined], [undefined, 'remote'],
  ['remote', 'refresh'], ['refresh', 'remote'], [undefined, 'refresh'], ['refresh', undefined]] as const) {
  test(`dirty marks ${JSON.stringify(marks)} pull the final payload once and preserve the contributing cause`, async () => {
    const h = queuedSelectionHost();
    await h.host.notificationsIdle();
    h.host.markSelectionDirty(marks[0]);
    h.move(9);
    h.host.markSelectionDirty(marks[1]);
    expect(h.reads()).toBe(0);
    await Promise.resolve();
    await h.host.notificationsIdle();
    expect(h.reads()).toBe(1);
    expect(h.causes).toEqual([marks.some(cause => cause === 'remote') && !marks.some(cause => cause === undefined) ? 'remote' : undefined]);
    expect(h.events.filter(event => event.type === 'selection-change')).toEqual([{
      type: 'selection-change', generation: h.host.generation()!, version: 'v1',
      selection: { formatting: null, displayRange: { story: 'body', from: 9, to: 9, layoutId: 'layout' } },
    }]);
    expect(h.keys).toHaveLength(1);
    h.host.markSelectionDirty();
    await h.host.notificationsIdle();
    expect(h.keys).toHaveLength(1);
    h.host.close('unmounted');
  });
}

for (const marks of [['remote', 'seed'], ['seed', 'remote']] as const) {
  test(`remote and seed marks ${JSON.stringify(marks)} coalesce without a local revision or settle`, async () => {
    const ownsDom = !GlobalRegistrator.isRegistered;
    if (ownsDom) GlobalRegistrator.register();
    const clock = selectionClock();
    const h = queuedSelectionHost();
    const machine = createSelectionSettle({ root: () => null,
      onUnsettle: () => h.host.selectionUnsettled(),
      onSettle: source => h.host.selectionSettled(source),
      onResettle: source => h.host.selectionSettled(source),
    });
    const publish = h.host.selectionChanged;
    h.host.selectionChanged = (selection, identity, publication) => {
      const key = publish(selection, identity, publication);
      if (h.keys.length) machine.notifySelectionChange(h.keys.at(-1)!);
      return key;
    };
    try {
      h.host.markSelectionDirty();
      await h.host.notificationsIdle();
      await clock.tick(350);
      await h.host.notificationsIdle();
      const key = h.keys.at(-1);
      h.events.length = 0;
      h.causes.length = 0;
      h.host.markSelectionDirty(marks[0]);
      h.move(9);
      h.host.markSelectionDirty(marks[1]);
      await h.host.notificationsIdle();
      expect(h.causes).toEqual(['remote']);
      expect(h.keys.at(-1)).toBe(key);
      expect(h.events.filter(event => event.type === 'selection-change')).toHaveLength(1);
      expect(h.host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
      await clock.tick(700);
      await h.host.notificationsIdle();
      expect(h.events.filter(event => event.type === 'selection-settle')).toEqual([]);
    } finally {
      machine.dispose();
      h.host.close('unmounted');
      clock.restore();
      if (ownsDom) await GlobalRegistrator.unregister();
    }
  });
}

for (const read of ['snapshot', 'guard'] as const) {
  test(`${read} flushes a reentrant layout and selection without nested snapshots or duplicate events`, async () => {
    const createRuntime = pluginRuntime.createPluginRuntime;
    let snapshots = 0;
    let depth = 0;
    let maxDepth = 0;
    let armSnapshot = false;
    let mark = () => {};
    const capture = spyOn(pluginRuntime, 'createPluginRuntime').mockImplementation(options =>
      createRuntime({ ...options, snapshot(pluginId, generation) {
        snapshots += 1;
        depth += 1;
        maxDepth = Math.max(maxDepth, depth);
        try {
          if (armSnapshot) { armSnapshot = false; mark(); }
          return options.snapshot(pluginId, generation);
        } finally { depth -= 1; }
      } }));
    const frame = { ...layout(0), id: 'reentrant-layout' };
    let pulls = 0;
    const observed: Array<{ layout: DocxPluginLayout | null; selection: DocxPluginSelection }> = [];
    const h = setup(undefined, undefined, undefined, { readSelection() {
      pulls += 1;
      h.host.layoutChanged(frame);
      return { selection: { formatting: null,
        displayRange: { story: 'body', from: 8, to: 8, layoutId: frame.id } }, identity: 'caret-8' };
    } });
    const plugin = defineDocxPlugin({ id: 'reentrant', createState: () => null,
      onEvent(context, event) {
        h.events.push(event);
        if (event.type === 'selection-change' || event.type === 'layout-change') {
          observed.push({ layout: context.snapshot.layout, selection: context.snapshot.selection });
        }
      } });
    try {
      h.host.setPlugins([plugin]);
      h.host.openWorker('v1');
      await h.host.notificationsIdle();
      h.events.length = 0;
      snapshots = 0;
      maxDepth = 0;
      mark = () => h.host.markSelectionDirty();
      if (read === 'snapshot') {
        armSnapshot = true;
        h.host.focusChanged(true);
      } else {
        mark();
        h.host.guard(plugin.id, 'command', context => {
          observed.push({ layout: context.snapshot.layout, selection: context.snapshot.selection });
        }, undefined);
      }
      await h.host.notificationsIdle();
      expect(pulls).toBe(1);
      expect(snapshots).toBeGreaterThan(0);
      expect(maxDepth).toBe(1);
      expect(observed.length).toBeGreaterThanOrEqual(2);
      for (const snapshot of observed) {
        expect(snapshot.layout).toEqual(frame);
        expect(snapshot.selection.displayRange).toEqual({ story: 'body', from: 8, to: 8, layoutId: frame.id });
      }
      expect(h.events.filter(event => event.type === 'layout-change')).toHaveLength(1);
      expect(h.events.filter(event => event.type === 'selection-change')).toHaveLength(1);
    } finally { h.host.close('unmounted'); capture.mockRestore(); }
  });
}

test('remote and local callers coalesce intermediate formatting into one complete local publication and settle', async () => {
  const ownsDom = !GlobalRegistrator.isRegistered;
  if (ownsDom) GlobalRegistrator.register();
  const clock = selectionClock();
  const intermediate: DocxPluginSelection = {
    formatting: { hasSelection: false, isMultiParagraph: false,
      textFormatting: { bold: false }, paragraphFormatting: {}, styleId: null,
      startParagraphIndex: 0, endParagraphIndex: 0 },
    displayRange: { story: 'body', from: 1, to: 1, layoutId: 'layout' },
  };
  const final: DocxPluginSelection = {
    formatting: { hasSelection: true, isMultiParagraph: false,
      textFormatting: { bold: true, italic: true, fontSize: 36, color: { rgb: '123456' } },
      paragraphFormatting: { alignment: 'center', spaceAfter: 120 }, styleId: 'Heading1',
      startParagraphIndex: 1, endParagraphIndex: 1 },
    displayRange: { story: 'hf:header', from: 3, to: 9, layoutId: 'layout' },
  };
  let current = intermediate;
  let machine: ReturnType<typeof createSelectionSettle> | undefined;
  const owner = {};
  const causes: Array<'remote' | 'seed' | undefined> = [];
  const { host, plugin, events } = setup(undefined, undefined,
    key => machine?.notifySelectionChange(key), {
      readSelection() {
        const range = current.displayRange!;
        return { selection: current, identity: docxSelectionIdentity(0,
          [range.story, range.from, range.to], [range.story, range.from, range.story, range.to, 1],
          range.from === range.to), owner };
      },
    });
  const publish = host.selectionChanged;
  host.selectionChanged = (selection, identity, publication) => {
    causes.push(publication?.cause);
    return publish(selection, identity, publication);
  };
  machine = createSelectionSettle({ root: () => null,
    onUnsettle: () => host.selectionUnsettled(),
    onSettle: source => host.selectionSettled(source),
    onResettle: source => host.selectionSettled(source),
  });
  try {
    host.setPlugins([plugin]);
    host.openWorker('v1');
    await host.notificationsIdle();
    const remoteCaller = (selection: DocxPluginSelection) => {
      current = selection;
      host.markSelectionDirty('remote');
    };
    const localCaller = (selection: DocxPluginSelection) => {
      current = selection;
      host.markSelectionDirty();
    };
    remoteCaller(intermediate);
    localCaller(final);
    await Promise.resolve();
    await host.notificationsIdle();
    const changes = events.filter(event => event.type === 'selection-change');
    expect(changes).toHaveLength(1);
    expect(causes).toEqual([undefined]);
    expect(changes[0]!.selection).toEqual(final);
    await clock.tick(349);
    await host.notificationsIdle();
    expect(events.filter(event => event.type === 'selection-settle')).toEqual([]);
    await clock.tick(1);
    await host.notificationsIdle();
    expect(events.filter(event => event.type === 'selection-settle')).toEqual([{
      type: 'selection-settle', generation: host.generation()!, version: 'v1',
      selection: final, source: 'keyboard',
    }]);
    await clock.tick(700);
    await host.notificationsIdle();
    expect(events.filter(event => event.type === 'selection-change' || event.type === 'selection-settle')
      .map(event => event.type)).toEqual(['selection-change', 'selection-settle']);
  } finally {
    machine.dispose();
    host.close('unmounted');
    clock.restore();
    if (ownsDom) await GlobalRegistrator.unregister();
  }
});

for (const local of [false, true]) test(`neutral worker refreshes read remote attribution at flush (local: ${local})`, async () => {
  const h = queuedSelectionHost('remote');
  await h.host.notificationsIdle();
  h.host.markSelectionDirty('refresh');
  h.move(9);
  if (local) h.host.markSelectionDirty();
  h.host.markSelectionDirty('refresh');
  await h.host.notificationsIdle();
  expect(h.causes).toEqual([local ? undefined : 'remote']);
  expect(h.events.filter(event => event.type === 'selection-change')).toHaveLength(1);
  expect(h.keys).toHaveLength(1);
  h.host.close('unmounted');
});

for (const operation of ['invocation', 'getAnchorGeometry', 'readAnchorGeometries'] as const) {
  test(`${operation} flushes a pending selection synchronously`, async () => {
    const h = queuedSelectionHost();
    h.host.markSelectionDirty();
    await h.host.notificationsIdle();
    const retained = h.host.activations()[0]!.context;
    h.move(6);
    h.host.markSelectionDirty();
    if (operation === 'invocation') {
      let selection: DocxPluginSelection | undefined;
      h.host.guard(h.plugin.id, 'command', context => { selection = context.snapshot.selection; }, undefined);
      expect(selection?.displayRange).toEqual({ story: 'body', from: 6, to: 6, layoutId: 'layout' });
    } else if (operation === 'getAnchorGeometry') {
      expect(retained.geometry!.getAnchorGeometry({ kind: 'selection' })).toMatchObject({
        ok: false, failure: { code: 'layout-unavailable' },
      });
    } else {
      const pending = retained.geometry!.readAnchorGeometries([{ kind: 'selection' }]);
      expect(h.reads()).toBe(2);
      expect(await pending).toMatchObject([{ ok: false, failure: { code: 'layout-unavailable' } }]);
    }
    expect(h.reads()).toBe(2);
    await Promise.resolve();
    await h.host.notificationsIdle();
    expect(h.reads()).toBe(2);
    expect(h.keys).toHaveLength(2);
    h.host.close('unmounted');
  });
}

test('worker mapping publishes the complete caret payload once and retains a local contributing cause', async () => {
  const h = queuedSelectionHost();
  h.host.markSelectionDirty();
  await h.host.notificationsIdle();
  h.events.length = 0;
  h.keys.length = 0;
  h.causes.length = 0;
  h.mapped(false);
  h.move(7);
  h.host.markSelectionDirty('remote');
  h.move(8);
  h.host.markSelectionDirty();
  await h.host.notificationsIdle();
  expect(h.events.filter(event => event.type === 'selection-change')).toEqual([]);
  expect(h.keys).toEqual([]);
  let range: DocxPluginSelection['displayRange'] | undefined;
  h.host.guard(h.plugin.id, 'command', context => { range = context.snapshot.selection.displayRange; }, undefined);
  expect(range).toBeNull();
  expect(h.host.activations()[0]!.context.snapshot.selectionSettled).toBe(false);
  h.mapped(true);
  h.host.markSelectionDirty('remote');
  await h.host.notificationsIdle();
  expect(h.causes).toEqual([undefined, undefined]);
  expect(h.events.filter(event => event.type === 'selection-change')).toEqual([{
    type: 'selection-change', generation: h.host.generation()!, version: 'v1',
    selection: { formatting: null, displayRange: { story: 'body', from: 8, to: 8, layoutId: 'layout' } },
  }]);
  expect(h.keys).toHaveLength(1);
  h.host.selectionSettled('keyboard');
  await h.host.notificationsIdle();
  h.host.markSelectionDirty();
  await h.host.notificationsIdle();
  expect(h.events.filter(event => event.type === 'selection-change')).toHaveLength(1);
  expect(h.keys).toHaveLength(1);
  h.host.close('unmounted');
});

test('worker mapping retains remote attribution through neutral view and geometry refreshes', async () => {
  const h = queuedSelectionHost();
  h.host.markSelectionDirty();
  await h.host.notificationsIdle();
  const settledKey = h.keys[0];
  h.events.length = 0;
  h.keys.length = 0;
  h.causes.length = 0;
  h.mapped(false);
  h.move(9);
  h.host.markSelectionDirty('remote');
  h.host.markSelectionDirty('refresh');
  await h.host.notificationsIdle();
  expect(h.events.filter(event => event.type === 'selection-change')).toEqual([]);
  expect(h.host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  h.mapped(true);
  h.host.markSelectionDirty('refresh');
  await h.host.notificationsIdle();
  expect(h.causes).toEqual(['remote', 'remote']);
  expect(h.keys).toEqual([settledKey]);
  expect(h.events.filter(event => event.type === 'selection-change')).toHaveLength(1);
  expect(h.host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  h.host.close('unmounted');
});

for (const lifecycle of ['unmount', 'document replacement', 'worker replacement', 'plugin removal'] as const) {
  test(`${lifecycle} cancels the pending selection flush`, async () => {
    const h = queuedSelectionHost();
    await h.host.notificationsIdle();
    h.host.markSelectionDirty();
    if (lifecycle === 'unmount') h.host.close('unmounted');
    else if (lifecycle === 'document replacement') h.host.open(stubSession().session);
    else if (lifecycle === 'worker replacement') h.host.openWorker('v2');
    else h.host.setPlugins([]);
    await Promise.resolve();
    await h.host.notificationsIdle();
    expect(h.reads()).toBe(0);
    expect(h.keys).toEqual([]);
    expect(h.events.filter(event => event.type === 'selection-change')).toEqual([]);
    if (lifecycle === 'plugin removal') h.host.setPlugins([h.plugin]);
    await h.host.notificationsIdle();
    expect(h.reads()).toBe(0);
    h.host.close('unmounted');
    h.host.markSelectionDirty();
    await Promise.resolve();
    expect(h.reads()).toBe(0);
  });
}

test('empty resolved publications clear stale formatting and display ranges without duplicate changes', async () => {
  const { host, plugin, events } = setup();
  host.setPlugins([plugin]);
  host.openWorker('v1');
  const selected = { formatting: { hasSelection: true, isMultiParagraph: false,
    textFormatting: { bold: true }, paragraphFormatting: {}, styleId: null,
    startParagraphIndex: 0, endParagraphIndex: 0 },
    displayRange: { story: 'body', from: 1, to: 4, layoutId: 'layout' } };
  host.selectionChanged(selected, docxSelectionIdentity(0, ['sticky'], ['body', 1, 'body', 4, 1], false), { owner: host });
  host.selectionUnsettled();
  host.selectionSettled('keyboard');
  await host.notificationsIdle();
  events.length = 0;
  const cleared = docxSelectionIdentity(0, null, null, false);
  host.selectionChanged(selected, cleared, { owner: host });
  host.selectionUnsettled();
  await host.notificationsIdle();
  host.selectionChanged({ ...selected, displayRange: null }, cleared, { owner: host });
  await host.notificationsIdle();
  host.selectionChanged({ formatting: null, displayRange: null }, cleared, { owner: host });
  await host.notificationsIdle();
  const changes = events.filter(event => event.type === 'selection-change');
  expect(changes).toHaveLength(1);
  expect(changes[0]!.selection).toEqual({ formatting: null, displayRange: null });
  host.selectionSettled('keyboard');
  await host.notificationsIdle();
  host.selectionSettled('keyboard');
  await host.notificationsIdle();
  const settled = events.filter(event => event.type === 'selection-settle');
  expect(settled).toHaveLength(1);
  expect(settled[0]!.selection).toEqual({ formatting: null, displayRange: null });
  expect(host.activations()[0]!.context.snapshot.selection).toEqual({ formatting: null, displayRange: null });
  host.close('unmounted');
});

test.each(['remote', 'seed'] as const)('silent %s movement rebases dedupe before a local return to the last settled position', async (cause) => {
  const { host } = setup();
  const events: DocxPluginEvent[] = [];
  host.setPlugins([defineDocxPlugin({ id: 'remote-return', createState: () => null,
    onEvent(_context, event) { events.push(event); } })]);
  host.openWorker('v1');
  const publish = (offset: number, remote = false) => {
    const identity = docxSelectionIdentity(0, ['sticky-a', 'sticky-b'], [offset, offset + 2], false);
    host.selectionChanged({ formatting: null, displayRange: { story: 'body', from: offset, to: offset + 2, layoutId: 'layout' } },
      identity, { owner: host, cause: remote ? cause : undefined });
  };
  publish(1);
  host.selectionUnsettled();
  host.selectionSettled('keyboard');
  await host.notificationsIdle();
  publish(4, true);
  await host.notificationsIdle();
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  expect(events.filter(event => event.type === 'selection-settle')).toHaveLength(1);
  publish(1);
  host.selectionUnsettled();
  host.selectionSettled('keyboard');
  await host.notificationsIdle();
  expect(events.filter(event => event.type === 'selection-settle')).toHaveLength(2);
  host.close('unmounted');
});

test('selection settling publishes the current snapshot once and ignores superseded generations', async () => {
  const { host } = setup();
  const open = () => host.openWorker('v1');
  const events: DocxPluginEvent[] = [];
  host.setPlugins([defineDocxPlugin({ id: 'selection-settle', createState: () => null,
    onEvent(_context, event) { events.push(event); } })]);
  open();
  await settle();
  expect(host.activations()[0]!.context.snapshot.selectionSettled).toBe(true);
  const selection = { formatting: null, displayRange: { story: 'body', from: 1, to: 5, layoutId: 'layout' } };
  host.selectionUnsettled();
  host.selectionChanged(selection);
  host.workerChanged('v2');
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
    const { host } = setup();
    const open = () => host.openWorker('v1');
    let entered!: () => void;
    let release!: () => void;
    let finished!: () => void;
    const started = new Promise<void>((resolve) => { entered = resolve; });
    const held = new Promise<void>((resolve) => { release = resolve; });
    const drained = new Promise<void>((resolve) => { finished = resolve; });
    const events: DocxPluginEvent[] = [];
    const delivered: Array<{ event: DocxPluginEvent; version: string; selection: unknown }> = [];
    host.setPlugins([defineDocxPlugin({ id: 'queued-settle', createState: () => null,
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
    const first = { formatting: null, displayRange: { story: 'body', from: 1, to: 5, layoutId: 'layout' } };
    const second = { formatting: null, displayRange: { story: 'body', from: 2, to: 6, layoutId: 'layout' } };
    host.modeChanged('viewing', true);
    await started;
    try {
      host.selectionChanged(first);
      host.selectionUnsettled();
      host.selectionSettled('pointer');
      if (changed === 'selection' || changed === 'selection-returned') host.selectionChanged(second);
      if (changed === 'selection-returned') host.selectionChanged(first);
      if (changed === 'selection-mutated') first.displayRange.to = 6;
      if (changed === 'version') host.workerChanged('v2');
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
