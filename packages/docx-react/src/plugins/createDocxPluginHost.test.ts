import { describe, expect, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { createDocxPluginHost } from './createDocxPluginHost';
import { defineDocxPlugin } from './defineDocxPlugin';
import type { DocxProposalSnapshot } from './proposalPreview';
import type { DocxPluginEvent, DocxPluginLayout } from './types';

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

function setup() {
  const events: DocxPluginEvent[] = [];
  const host = createDocxPluginHost({
    pagedEditorRef: { current: null },
    writeMode: () => 'viewing',
    commands: () => null,
    layout: () => ({ queries: null, complete: false, failed: false }),
    subscribeLayout: () => () => {},
    geometry: () => null,
    translate: (key) => key,
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
  zoom: 1,
  pageCount: 1,
});
const settle = () => new Promise((done) => setTimeout(done, 0));

describe('plugin host proposal previews', () => {
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
