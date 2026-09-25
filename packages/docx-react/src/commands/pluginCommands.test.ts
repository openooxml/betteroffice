import { describe, expect, test } from 'bun:test';
import {
  createDocxCommandController,
  docxCommandController,
  type DocxCommandScope,
  type DocxPluginCommandBinding,
} from './createDocxCommandStore';
import { BUILT_IN_CHORDS, isPluginCommandId, normalizeChord } from './descriptors';
import { testBinding } from './testing';
import type { DocxCommandFailureCode, DocxCommandId, DocxPluginCommandId } from './types';

function setup(initial = {}) {
  const harness = testBinding(initial);
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  return { harness, controller, store: controller.store };
}

function scope(granted: readonly string[], pluginId = 'acme') {
  const listeners = new Set<() => void>();
  const state = { granted: new Set(granted), ended: null as DocxCommandFailureCode | null };
  const value: DocxCommandScope = {
    deny(id) {
      if (state.ended) return state.ended;
      if (isPluginCommandId(id))
        return id.startsWith(`plugin:${pluginId}/`) ? null : 'permission-denied';
      return state.granted.has(id) ? null : 'permission-denied';
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
  return {
    value,
    state,
    changed: () => {
      for (const listener of listeners) listener();
    },
  };
}

function contributed(
  id: DocxPluginCommandId,
  mutatesDocument: boolean,
  runs: string[]
): DocxPluginCommandBinding {
  return {
    descriptor: {
      id,
      label: id,
      mutatesDocument,
      shortcuts: [{ chord: 'Mod+Shift+M', args: null }],
    },
    state: () => ({ enabled: true, active: runs.length > 0 }),
    execute: async () => {
      runs.push(id);
      return { ok: true, status: 'executed' };
    },
  };
}

describe('scoped command stores', () => {
  test('refuse ungranted commands and mutating built-ins without a policy path', async () => {
    const { store, controller, harness } = setup();
    const plugin = scope(['reviewNext', 'bold']);
    const scoped = controller.scoped(plugin.value);
    expect(controller.scoped(plugin.value)).toBe(scoped);
    expect(docxCommandController(scoped)?.store).toBe(scoped);
    expect(scoped.getState('italic').disabledReason?.code).toBe('permission-denied');
    expect(scoped.getState('bold').disabledReason?.code).toBe('unsupported-policy');
    expect(scoped.getState('bold')).toBe(scoped.getState('bold'));
    expect(store.getState('bold').enabled).toBe(true);
    expect(await scoped.execute('italic', null)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
    expect(await scoped.execute('bold', null)).toMatchObject({
      ok: false,
      failure: { code: 'unsupported-policy' },
    });
    expect(await scoped.execute('reviewNext', null)).toEqual({ ok: true, status: 'executed' });
    expect(harness.calls.map((call) => call.id)).toEqual(['reviewNext']);

    harness.update({ readOnly: true });
    controller.refresh();
    expect(scoped.getState('bold').disabledReason?.code).toBe('read-only');
  });

  test('check the grant again once the command is admitted', async () => {
    const { controller, harness } = setup();
    const plugin = scope(['reviewNext', 'reviewPrevious']);
    const scoped = controller.scoped(plugin.value);
    let admit!: () => void;
    harness.state.admission = () => new Promise<void>((resolve) => (admit = resolve));
    const pending = scoped.execute('reviewNext', null);
    plugin.state.granted.delete('reviewNext');
    plugin.changed();
    admit();
    expect(await pending).toMatchObject({ ok: false, failure: { code: 'permission-denied' } });
    const ended = scoped.execute('reviewPrevious', null);
    plugin.state.ended = 'document-replaced';
    admit();
    expect(await ended).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
    expect(scoped.getState('reviewPrevious').disabledReason?.code).toBe('plugin-unavailable');
    expect(harness.calls).toEqual([]);
  });
});

describe('contributed commands', () => {
  test('are registered under their namespace and gated by the editor', async () => {
    const { store, controller, harness } = setup();
    const runs: string[] = [];
    const mark = 'plugin:acme/mark' as const;
    const stamp = 'plugin:acme/stamp' as const;
    controller.setPluginCommands(
      [contributed(mark, false, runs), contributed(stamp, true, runs)],
      [stamp, 'plugin:gone/x', mark]
    );
    expect(controller.pluginToolbar()).toEqual([stamp, mark]);
    expect(controller.pluginShortcuts().map((binding) => binding.id)).toEqual([mark, stamp]);
    expect(store.getDescriptor(mark)).toMatchObject({ label: mark, mutatesDocument: false });
    expect(store.getDescriptor('plugin:acme/none')).toBeNull();
    expect(store.getState('plugin:acme/none').disabledReason?.code).toBe('unsupported-command');

    harness.state.admission = () => new Promise<void>(() => {});
    expect(await store.execute(mark, null)).toEqual({ ok: true, status: 'executed' });
    expect(store.getState(mark).active).toBe(true);

    harness.update({ mode: 'viewing' });
    controller.refresh();
    expect(store.getState(stamp).disabledReason?.code).toBe('viewing-mode');
    expect(store.getState(mark).enabled).toBe(true);
    expect(await store.execute(stamp, null)).toMatchObject({
      ok: false,
      failure: { code: 'viewing-mode' },
    });
    expect(runs).toEqual([mark]);
  });

  test('run only from their own plugin scope', async () => {
    const { controller } = setup();
    const runs: string[] = [];
    const mark = 'plugin:acme/mark' as const;
    controller.setPluginCommands([contributed(mark, false, runs)], [mark]);
    const own = controller.scoped(scope([], 'acme').value);
    const other = controller.scoped(scope([], 'other').value);
    expect(await other.execute(mark, null)).toMatchObject({
      ok: false,
      failure: { code: 'permission-denied' },
    });
    expect(await own.execute(mark, null)).toEqual({ ok: true, status: 'executed' });
    expect(runs).toEqual([mark]);
  });

  test('chords normalize, and built-in chords are reserved', () => {
    expect(normalizeChord('Shift+Mod+b')).toBe('Mod+Shift+b');
    expect(normalizeChord('Ctrl+B')).toBeNull();
    expect(normalizeChord('Mod++')).toBe('Mod++');
    expect(BUILT_IN_CHORDS.has('Mod+b')).toBe(true);
    expect(BUILT_IN_CHORDS.has(normalizeChord('Mod+Shift+M')!)).toBe(false);
    const ids: unknown[] = ['plugin:acme/mark', 'plugin:acme/', 'plugin:/x', 'bold'];
    expect(ids.map(isPluginCommandId)).toEqual([true, false, false, false]);
    const builtIn: DocxCommandId = 'bold';
    expect(isPluginCommandId(builtIn)).toBe(false);
  });
});
