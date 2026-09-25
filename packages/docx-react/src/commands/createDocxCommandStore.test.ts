import { describe, expect, test } from 'bun:test';
import { en } from '@betteroffice/docx-i18n';
import {
  createDocxCommandController,
  DocxCommandAdmissionError,
  UNAVAILABLE_DOCX_COMMANDS,
} from './createDocxCommandStore';
import { DOCX_COMMAND_DESCRIPTORS, DOCX_COMMAND_IDS, formatChord, matchesChord } from './descriptors';
import { PLAIN_CONTEXT, testBinding } from './testing';

function lookup(key: string): unknown {
  return key.split('.').reduce<unknown>(
    (node, part) => (node && typeof node === 'object' ? (node as Record<string, unknown>)[part] : undefined),
    en
  );
}

function setup(initial = {}) {
  const harness = testBinding(initial);
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  let notifications = 0;
  const unsubscribe = controller.store.subscribe(() => {
    notifications += 1;
  });
  return {
    harness,
    controller,
    store: controller.store,
    notifications: () => notifications,
    unsubscribe,
  };
}

describe('command store snapshots', () => {
  test('keep their identity until the state changes', () => {
    const { store, controller, harness, notifications } = setup();
    const bold = store.getState('bold');
    expect(store.getState('bold')).toBe(bold);
    controller.refresh();
    expect(store.getState('bold')).toBe(bold);
    expect(notifications()).toBe(0);

    harness.update({
      selection: { context: { ...PLAIN_CONTEXT, bold: true }, fontFamily: null, fontSize: null },
    });
    controller.refresh();
    const next = store.getState('bold');
    expect(next).not.toBe(bold);
    expect(next.active).toBe(true);
    expect(bold.active).toBe(false);
    expect(notifications()).toBe(1);
  });

  test('evaluate options and arguments separately', () => {
    const { store } = setup();
    expect(store.getState('fontSize', { points: 12 })).toBe(store.getState('fontSize', { points: 12 }));
    expect(store.getState('fontSize', { points: 12 })).not.toBe(store.getState('fontSize'));
    expect(store.getState('fontSize', { points: 0 }).enabled).toBe(false);
  });

  test('publish selection, mode and history changes and stop after unsubscribing', () => {
    const { store, controller, harness, notifications, unsubscribe } = setup({ canUndo: false });
    store.getState('undo');
    store.getState('editingMode');
    store.getState('alignment');
    harness.update({ canUndo: true });
    controller.refresh();
    expect(store.getState('undo').enabled).toBe(true);
    harness.update({ mode: 'suggesting' });
    controller.refresh();
    expect(store.getState('editingMode').value).toBe('suggesting');
    harness.update({
      selection: {
        context: { ...PLAIN_CONTEXT, alignment: 'right' },
        fontFamily: null,
        fontSize: null,
      },
    });
    controller.refresh();
    expect(store.getState('alignment').value).toBe('right');
    expect(notifications()).toBe(3);
    unsubscribe();
    harness.update({ mode: 'viewing' });
    controller.refresh();
    expect(notifications()).toBe(3);
    expect(store.getState('alignment').enabled).toBe(false);
  });

  test('become unavailable when the editor detaches and recover on reattachment', async () => {
    const { store, controller, harness, notifications } = setup();
    const before = store.getState('save');
    expect(before.enabled).toBe(true);
    controller.detach(harness.binding);
    const detached = store.getState('save');
    expect(detached.enabled ? null : detached.disabledReason.code).toBe('editor-unavailable');
    expect(notifications()).toBe(1);
    const failed = await store.execute('save', null);
    expect(failed.ok ? null : failed.failure.code).toBe('editor-unavailable');

    const replacement = testBinding({ status: 'loading' });
    controller.attach(replacement.binding);
    const loading = store.getState('save');
    expect(loading.enabled ? null : loading.disabledReason.code).toBe('document-loading');
    expect(harness.calls).toEqual([]);
  });

  test('keep held snapshots stable past the cache limit and drop stale ones', () => {
    const { store, controller, harness } = setup({ canUndo: true });
    const flood = (from: number) => {
      for (let points = from; points < from + 1200; points += 1) store.getState('fontSize', { points });
    };
    const notify = () => {
      harness.update({ canUndo: !store.getState('undo').enabled });
      controller.refresh();
    };
    const release = controller.hold('fontFamily', { family: 'Georgia' });
    const held = store.getState('fontFamily', { family: 'Georgia' });
    const oneOff = store.getState('fontSize', { points: 1 });
    flood(2);
    notify();
    flood(2000);
    expect(store.getState('fontFamily', { family: 'Georgia' })).toBe(held);
    expect(store.getState('fontSize', { points: 1 })).not.toBe(oneOff);

    release();
    notify();
    flood(4000);
    expect(store.getState('fontFamily', { family: 'Georgia' })).not.toBe(held);
  });

  test('publish chrome changes to chrome subscribers only', () => {
    const { controller, harness, notifications } = setup();
    let changes = 0;
    const unsubscribe = controller.subscribeChrome(() => {
      changes += 1;
    });
    const initial = controller.chrome();
    controller.refresh();
    expect(controller.chrome()).toBe(initial);
    expect(changes).toBe(0);

    harness.state.chrome = { ...harness.state.chrome, isDark: true };
    controller.refresh();
    expect(changes).toBe(1);
    expect(controller.chrome()?.isDark).toBe(true);
    expect(notifications()).toBe(0);
    unsubscribe();
    harness.state.chrome = { ...harness.state.chrome, isDark: false };
    controller.refresh();
    expect(changes).toBe(1);
  });

  test('without an editor every command is stably unavailable', () => {
    const state = UNAVAILABLE_DOCX_COMMANDS.getState('bold');
    expect(UNAVAILABLE_DOCX_COMMANDS.getState('bold')).toBe(state);
    expect(state).toEqual({
      enabled: false,
      disabledReason: { code: 'editor-unavailable', message: 'The editor is not ready.' },
    });
  });
});

describe('command store execution', () => {
  test('orders document commands behind input and runs the rest immediately', async () => {
    const { store, harness } = setup();
    const admitted: string[] = [];
    harness.state.admission = () => {
      admitted.push('admit');
      return Promise.resolve();
    };
    expect(await store.execute('bold', null)).toEqual({ ok: true, status: 'executed' });
    expect(await store.execute('zoom', { scale: 1.5 })).toEqual({ ok: true, status: 'executed' });
    expect(admitted).toEqual(['admit']);
    expect(harness.calls).toEqual([
      { id: 'bold', args: null, ordered: true },
      { id: 'zoom', args: { scale: 1.5 }, ordered: false },
    ]);
  });

  test('enters the queue synchronously, in invocation order', async () => {
    const { store, harness } = setup();
    const order: string[] = [];
    harness.state.admission = () => {
      order.push(`admit:${order.length}`);
      return Promise.resolve();
    };
    const first = store.execute('undo', null);
    const second = store.execute('undo', null);
    expect(order).toEqual(['admit:0', 'admit:1']);
    await Promise.all([first, second]);
    expect(harness.calls.map((call) => call.id)).toEqual(['undo', 'undo']);
  });

  test('maps admission and execution failures to typed results', async () => {
    const { store, harness } = setup();
    for (const code of ['input-failed', 'document-replaced', 'editor-unavailable'] as const) {
      harness.state.admission = () => Promise.reject(new DocxCommandAdmissionError(code));
      const result = await store.execute('bold', null);
      expect(result.ok ? null : result.failure.code).toBe(code);
    }
    harness.state.admission = () => Promise.resolve();
    harness.state.result = () => {
      throw new Error('engine refused');
    };
    const originalError = console.error;
    console.error = () => {};
    try {
      const failed = await store.execute('italic', null);
      expect(failed.ok ? null : failed.failure.code).toBe('command-failed');
    } finally {
      console.error = originalError;
    }
    harness.state.result = () => ({ ok: true, status: 'noop' });
    const noop = await store.execute('bold', null);
    expect(JSON.parse(JSON.stringify(noop))).toEqual({ ok: true, status: 'noop' });
  });

  test('refuses unknown commands without running anything', async () => {
    const { store, harness } = setup();
    const result = await store.execute('explode' as never, null as never);
    expect(result.ok ? null : result.failure.code).toBe('unsupported-command');
    expect(harness.calls).toEqual([]);
  });
});

describe('command descriptors', () => {
  test('are serializable, labelled and cover every command', () => {
    expect(DOCX_COMMAND_IDS.length).toBeGreaterThan(40);
    for (const id of DOCX_COMMAND_IDS) {
      const descriptor = DOCX_COMMAND_DESCRIPTORS[id];
      expect(descriptor.id).toBe(id);
      expect(typeof lookup(descriptor.labelKey)).toBe('string');
      expect(JSON.parse(JSON.stringify(descriptor))).toEqual(descriptor);
    }
  });

  test('bind each chord to exactly one command', () => {
    const chords = DOCX_COMMAND_IDS.flatMap((id) =>
      DOCX_COMMAND_DESCRIPTORS[id].shortcuts.map((shortcut) => shortcut.chord)
    );
    expect(new Set(chords).size).toBe(chords.length);
  });

  test('match and format chords per platform', () => {
    const event = (init: Partial<KeyboardEvent>) =>
      ({ ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, code: '', ...init }) as KeyboardEvent;
    expect(matchesChord('Mod+B', event({ key: 'b', ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+B', event({ key: 'b', metaKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+B', event({ key: 'b', metaKey: true }), false)).toBe(false);
    expect(matchesChord('Mod+B', event({ key: 'B', ctrlKey: true, shiftKey: true }), false)).toBe(false);
    expect(matchesChord('Mod+Shift+Z', event({ key: 'Z', ctrlKey: true, shiftKey: true }), false)).toBe(true);
    expect(
      matchesChord('Mod+Shift+=', event({ key: '+', code: 'Equal', ctrlKey: true, shiftKey: true }), false)
    ).toBe(true);
    expect(matchesChord('Mod+E', event({ key: 'e', ctrlKey: true, altKey: true }), false)).toBe(false);
    expect(formatChord('Mod+Shift+Z', false)).toBe('Ctrl+Shift+Z');
    expect(formatChord('Mod+B', true)).toBe('⌘B');
  });
});
