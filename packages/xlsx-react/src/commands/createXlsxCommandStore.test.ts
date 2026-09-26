import { describe, expect, test } from 'bun:test';
import { en } from '@betteroffice/xlsx-i18n';
import {
  createXlsxCommandController,
  UNAVAILABLE_XLSX_COMMANDS,
  XlsxCommandAdmissionError,
} from './createXlsxCommandStore';
import {
  commandForEvent,
  commandLabelKey,
  formatChord,
  isMacPlatform,
  matchesChord,
  XLSX_COMMAND_DESCRIPTORS,
  XLSX_COMMAND_IDS,
} from './descriptors';
import { PLAIN_FORMATTING, testBinding, testEnvironment } from './testing';

function lookup(key: string): unknown {
  return key
    .split('.')
    .reduce<unknown>(
      (node, part) =>
        node && typeof node === 'object' ? (node as Record<string, unknown>)[part] : undefined,
      en
    );
}

function setup(initial = {}) {
  const harness = testBinding(initial);
  const controller = createXlsxCommandController();
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

const selection = testEnvironment().selection as Exclude<
  ReturnType<typeof testEnvironment>['selection'],
  'chart' | null
>;

describe('xlsx command store snapshots', () => {
  test('keep their identity until the state changes', () => {
    const { store, controller, harness, notifications } = setup();
    const bold = store.getState('bold');
    expect(store.getState('bold')).toBe(bold);
    controller.refresh();
    expect(store.getState('bold')).toBe(bold);
    expect(notifications()).toBe(0);

    harness.update({ selection: { ...selection, formatting: { ...PLAIN_FORMATTING, bold: true } } });
    controller.refresh();
    expect(store.getState('bold')).not.toBe(bold);
    expect(store.getState('bold').active).toBe(true);
    expect(bold.active).toBe(false);
    expect(notifications()).toBe(1);
  });

  test('evaluate options and arguments separately', () => {
    const { store } = setup();
    expect(store.getState('fontSize', { points: 12 })).toBe(
      store.getState('fontSize', { points: 12 })
    );
    expect(store.getState('fontSize', { points: 12 })).not.toBe(store.getState('fontSize'));
    const invalid = store.getState('fontSize', { points: 0 });
    expect(invalid.enabled ? null : invalid.disabledReason.code).toBe('invalid-arguments');
  });

  test('stop notifying after unsubscribing', () => {
    const { store, controller, harness, notifications, unsubscribe } = setup({ canUndo: false });
    store.getState('undo');
    harness.update({ canUndo: true });
    controller.refresh();
    expect(store.getState('undo').enabled).toBe(true);
    expect(notifications()).toBe(1);
    unsubscribe();
    harness.update({ canUndo: false });
    controller.refresh();
    expect(notifications()).toBe(1);
  });

  test('become unavailable when the editor detaches and recover on reattachment', async () => {
    const { store, controller, harness, notifications } = setup();
    expect(store.getState('save').enabled).toBe(true);
    controller.detach(harness.binding);
    const detached = store.getState('save');
    expect(detached.enabled ? null : detached.disabledReason.code).toBe('editor-unavailable');
    expect(notifications()).toBe(1);
    const failed = await store.execute('save', null);
    expect(failed.ok ? null : failed.failure.code).toBe('editor-unavailable');

    controller.attach(testBinding({ status: 'loading' }).binding);
    const loading = store.getState('save');
    expect(loading.enabled ? null : loading.disabledReason.code).toBe('document-loading');
    expect(harness.calls).toEqual([]);
  });

  test('without an editor report editor-unavailable in English', async () => {
    const state = UNAVAILABLE_XLSX_COMMANDS.getState('bold');
    expect(state).toEqual({
      enabled: false,
      disabledReason: { code: 'editor-unavailable', message: 'The editor is not ready.' },
    });
    const result = await UNAVAILABLE_XLSX_COMMANDS.execute('bold', null);
    expect(result.ok ? null : result.failure.code).toBe('editor-unavailable');
  });

  test('are serializable data', () => {
    const { store } = setup();
    for (const id of XLSX_COMMAND_IDS) {
      const state = store.getState(id);
      expect(JSON.parse(JSON.stringify(state))).toEqual(state);
      const descriptor = store.getDescriptor(id);
      expect(JSON.parse(JSON.stringify(descriptor))).toEqual(descriptor);
    }
  });

  test('keep held snapshots past the cache limit', () => {
    const { store, controller } = setup();
    const release = controller.hold('fontFamily', { family: 'Georgia' });
    const held = store.getState('fontFamily', { family: 'Georgia' });
    for (let points = 1; points <= 400; points += 0.5) store.getState('fontSize', { points });
    controller.refresh();
    expect(store.getState('fontFamily', { family: 'Georgia' })).toBe(held);
    release();
  });
});

describe('xlsx command execution', () => {
  test('checks availability again when it runs', async () => {
    const { store, harness } = setup();
    expect(store.getState('merge', { value: 'vertical' }).enabled).toBe(true);
    harness.update({ selection: { ...selection, bottom: 0 } });
    const result = await store.execute('merge', { value: 'vertical' });
    expect(result.ok ? null : result.failure.code).toBe('multiple-rows-required');
    expect(harness.calls).toEqual([]);
  });

  test('orders document commands behind accepted input and runs view commands at once', async () => {
    const { store, harness } = setup();
    let release!: () => void;
    const held = new Promise<void>((resolve) => (release = resolve));
    harness.state.admission = () => held;
    const bold = store.execute('bold', null);
    const zoom = store.execute('zoom', { scale: 2 });
    expect(await zoom).toEqual({ ok: true, status: 'executed' });
    expect(harness.calls.map((call) => call.id)).toEqual(['zoom']);
    release();
    expect(await bold).toEqual({ ok: true, status: 'executed' });
    expect(harness.calls.map((call) => call.id)).toEqual(['zoom', 'bold']);
  });

  test('reports admission failures with their code', async () => {
    const { store, harness } = setup();
    for (const code of ['input-failed', 'document-replaced', 'gesture-active'] as const) {
      harness.state.admission = () => Promise.reject(new XlsxCommandAdmissionError(code));
      const result = await store.execute('undo', null);
      expect(result.ok ? null : result.failure.code).toBe(code);
    }
    expect(harness.calls).toEqual([]);
  });

  test('turns unexpected errors into command-failed', async () => {
    const { store, harness } = setup();
    const originalError = console.error;
    console.error = () => {};
    try {
      harness.state.result = () => {
        throw new Error('engine refused');
      };
      const result = await store.execute('bold', null);
      expect(result.ok ? null : result.failure).toEqual({
        code: 'command-failed',
        message: 'commands.reasons.commandFailed',
      });
    } finally {
      console.error = originalError;
    }
  });

  test('rejects unknown ids and malformed arguments', async () => {
    const { store } = setup();
    const unknown = await store.execute('share' as never, null as never);
    expect(unknown.ok ? null : unknown.failure.code).toBe('unsupported-command');
    const malformed = await store.execute('textColor', { color: 'red' });
    expect(malformed.ok ? null : malformed.failure.code).toBe('invalid-arguments');
  });

  test('commands waiting behind input fail when their target moved on', async () => {
    const { store, harness } = setup();
    let release!: () => void;
    let held = new Promise<void>((resolve) => (release = resolve));
    harness.state.admission = () => held;
    const bold = store.execute('bold', null);
    const save = store.execute('save', null);
    harness.state.target = 'C1:C1';
    release();
    const moved = await bold;
    expect(moved.ok ? null : moved.failure.code).toBe('target-changed');
    expect(await save).toEqual({ ok: true, status: 'executed' });
    expect(harness.calls.map((call) => call.id)).toEqual(['save']);

    held = new Promise<void>((resolve) => (release = resolve));
    const undo = store.execute('undo', null);
    harness.state.generation += 1;
    release();
    const replaced = await undo;
    expect(replaced.ok ? null : replaced.failure.code).toBe('document-replaced');
  });

  test('prepared commands fail once the document or selection moved on', async () => {
    const { controller, harness } = setup();
    const replaced = controller.prepare('fontSize');
    harness.state.generation += 1;
    const first = await replaced.execute({ points: 14 });
    expect(first.ok ? null : first.failure.code).toBe('document-replaced');

    const moved = controller.prepare('textColor');
    harness.state.target = 'C3:C3';
    const second = await moved.execute({ color: '#ff0000' });
    expect(second.ok ? null : second.failure.code).toBe('target-changed');

    const current = controller.prepare('textColor');
    expect(await current.execute({ color: '#ff0000' })).toEqual({ ok: true, status: 'executed' });
    expect(harness.calls.map((call) => call.id)).toEqual(['textColor']);
  });
});

describe('xlsx command descriptors', () => {
  test('name existing locale keys for every command and argument label', () => {
    for (const id of XLSX_COMMAND_IDS) {
      expect(typeof lookup(XLSX_COMMAND_DESCRIPTORS[id].labelKey)).toBe('string');
      const state = setup().store.getState(id);
      for (const option of state.options ?? []) {
        expect(typeof lookup(commandLabelKey(id, option.args as never))).toBe('string');
      }
    }
    for (const args of [
      { value: 'currency' },
      { value: 'percent' },
      { direction: 'increase' },
      { direction: 'decrease' },
    ]) {
      const id = 'value' in args ? 'numberFormat' : 'decimalPlaces';
      expect(typeof lookup(commandLabelKey(id, args as never))).toBe('string');
    }
  });

  test('match chords exactly and format them per platform', () => {
    const event = (init: Partial<KeyboardEvent>) =>
      ({ key: 'z', metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...init }) as KeyboardEvent;
    expect(matchesChord('Mod+Z', event({ ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+Z', event({ metaKey: true }), false)).toBe(false);
    expect(matchesChord('Mod+Z', event({ metaKey: true }), true)).toBe(true);
    expect(matchesChord('Mod+Z', event({ ctrlKey: true, shiftKey: true }), false)).toBe(false);
    expect(formatChord('Mod+Shift+Z', false)).toBe('Ctrl+Shift+Z');
    expect(formatChord('Mod+Shift+Z', true)).toBe('⌘⇧Z');
    const mod = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };
    expect(commandForEvent(event({ key: 'Z', shiftKey: true, ...mod }))?.id).toBe('redo');
    expect(commandForEvent(event({ key: 'c', ...mod }))).toBeNull();
  });
});
