import { describe, expect, test } from 'bun:test';
import {
  createPptxCommandController,
  PptxCommandAdmissionError,
  UNAVAILABLE_PPTX_COMMANDS,
} from './createPptxCommandStore';
import { PPTX_COMMAND_IDS } from './descriptors';
import { testBinding } from './testing';
import type { PptxCommandId, PptxCommandState } from './types';

function mounted(overrides = {}) {
  const harness = testBinding(overrides);
  const controller = createPptxCommandController();
  controller.attach(harness.binding);
  return { harness, controller, store: controller.store };
}

describe('PPTX command store', () => {
  test('keeps a snapshot until its state changes and notifies only then', () => {
    const { harness, controller, store } = mounted();
    let notified = 0;
    store.subscribe(() => {
      notified += 1;
    });
    const bold = store.getState('bold');
    controller.refresh();
    expect(store.getState('bold')).toBe(bold);
    expect(notified).toBe(0);

    harness.update({ text: { ...(harness.state.env.text as object), bold: 'mixed' } as never });
    controller.refresh();
    expect(notified).toBe(1);
    expect(store.getState('bold')).not.toBe(bold);
    expect(store.getState('bold').active).toBe('mixed');
  });

  test('every disabled command states a coded, localized reason', () => {
    const { store } = mounted({ status: 'loading' });
    for (const id of PPTX_COMMAND_IDS) {
      const state = store.getState(id) as PptxCommandState;
      expect(state.enabled).toBe(false);
      if (!state.enabled) {
        expect(state.disabledReason.code).toBe('document-loading');
        expect(state.disabledReason.message).toBe('The presentation is still opening.');
      }
    }
  });

  test('descriptors, states and results survive JSON', () => {
    const { store } = mounted();
    for (const id of PPTX_COMMAND_IDS) {
      const descriptor = store.getDescriptor(id);
      expect(JSON.parse(JSON.stringify(descriptor))).toEqual(descriptor);
      const state = store.getState(id);
      expect(JSON.parse(JSON.stringify(state))).toEqual(state);
      for (const option of state.options ?? []) {
        const optionState = store.getState(id as PptxCommandId, option.args as never);
        expect(JSON.parse(JSON.stringify(optionState))).toEqual(optionState);
      }
    }
  });

  test('checks availability again when the command runs', async () => {
    const { harness, store } = mounted();
    let admit!: () => void;
    harness.state.admission = () => new Promise<void>((resolve) => (admit = resolve));
    expect(store.getState('bold').enabled).toBe(true);
    const running = store.execute('bold', null);
    harness.update({ readOnly: true });
    admit();
    expect(await running).toEqual({
      ok: false,
      failure: { code: 'read-only', message: 'The presentation is read-only.' },
    });
    expect(harness.calls).toEqual([]);
  });

  test('fails an admitted command whose document or target changed', async () => {
    const { harness, store } = mounted();
    let admit!: () => void;
    harness.state.admission = () => new Promise<void>((resolve) => (admit = resolve));
    const replaced = store.execute('italic', null);
    harness.state.generation += 1;
    admit();
    expect(await replaced).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });

    const moved = store.execute('italic', null);
    harness.state.targetChanged = true;
    admit();
    expect(await moved).toMatchObject({ ok: false, failure: { code: 'target-changed' } });
    expect(harness.calls).toEqual([]);
  });

  test('reports admission failures and thrown errors with codes', async () => {
    const { harness, store } = mounted();
    harness.state.admission = () => Promise.reject(new PptxCommandAdmissionError('input-failed'));
    expect(await store.execute('undo', null)).toMatchObject({
      ok: false,
      failure: { code: 'input-failed', message: 'Earlier input could not be applied.' },
    });
    harness.state.admission = () => Promise.resolve();
    const error = console.error;
    console.error = () => {};
    try {
      harness.state.result = () => {
        throw new Error('engine refused');
      };
      expect(await store.execute('undo', null)).toMatchObject({
        ok: false,
        failure: { code: 'command-failed' },
      });
    } finally {
      console.error = error;
    }
  });

  test('runs immediate commands without waiting and rejects unknown ids', async () => {
    const { harness, store } = mounted();
    harness.state.admission = () => new Promise<void>(() => {});
    expect(await store.execute('zoom', { scale: 1.5 })).toEqual({ ok: true, status: 'executed' });
    expect(await store.execute('nope' as PptxCommandId, null as never)).toMatchObject({
      ok: false,
      failure: { code: 'unsupported-command' },
    });
    expect(harness.calls).toEqual([{ id: 'zoom', args: { scale: 1.5 } }]);
  });

  test('prepared commands complete against what they opened for', async () => {
    const { harness, controller } = mounted();
    const pending = controller.prepare('textColor');
    harness.state.generation += 1;
    expect(await pending.execute({ color: '#ff0000' })).toMatchObject({
      ok: false,
      failure: { code: 'document-replaced' },
    });
    const fresh = controller.prepare('textColor');
    expect(await fresh.execute({ color: '#ff0000' })).toEqual({ ok: true, status: 'executed' });
  });

  test('without an editor every command is unavailable, and detaching restores that', () => {
    expect(UNAVAILABLE_PPTX_COMMANDS.getState('save')).toEqual({
      enabled: false,
      disabledReason: { code: 'editor-unavailable', message: 'The editor is not ready.' },
    });
    const { harness, controller, store } = mounted();
    expect(store.getState('save').enabled).toBe(true);
    controller.detach(harness.binding);
    expect(store.getState('save')).toMatchObject({
      enabled: false,
      disabledReason: { code: 'editor-unavailable' },
    });
  });

  test('keeps held snapshots stable past the cache limit', () => {
    const { controller, store } = mounted();
    const release = controller.hold('bold');
    const bold = store.getState('bold');
    for (let points = 1; points <= 400; points += 0.5) store.getState('fontSize', { points });
    controller.refresh();
    expect(store.getState('bold')).toBe(bold);
    release();
  });
});
