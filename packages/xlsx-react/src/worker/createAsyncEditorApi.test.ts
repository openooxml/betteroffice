import { describe, expect, it, mock } from 'bun:test';
import { createXlsxCommandController, XlsxCommandAdmissionError, xlsxCommandController } from '../commands/createXlsxCommandStore';
import { testBinding } from '../commands/testing';
import { createAsyncEditorApi, XlsxAsyncApiSyncAccessError } from './createAsyncEditorApi';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

function apiSession() {
  return { current: true, retiring: false, ready: true, failure: null, session: {},
    whenReady: async () => {} } as unknown as PeerlessWorkbookSession;
}

const bridge = { focus: mock(() => {}) } as unknown as WorkerEditorApiBridge;

describe('async editor commands', () => {
  for (const [member, alternative] of [
    ['saveAsync', 'save()'], ['handle', 'commands'],
    ['cell', 'readCells(request)'], ['rangeCells', 'readCells(request)'],
    ['readCellsSync', 'readCells(request)'], ['editCell', 'applyEdits(request)'],
    ['selectCells', 'plugin.navigation.selectCells'],
    ['undo', 'commands.execute("undo", null)'], ['redo', 'commands.execute("redo", null)'],
  ]) {
    it(`directs legacy ${member} access to ${alternative}`, () => {
      const api = createAsyncEditorApi(apiSession(), createXlsxCommandController().store, () => bridge);
      let refusal: unknown;
      try {
        const value = Reflect.get(api, member);
        if (typeof value === 'function') value();
      } catch (error) { refusal = error; }
      expect(refusal).toBeInstanceOf(XlsxAsyncApiSyncAccessError);
      expect(refusal).toMatchObject({ name: 'XlsxAsyncApiSyncAccessError', code: 'sync-api-unavailable', member, alternative });
      expect((refusal as Error).message).toContain(alternative);
      expect((refusal as Error).message).toContain('legacySyncApi');
      expect((refusal as Error).message).not.toContain('with asyncApi');
    });
  }

  it('delegates state and subscriptions while keeping execution on its workbook', async () => {
    const controller = createXlsxCommandController();
    const firstBinding = testBinding();
    controller.attach(firstBinding.binding);
    const firstSession = apiSession();
    const first = createAsyncEditorApi(firstSession, controller.store, () => bridge);
    expect(first.commands).not.toBe(controller.store);
    expect(first.commands.getDescriptor('undo')).toBe(controller.store.getDescriptor('undo'));
    expect(first.commands.getState('undo')).toBe(controller.store.getState('undo'));
    const changed = mock(() => {});
    const unsubscribe = first.commands.subscribe(changed);
    firstBinding.update({ canUndo: false });
    controller.refresh();
    expect(changed).toHaveBeenCalledTimes(1);
    unsubscribe();
    firstBinding.update({ canUndo: true });
    controller.refresh();
    expect(changed).toHaveBeenCalledTimes(1);
    Object.defineProperty(firstSession, 'current', { value: false });
    const nextBinding = testBinding();
    controller.attach(nextBinding.binding);
    const next = createAsyncEditorApi(apiSession(), controller.store, () => bridge);
    await expect(first.commands.execute('undo', null)).rejects.toBeInstanceOf(XlsxCommandAdmissionError);
    await expect(first.commands.execute('undo', null)).rejects.toMatchObject({ code: 'document-replaced' });
    expect(nextBinding.calls).toEqual([]);
    await expect(next.commands.execute('undo', null)).resolves.toMatchObject({ ok: true });
    expect(nextBinding.calls).toHaveLength(1);
  });

  for (const stale of ['replaced', 'retiring'] as const) {
    it(`checks its workbook again after command admission (${stale})`, async () => {
      const controller = createXlsxCommandController();
      const binding = testBinding();
      let release!: () => void;
      const admission = new Promise<void>((resolve) => { release = resolve; });
      binding.state.admission = () => admission;
      controller.attach(binding.binding);
      const session = apiSession();
      const api = createAsyncEditorApi(session, controller.store, () => bridge);
      const command = api.commands.execute('undo', null);
      if (stale === 'retiring') session.retiring = true;
      else Object.defineProperty(session, 'current', { value: false });
      release();
      await expect(command).rejects.toMatchObject({ code: 'document-replaced' });
      expect(binding.calls).toEqual([]);
    });

    it(`checks its workbook again after prepared command admission (${stale})`, async () => {
      const controller = createXlsxCommandController();
      const binding = testBinding();
      let release!: () => void;
      const admission = new Promise<void>((resolve) => { release = resolve; });
      binding.state.admission = () => admission;
      controller.attach(binding.binding);
      const session = apiSession();
      const api = createAsyncEditorApi(session, controller.store, () => bridge);
      const prepared = xlsxCommandController(api.commands)!.prepare('fontSize');
      const command = prepared.execute({ points: 14 });
      if (stale === 'retiring') session.retiring = true;
      else Object.defineProperty(session, 'current', { value: false });
      release();
      await expect(command).rejects.toMatchObject({ code: 'document-replaced' });
      expect(binding.calls).toEqual([]);
    });
  }

  it('keeps prepared commands on the workbook that opened them', async () => {
    const controller = createXlsxCommandController();
    const binding = testBinding();
    controller.attach(binding.binding);
    const session = apiSession();
    const api = createAsyncEditorApi(session, controller.store, () => bridge);
    const prepared = xlsxCommandController(api.commands)!.prepare('textColor');
    Object.defineProperty(session, 'current', { value: false });
    const next = testBinding();
    controller.attach(next.binding);
    await expect(prepared.execute({ color: '#ff0000' })).rejects.toMatchObject({ code: 'document-replaced' });
    expect(binding.calls).toEqual([]);
    expect(next.calls).toEqual([]);
  });

  it('refuses stale preparation before capturing the replacement workbook', () => {
    const controller = createXlsxCommandController();
    controller.attach(testBinding().binding);
    const session = apiSession();
    const api = createAsyncEditorApi(session, controller.store, () => bridge);
    Object.defineProperty(session, 'current', { value: false });
    const next = testBinding();
    const capture = mock(next.binding.capture);
    next.binding.capture = capture;
    controller.attach(next.binding);
    expect(() => xlsxCommandController(api.commands)!.prepare('fontSize')).toThrow('The workbook was replaced');
    expect(capture).not.toHaveBeenCalled();
    expect(next.calls).toEqual([]);
  });

  it('captures the selection when a guarded choice opens', async () => {
    const controller = createXlsxCommandController();
    const binding = testBinding();
    controller.attach(binding.binding);
    const api = createAsyncEditorApi(apiSession(), controller.store, () => bridge);
    const choices = xlsxCommandController(api.commands)!;
    const moved = choices.prepare('textColor');
    binding.state.target = 'C3:C3';
    await expect(moved.execute({ color: '#ff0000' })).resolves.toMatchObject({
      ok: false, failure: { code: 'target-changed' },
    });
    expect(binding.calls).toEqual([]);
    await expect(choices.prepare('fontSize').execute({ points: 14 })).resolves.toMatchObject({ ok: true });
    expect(binding.calls).toEqual([{ id: 'fontSize', args: { points: 14 }, ordered: true }]);
  });
});
