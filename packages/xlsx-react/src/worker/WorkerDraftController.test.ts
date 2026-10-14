import { describe, expect, mock, test } from 'bun:test';
import type { Direction, Selection } from '@betteroffice/xlsx';
import type { PeerlessWorkerInputCoordinator, PeerlessWorkerPendingInput } from '../commands/peerlessWorkerInputCoordinator';
import { WorkerDraftController, type WorkerNativeDraft } from './WorkerDraftController';

function setup() {
  let resolve!: (value: string) => void;
  const prefill = mock(() => new Promise<string>((done) => { resolve = done; }));
  const edits: WorkerNativeDraft[] = [];
  const state = { navigation: 0, selection: { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } } as Selection };
  const select = mock((_sheet: number, _cell: { row: number; col: number }, _move?: Direction) => {});
  const focus = mock(() => {});
  const controller = new WorkerDraftController({
    generation: () => 1, identity: () => 'book', navigation: () => state.navigation,
    selection: () => state.selection, sheet: () => 0,
    target: () => ({ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }),
    resolveSheet: () => 0, editable: () => true, prefill, select,
    selectionTarget: () => undefined, dependency: () => null, navigate: () => {},
    committed: () => {}, fail: (error) => { throw error; }, report: (error) => { throw error; }, focus,
  });
  controller.attach({
    setDraft: () => {},
    runAfterPendingInput: (operation: () => Promise<string>) => operation(),
    reserveInput: (input: PeerlessWorkerPendingInput) => ({ seq: null, result: input.ready.then((ready) => {
      const draft = input.draft() as WorkerNativeDraft | null;
      if (ready && draft?.modified) edits.push(draft);
      return { seq: edits.length, status: 'applied' as const, committed: true };
    }) }),
    submit: (draft: WorkerNativeDraft) => {
      edits.push(draft);
      return { seq: edits.length, result: Promise.resolve({ seq: edits.length, status: 'applied', committed: true }) };
    },
  } as unknown as PeerlessWorkerInputCoordinator);
  const complete = async (value = 'Q3 Sales Dashboard') => { resolve(value); await Promise.resolve(); await Promise.resolve(); };
  const beforeInput = (inputType: string, data: string | null = null, dataTransfer: DataTransfer | null = null) =>
    controller.holdCellBeforeInput({ inputType, data, dataTransfer, isComposing: false });
  return { controller, prefill, edits, state, select, focus, complete, beforeInput };
}

describe('loading cell input', () => {
  for (const inputType of ['insertText', 'insertReplacementText', 'insertFromPaste', 'insertFromDrop']) {
    test(`holds ${inputType} text until the existing input arrives`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.open();
      expect(beforeInput(inputType, 'X')).toBe(true);
      expect(controller.getSnapshot().cell?.value).toBe('');
      await complete();
      expect(controller.getSnapshot().cell?.value).toBe('Q3 Sales DashboardX');
      expect(controller.getSnapshot().cell?.modified).toBe(true);
    });
  }

  for (const inputType of ['insertFromPaste', 'insertFromDrop']) {
    test(`uses text/plain transfer data for ${inputType}`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.open();
      const getData = mock((_format: string) => 'X');
      expect(beforeInput(inputType, null, { getData } as unknown as DataTransfer)).toBe(true);
      await complete();
      expect(getData).toHaveBeenCalledWith('text/plain');
      expect(controller.getSnapshot().cell?.value).toBe('Q3 Sales DashboardX');
    });
  }

  for (const [inputType, expected] of [
    ['deleteContentBackward', 'Q3 Sales Dashboar'], ['deleteWordBackward', 'Q3 Sales '],
    ['deleteSoftLineBackward', ''], ['deleteHardLineBackward', ''],
    ['deleteContentForward', '3 Sales Dashboard'], ['deleteWordForward', ' Sales Dashboard'],
    ['deleteSoftLineForward', ''], ['deleteHardLineForward', ''],
  ]) {
    test(`replays ${inputType}`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.open();
      if (inputType.endsWith('Forward')) controller.holdCellEdit({ kind: 'move', direction: 'start' });
      expect(beforeInput(inputType)).toBe(true);
      await complete();
      expect(controller.getSnapshot().cell?.value).toBe(expected);
    });
  }

  test('keeps the prefill pending while composing after held input', async () => {
    const { controller, beforeInput, complete, edits } = setup();
    controller.open();
    controller.holdCellEdit({ kind: 'selectAll' });
    beforeInput('insertText', 'probe');
    controller.start('cell');
    expect(controller.getSnapshot().cell?.value).toBe('');
    expect(beforeInput('insertCompositionText', '変')).toBe(false);
    controller.change('cell', '変', null);
    controller.end();
    await Promise.resolve();
    controller.commit('cell');
    await complete();
    expect(edits.map((draft) => draft.value)).toEqual(['probe変']);
  });

  test('closes a held commit so subsequent input belongs to another editor', async () => {
    const { controller, beforeInput, complete, edits } = setup();
    controller.open();
    controller.holdCellEdit({ kind: 'selectAll' });
    beforeInput('insertText', 'probe');
    controller.commit('cell', 'down');
    expect(controller.getSnapshot().cell).toBeNull();
    expect(controller.holdCellKey({ key: 'X', ctrlKey: false, metaKey: false, altKey: false, shiftKey: false })).toBe(false);
    expect(beforeInput('insertText', 'later')).toBe(false);
    await complete();
    expect(edits.map((draft) => draft.value)).toEqual(['probe']);
  });

  test('submits a held commit after navigation without selecting or focusing the old cell', async () => {
    const { controller, beforeInput, complete, edits, state, select, focus } = setup();
    controller.open();
    beforeInput('insertText', 'X');
    controller.commit('cell', 'down');
    expect(select).toHaveBeenCalledTimes(1);
    expect(focus).toHaveBeenCalledTimes(1);
    select.mockClear();
    focus.mockClear();
    state.navigation++;
    state.selection = { anchor: { row: 2, col: 1 }, focus: { row: 2, col: 1 } };
    controller.selectionChanged();
    expect(controller.getSnapshot().cell).toBeNull();
    await complete();
    expect(edits.map((draft) => draft.value)).toEqual(['Q3 Sales DashboardX']);
    expect(select).not.toHaveBeenCalled();
    expect(focus).not.toHaveBeenCalled();
  });

  test('finishes loading input by reserving held keys for submission', async () => {
    const { controller, beforeInput, complete, edits } = setup();
    controller.open();
    beforeInput('insertText', 'X');
    controller.finish();
    await complete();
    expect(edits.map((draft) => draft.value)).toEqual(['Q3 Sales DashboardX']);
    expect(controller.getSnapshot().cell).toBeNull();
  });

  test('keeps a detached commit when composition starts in the newly selected cell', async () => {
    const { controller, beforeInput, complete, edits, state } = setup();
    controller.open();
    beforeInput('insertText', 'X');
    controller.commit('cell');
    state.navigation++;
    state.selection = { anchor: { row: 2, col: 1 }, focus: { row: 2, col: 1 } };
    controller.selectionChanged();
    controller.start('formula');
    await complete();
    expect(edits.map((draft) => draft.value)).toEqual(['Q3 Sales DashboardX']);
    expect(controller.isComposing('formula')).toBe(true);
  });

  test('commits untouched loading input immediately without an edit', async () => {
    const { controller, complete, edits, select } = setup();
    controller.open();
    controller.commit('cell', 'down');
    expect(controller.getSnapshot().cell).toBeNull();
    expect(select).toHaveBeenCalledTimes(1);
    await complete();
    expect(edits).toEqual([]);
    expect(controller.getSnapshot().cell).toBeNull();
  });

  test('seeded opens never hold input or fetch a prefill', () => {
    const { controller, prefill, beforeInput } = setup();
    controller.open('X');
    expect(prefill).not.toHaveBeenCalled();
    expect(beforeInput('insertText', 'Y')).toBe(false);
    expect(controller.getSnapshot().cell?.value).toBe('X');
  });

  test('restores the replayed selection when the cell input mounts later', async () => {
    const { controller, complete } = setup();
    controller.open();
    controller.holdCellEdit({ kind: 'selectAll' });
    await complete();
    const setSelectionRange = mock((_start: number, _end: number, _direction?: string) => {});
    controller.cellInput = { setSelectionRange } as unknown as HTMLInputElement;
    controller.restoreCellSelection();
    expect(setSelectionRange).toHaveBeenCalledWith(0, 'Q3 Sales Dashboard'.length, 'forward');
  });
});
