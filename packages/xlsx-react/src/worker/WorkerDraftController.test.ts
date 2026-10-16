import { afterAll, afterEach, describe, expect, mock, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createElement } from 'react';
import type { Direction, Selection } from '@betteroffice/xlsx';
import type { PeerlessWorkerInputCoordinator, PeerlessWorkerPendingInput } from '../commands/peerlessWorkerInputCoordinator';
import { WorkerDraftController, type WorkerNativeDraft } from './WorkerDraftController';

import { PeerlessWorkerCellInput } from './PeerlessWorkerInputs';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function setup() {
  let resolve!: (value: string) => void;
  const prefill = mock(() => new Promise<string>((done) => { resolve = done; }));
  const edits: WorkerNativeDraft[] = [];
  const state = { navigation: 0, identity: 'book', sheet: 0, throughSeq: 0, pending: false, valueRevision: 'v1',
    selection: { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } } as Selection };
  const select = mock((_sheet: number, _cell: { row: number; col: number }, _move?: Direction) => {});
  const focus = mock(() => {});
  const controller = new WorkerDraftController({
    generation: () => 1, identity: () => state.identity, valueRevision: () => state.valueRevision, navigation: () => state.navigation,
    selection: () => state.selection, sheet: () => state.sheet,
    target: () => ({ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }),
    resolveSheet: () => 0, editable: () => true, prefill, select,
    selectionTarget: () => undefined, dependency: () => null, navigate: () => {},
    committed: () => {}, fail: (error) => { throw error; }, report: (error) => { throw error; }, focus,
  });
  controller.attach({
    get submittedThroughSeq() { return state.throughSeq; },
    get pending() { return state.pending; },
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
  for (const inputType of ['insertText', 'insertFromPaste', 'insertFromDrop']) {
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

  test('replacement text without a known full range appends to the pending prefill', async () => {
    const { controller, beforeInput, complete } = setup();
    controller.open();
    expect(beforeInput('insertReplacementText', 'X')).toBe(true);
    await complete();
    expect(controller.getSnapshot().cell?.value).toBe('Q3 Sales DashboardX');
  });

  for (const inputType of ['insertText', 'insertFromPaste']) {
    test(`a native select event replaces the pending prefill with ${inputType}`, async () => {
      const { controller, complete } = setup();
      controller.focusedCellInput = {
        identity: controller.focusedCellIdentity,
        value: 'initial',
      };
      controller.open();
      const view = render(
        createElement(PeerlessWorkerCellInput, {
          controller,
          rect: () => ({ x: 0, y: 0, w: 100, h: 24 }),
          disabled: false,
          zoom: 1,
        }),
      );
      const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
      expect(input.value).toBe('initial');
      input.setSelectionRange(0, input.value.length);
      fireEvent(input, new Event('select', { bubbles: true }));
      const event = new InputEvent('beforeinput', {
        inputType,
        data: 'X',
        cancelable: true,
        bubbles: true,
      });
      fireEvent(input, event);
      expect(event.defaultPrevented).toBe(true);
      await act(() => complete());
      expect(input.value).toBe('X');
    });
  }

  test('mounts and focuses the native input while an unknown prefill is pending', async () => {
    const { controller, complete } = setup();
    controller.open();
    const view = render(
      createElement(PeerlessWorkerCellInput, {
        controller,
        rect: () => ({ x: 0, y: 0, w: 100, h: 24 }),
        disabled: false,
        zoom: 1,
      }),
    );
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('');
    await act(() => complete('initial'));
    expect(input.value).toBe('initial');
    expect(input.selectionStart).toBe('initial'.length);
    expect(input.selectionEnd).toBe('initial'.length);
    input.setSelectionRange(0, input.value.length);
    fireEvent.change(input, { target: { value: 'X' } });
    expect(controller.getSnapshot().cell?.value).toBe('X');
  });

  for (const [inputType, text] of [
    ['insertFromPaste', 'paste'],
    ['insertText', '😀'],
  ]) {
    test(`holds native ${inputType} on the mounted uncached input`, async () => {
      const { controller, complete } = setup();
      controller.open();
      const view = render(
        createElement(PeerlessWorkerCellInput, {
          controller,
          rect: () => ({ x: 0, y: 0, w: 100, h: 24 }),
          disabled: false,
          zoom: 1,
        }),
      );
      const input = view.getByTestId('xlsx-cell-editor');
      expect(document.activeElement).toBe(input);
      const event = new InputEvent('beforeinput', {
        inputType,
        data: text,
        cancelable: true,
        bubbles: true,
      });
      fireEvent(input, event);
      expect(event.defaultPrevented).toBe(true);
      await act(() => complete('initial'));
      expect((input as HTMLInputElement).value).toBe(`initial${text}`);
    });
  }

  for (const changed of ['identity', 'navigation', 'sheet', 'focus', 'value'] as const) {
    test(`discards the focused input preview after ${changed} changes`, () => {
      const { controller, state } = setup();
      controller.focusedCellInput = {
        identity: controller.focusedCellIdentity,
        value: 'stale',
      };
      if (changed === 'identity') state.identity = 'replacement';
      else if (changed === 'navigation') state.navigation++;
      else if (changed === 'sheet') state.sheet++;
      else if (changed === 'focus') state.selection.focus = { row: 1, col: 0 };
      else state.valueRevision = 'v2';
      controller.open();
      expect(controller.getSnapshot().cell?.value).toBe('');
    });
  }

  test('reuses the focused input through open, cancel and nonmutating command sequences', () => {
    const { controller, state } = setup();
    controller.focusedCellInput = {
      identity: controller.focusedCellIdentity,
      value: 'initial',
    };
    controller.open();
    controller.cancel();
    state.throughSeq++;
    controller.open();
    expect(controller.getSnapshot().cell?.value).toBe('initial');
  });

  for (const value of ['', 'initial']) {
    test(`keeps ${value === '' ? 'an empty' : 'a nonempty'} preview safe behind a pending fence`, () => {
      const { controller, state } = setup();
      controller.focusedCellInput = {
        identity: controller.focusedCellIdentity,
        value,
      };
      state.pending = true;
      controller.open();
      expect(controller.getSnapshot().cell?.value).toBe('');
    });
  }

  for (const arrival of ['prefill', 'beforeinput']) {
    test(`keeps a full DOM selection when ${arrival} arrives before its select event`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.focusedCellInput = {
        identity: controller.focusedCellIdentity,
        value: 'initial',
      };
      controller.open();
      const input = document.createElement('input');
      input.value = controller.getSnapshot().cell!.value;
      controller.cellInput = input;
      input.setSelectionRange(0, input.value.length);
      if (arrival === 'beforeinput') beforeInput('insertText', 'X');
      await complete('initial');
      if (arrival === 'prefill') {
        input.setRangeText('X', input.selectionStart!, input.selectionEnd!, 'end');
        controller.change('cell', input.value, input);
      }
      expect(controller.getSnapshot().cell?.value).toBe('X');
    });
  }

  for (const key of ['ArrowRight', 'ArrowLeft', 'End', 'Home', 'Delete', 'Backspace']) {
    test(`captures native select-all before the held ${key} operation`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.focusedCellInput = {
        identity: controller.focusedCellIdentity,
        value: 'initial',
      };
      controller.open();
      const input = document.createElement('input');
      input.value = 'initial';
      controller.cellInput = input;
      input.setSelectionRange(0, 7);
      expect(
        controller.holdCellKey({
          key,
          ctrlKey: false,
          metaKey: false,
          altKey: false,
          shiftKey: false,
        }),
      ).toBe(true);
      controller.captureCellSelection();
      if (key !== 'Delete' && key !== 'Backspace') beforeInput('insertText', 'X');
      await complete('initial');
      const expected =
        key === 'Delete' || key === 'Backspace' ? '' : key === 'Home' || key === 'ArrowLeft' ? 'Xinitial' : 'initialX';
      expect(controller.getSnapshot().cell?.value).toBe(expected);
    });
  }

  for (const caret of [0, 3, 7]) {
    test(`respects a native selection collapsed to ${caret} after select-all`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.focusedCellInput = {
        identity: controller.focusedCellIdentity,
        value: 'initial',
      };
      controller.open();
      const input = document.createElement('input');
      input.value = 'initial';
      controller.cellInput = input;
      input.setSelectionRange(0, 7);
      controller.captureCellSelection();
      input.setSelectionRange(caret, caret);
      beforeInput('insertText', 'X');
      await complete('initial');
      expect(controller.getSnapshot().cell?.value).toBe('initial'.slice(0, caret) + 'X' + 'initial'.slice(caret));
    });
  }

  for (const route of ['selection', 'target ranges']) {
    for (const [value, start, end, text, expected] of [
      ['teh report', 0, 3, 'the', 'the report'],
      ['Q3 Sales e', 9, 10, 'é', 'Q3 Sales é'],
      ['initial', 0, 7, 'X', 'X'],
    ] as const) {
      test(`replaces only ${start}..${end} of ${value} using ${route}`, async () => {
        const { controller, complete } = setup();
        controller.focusedCellInput = {
          identity: controller.focusedCellIdentity,
          value,
        };
        controller.open();
        const input = document.createElement('input');
        input.value = value;
        controller.cellInput = input;
        input.setSelectionRange(route === 'selection' ? start : value.length, route === 'selection' ? end : value.length);
        controller.holdCellBeforeInput({
          inputType: 'insertReplacementText',
          data: text,
          dataTransfer: null,
          isComposing: false,
          ...(route === 'target ranges'
            ? {
                getTargetRanges: () => [{ startOffset: start, endOffset: end }] as StaticRange[],
              }
            : {}),
        });
        await complete(value);
        expect(controller.getSnapshot().cell?.value).toBe(expected);
      });
    }
  }

  test('preserves replayed select-all until an unfocused input gains focus', async () => {
    const { controller, complete } = setup();
    controller.open();
    const input = document.createElement('input');
    controller.cellInput = input;
    controller.holdCellEdit({ kind: 'selectAll' });
    await complete('initial');
    input.setSelectionRange(7, 7);
    controller.restoreCellSelection();
    expect(input.selectionStart).toBe(0);
    expect(input.selectionEnd).toBe(7);
  });

  test('typing after F2 without selecting appends at the prefill caret', async () => {
    const { controller, beforeInput, complete } = setup();
    controller.open();
    expect(
      controller.holdCellKey({
        key: 'X',
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        shiftKey: false,
      }),
    ).toBe(false);
    beforeInput('insertText', 'X');
    await complete('initial');
    expect(controller.getSnapshot().cell?.value).toBe('initialX');
  });

  for (const modifier of ['ctrlKey', 'metaKey']) {
    test(`${modifier}+A followed by typing replaces the pending prefill`, async () => {
      const { controller, beforeInput, complete } = setup();
      controller.open();
      expect(
        controller.holdCellKey({
          key: 'a',
          ctrlKey: false,
          metaKey: false,
          altKey: false,
          shiftKey: false,
          [modifier]: true,
        }),
      ).toBe(true);
      beforeInput('insertText', 'X');
      await complete('initial');
      expect(controller.getSnapshot().cell?.value).toBe('X');
    });
  }

  for (const route of ['keyboard', 'DOM']) {
    test(`${route} selection after prefill survives the layout commit`, async () => {
      const { controller, complete } = setup();
      controller.open();
      const input = document.createElement('input');
      document.body.append(input);
      input.focus();
      controller.cellInput = input;
      await complete('initial');
      expect(controller.loadingCell).toBe(false);
      expect(input.selectionStart).toBe('initial'.length);
      if (route === 'keyboard') {
        expect(
          controller.holdCellKey({
            key: 'a',
            ctrlKey: true,
            metaKey: false,
            altKey: false,
            shiftKey: false,
          }),
        ).toBe(false);
      }
      input.setSelectionRange(0, input.value.length);
      controller.restoreCellSelection();
      input.setRangeText('X', input.selectionStart!, input.selectionEnd!, 'end');
      controller.change('cell', input.value, input);
      expect(controller.getSnapshot().cell?.value).toBe('X');
      input.remove();
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
