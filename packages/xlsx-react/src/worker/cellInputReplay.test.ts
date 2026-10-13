import { describe, expect, test } from 'bun:test';
import { cellInputOperation, replayCellInput } from './cellInputReplay';

describe('cell input replay', () => {
  test('maps platform word and line keys through the same operation model', () => {
    const key = (key: string, modifiers: Partial<{ ctrlKey: boolean; metaKey: boolean; altKey: boolean; shiftKey: boolean }> = {}) =>
      ({ key, ctrlKey: false, metaKey: false, altKey: false, shiftKey: false, ...modifiers });
    expect(cellInputOperation(key('ArrowLeft', { ctrlKey: true, shiftKey: true }), false))
      .toEqual({ kind: 'move', direction: 'left', unit: 'word', extend: true });
    expect(cellInputOperation(key('ArrowRight', { altKey: true }), true))
      .toEqual({ kind: 'move', direction: 'right', unit: 'word', extend: false });
    expect(cellInputOperation(key('ArrowLeft', { metaKey: true }), true))
      .toEqual({ kind: 'move', direction: 'start', extend: false });
    expect(cellInputOperation(key('Backspace', { ctrlKey: true }), false))
      .toEqual({ kind: 'delete', direction: 'backward', unit: 'word' });
    expect(cellInputOperation(key('Delete', { altKey: true }), true))
      .toEqual({ kind: 'delete', direction: 'forward', unit: 'word' });
    expect(cellInputOperation(key('Backspace', { metaKey: true }), true))
      .toEqual({ kind: 'delete', direction: 'backward', unit: 'line' });
  });

  test('keeps a caret move when another edit arrives before the input mounts', () => {
    const state = replayCellInput('value', [{ kind: 'move', direction: 'start' }]);
    expect(replayCellInput(state.value, [{ kind: 'insert', text: 'X' }], state))
      .toEqual({ value: 'Xvalue', anchor: 1, focus: 1 });
  });
  test('select-all followed by insertion replaces the value', () => {
    expect(replayCellInput('Q3 Sales Dashboard', [{ kind: 'selectAll' }, { kind: 'insert', text: 'probe' }]))
      .toEqual({ value: 'probe', anchor: 5, focus: 5 });
  });

  test('inserts append at the initial caret and preserve their order', () => {
    expect(replayCellInput('value', [{ kind: 'insert', text: 'X' }, { kind: 'insert', text: 'Y' }]))
      .toEqual({ value: 'valueXY', anchor: 7, focus: 7 });
  });

  test('Home followed by insertion prepends', () => {
    expect(replayCellInput('value', [{ kind: 'move', direction: 'start' }, { kind: 'insert', text: 'X' }]))
      .toEqual({ value: 'Xvalue', anchor: 1, focus: 1 });
  });

  test('Shift+Home extends from the end and insertion replaces that selection', () => {
    expect(replayCellInput('value', [{ kind: 'move', direction: 'start', extend: true }]))
      .toEqual({ value: 'value', anchor: 5, focus: 0 });
    expect(replayCellInput('value', [{ kind: 'move', direction: 'start', extend: true }, { kind: 'insert', text: 'X' }]))
      .toEqual({ value: 'X', anchor: 1, focus: 1 });
  });

  test('backward deletion removes one character without a selection', () => {
    expect(replayCellInput('value', [{ kind: 'delete', direction: 'backward', unit: 'character' }]))
      .toEqual({ value: 'valu', anchor: 4, focus: 4 });
  });

  test('backward deletion removes the selection', () => {
    expect(replayCellInput('value', [{ kind: 'move', direction: 'left', extend: true },
      { kind: 'move', direction: 'left', extend: true }, { kind: 'delete', direction: 'backward', unit: 'character' }]))
      .toEqual({ value: 'val', anchor: 3, focus: 3 });
  });

  test('word deletion skips trailing spaces and removes the preceding word', () => {
    expect(replayCellInput('Q3 Sales Dashboard  ', [{ kind: 'delete', direction: 'backward', unit: 'word' }]))
      .toEqual({ value: 'Q3 Sales ', anchor: 9, focus: 9 });
  });

  test('word movement with Shift selects only the preceding word', () => {
    expect(replayCellInput('Q3 Sales Dashboard', [{ kind: 'move', direction: 'left', unit: 'word', extend: true },
      { kind: 'insert', text: 'X' }])).toEqual({ value: 'Q3 Sales X', anchor: 10, focus: 10 });
  });

  test('word movement skips whitespace and supports forward selection', () => {
    expect(replayCellInput('Q3 Sales  ', [{ kind: 'move', direction: 'left', unit: 'word' }]))
      .toEqual({ value: 'Q3 Sales  ', anchor: 3, focus: 3 });
    expect(replayCellInput('Q3 Sales', [{ kind: 'move', direction: 'start' },
      { kind: 'move', direction: 'right', unit: 'word', extend: true }, { kind: 'insert', text: 'X' }]))
      .toEqual({ value: 'X Sales', anchor: 1, focus: 1 });
  });

  test('forward word deletion starts at the current caret', () => {
    expect(replayCellInput('Q3 Sales', [{ kind: 'move', direction: 'start' },
      { kind: 'delete', direction: 'forward', unit: 'word' }]))
      .toEqual({ value: ' Sales', anchor: 0, focus: 0 });
  });

  test('line deletion removes the text on the specified side', () => {
    expect(replayCellInput('value', [{ kind: 'move', direction: 'left' },
      { kind: 'delete', direction: 'backward', unit: 'line' }]))
      .toEqual({ value: 'e', anchor: 0, focus: 0 });
    expect(replayCellInput('value', [{ kind: 'move', direction: 'start' },
      { kind: 'move', direction: 'right' }, { kind: 'delete', direction: 'forward', unit: 'line' }]))
      .toEqual({ value: 'v', anchor: 1, focus: 1 });
  });

  test('arrow moves collapse a selection and stop at the value bounds', () => {
    expect(replayCellInput('value', [{ kind: 'selectAll' }, { kind: 'move', direction: 'left' },
      { kind: 'move', direction: 'left' }])).toEqual({ value: 'value', anchor: 0, focus: 0 });
    expect(replayCellInput('value', [{ kind: 'selectAll' }, { kind: 'move', direction: 'right' },
      { kind: 'move', direction: 'right' }])).toEqual({ value: 'value', anchor: 5, focus: 5 });
  });

  test('character edits preserve complete Unicode code points', () => {
    expect(replayCellInput('A😀', [{ kind: 'delete', direction: 'backward', unit: 'character' }]))
      .toEqual({ value: 'A', anchor: 1, focus: 1 });
  });
});
