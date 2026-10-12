import { describe, expect, test } from 'bun:test';
import { replayCellInput } from './cellInputReplay';

describe('cell input replay', () => {
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
