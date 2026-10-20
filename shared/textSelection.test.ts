import { expect, test } from 'bun:test';
import { caretGoalX, verticalCaretMove } from './textSelection';

test('text_selection_import_without_segmenter', async () => {
  const original = Object.getOwnPropertyDescriptor(Intl, 'Segmenter');
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    Object.defineProperty(Intl, 'Segmenter', { configurable: true, value: undefined });
    const path = './textSelection.ts?lazy-segmenters';
    const helpers = await Promise.race([import(path), new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error('text selection import timed out')), 5000);
    })]);
    expect(helpers.paragraphRangeAt('one\ntwo', 1)).toEqual({ start: 0, end: 3 });
    if (original) Object.defineProperty(Intl, 'Segmenter', original);
    expect(helpers.wordRangeAt('Hello, café world', 8)).toEqual({ start: 7, end: 11 });
    expect(helpers.wordBoundary('Hello brave world', 8, -1)).toBe(6);
    expect([...helpers.graphemeBoundaries('e\u0301😀a')]).toEqual([0, 2, 4, 5]);
  } finally {
    clearTimeout(timer);
    if (original) Object.defineProperty(Intl, 'Segmenter', original);
    else Reflect.deleteProperty(Intl, 'Segmenter');
  }
});

test('vertical caret movement preserves the column-relative goal across columns', () => {
  for (const offset of [100, -100]) {
    const row = (start: number, x: number, y: number, width = 40, columnOffset = 0) => ({
      x, y, columnOffset, start, end: start + width / 20,
      caretStops: Array.from({ length: width / 20 + 1 }, (_, index) => ({
        position: start + index, x: x + index * 20,
      })),
    });
    const lines = [row(0, 10, 0), row(3, 10, 20, 20),
      row(5, 10 + offset, 0, 40, offset), row(8, 10 + offset, 20, 40, offset)];
    const start = { position: 2, lineIndex: 0 };
    const goal = caretGoalX(lines, start);
    const short = verticalCaretMove(lines, start, 'down', goal);
    expect(short).toEqual({ position: 4, lineIndex: 1 });
    const next = verticalCaretMove(lines, short, 'down', goal);
    expect(next).toEqual({ position: 7, lineIndex: 2 });
    expect(verticalCaretMove(lines, next, 'down', goal)).toEqual({ position: 10, lineIndex: 3 });
    const reverseGoal = caretGoalX(lines, next);
    expect(verticalCaretMove(lines, next, 'up', reverseGoal)).toEqual(short);
    expect(verticalCaretMove(lines, short, 'up', reverseGoal)).toEqual(start);
  }
});

test('vertical caret movement uses the true column origin for centered short lines', () => {
  for (const offset of [100, -100]) {
    for (const inset of [40, 80]) {
      const row = (start: number, x: number, y: number, width: number, columnOffset: number) => ({
        x, y, columnOffset, start, end: start + width / 10,
        caretStops: Array.from({ length: width / 10 + 1 }, (_, index) => ({
          position: start + index, x: x + index * 10,
        })),
      });
      const lines = [row(0, 0, 0, 100, 0), row(11, inset, 20, 20, 0),
        row(14, offset + inset, 0, 20, offset)];
      const start = { position: 12, lineIndex: 1 };
      const next = { position: 15, lineIndex: 2 };
      const goal = caretGoalX(lines, start);
      expect(verticalCaretMove(lines, start, 'down', goal)).toEqual(next);
      expect(caretGoalX(lines, next)).toBe(goal);
      expect(verticalCaretMove(lines, next, 'up')).toEqual(start);
    }
  }
});
