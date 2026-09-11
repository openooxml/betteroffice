import { describe, expect, it } from 'bun:test';
import {
  caretGoalX,
  extendTextRange,
  lineEdge,
  paragraphRangeAt,
  sameCaretGoalKey,
  textRangeAt,
  verticalCaretMove,
  wordBoundary,
  wordRangeAt,
} from './textSelection';
import type { CaretLine } from './textSelection';

describe('pptx text selection boundaries', () => {
  it('uses Unicode word boundaries', () => {
    expect(wordRangeAt('Hello, café world', 8)).toEqual({ start: 7, end: 11 });
    expect(textRangeAt('Hello, café world', 14, 'word')).toEqual({
      start: 12,
      end: 17,
    });
  });

  it('selects the whitespace segment under the pointer', () => {
    expect(wordRangeAt('Hello,   world', 7)).toEqual({ start: 6, end: 9 });
  });

  it('selects the paragraph containing the caret position', () => {
    const text = 'First line\nSecond paragraph\nThird';
    expect(paragraphRangeAt(text, 3)).toEqual({ start: 0, end: 10 });
    expect(paragraphRangeAt(text, 10)).toEqual({ start: 0, end: 10 });
    expect(textRangeAt(text, 18, 'paragraph')).toEqual({ start: 11, end: 27 });
    expect(paragraphRangeAt(text, text.length)).toEqual({
      start: 28,
      end: 33,
    });
  });

  it('extends a unit selection without splitting either boundary', () => {
    expect(
      extendTextRange({ start: 6, end: 10 }, { start: 0, end: 5 })
    ).toEqual({ anchor: 10, focus: 0 });
    expect(
      extendTextRange({ start: 6, end: 10 }, { start: 11, end: 15 })
    ).toEqual({ anchor: 6, focus: 15 });
  });
});

describe('pptx caret movement', () => {
  const lines: CaretLine[] = [
    { start: 0, end: 5, caretStops: stops(0, [0, 10, 20, 30, 40, 50]) },
    { start: 5, end: 8, caretStops: stops(5, [0, 12, 24, 36]) },
    { start: 8, end: 13, caretStops: stops(8, [0, 10, 20, 30, 40, 50]) },
  ];

  it('moves between lines at the nearest column', () => {
    expect(verticalCaretMove(lines, { position: 3, lineIndex: 0 }, 'down')).toEqual({
      position: 7,
      lineIndex: 1,
    });
    expect(verticalCaretMove(lines, { position: 7, lineIndex: 1 }, 'up')).toEqual({
      position: 2,
      lineIndex: 0,
    });
  });

  it('stays put at the first and last line', () => {
    expect(verticalCaretMove(lines, { position: 2, lineIndex: 0 }, 'up')).toEqual({
      position: 2,
      lineIndex: 0,
    });
    expect(verticalCaretMove(lines, { position: 11, lineIndex: 2 }, 'down')).toEqual({
      position: 11,
      lineIndex: 2,
    });
  });

  it('holds the goal column across a short line', () => {
    const goal = caretGoalX(lines, { position: 5, lineIndex: 0 });
    expect(goal).toBe(50);
    const middle = verticalCaretMove(lines, { position: 5, lineIndex: 0 }, 'down', goal);
    expect(middle).toEqual({ position: 8, lineIndex: 1 });
    expect(verticalCaretMove(lines, middle, 'down', goal)).toEqual({
      position: 13,
      lineIndex: 2,
    });
  });

  it('keeps shared endpoints on their visual line', () => {
    expect(lineEdge(lines, { position: 5, lineIndex: 0 }, 'start')).toEqual({
      position: 0,
      lineIndex: 0,
    });
    expect(lineEdge(lines, { position: 5, lineIndex: 1 }, 'start')).toEqual({
      position: 5,
      lineIndex: 1,
    });
    expect(lineEdge(lines, { position: 4, lineIndex: 0 }, 'end')).toEqual({
      position: 5,
      lineIndex: 0,
    });
  });

  it('moves through an empty shared-endpoint line', () => {
    const empty: CaretLine[] = [
      { start: 0, end: 1, caretStops: stops(0, [0, 10]) },
      { start: 1, end: 1, caretStops: [{ position: 1, x: 0 }] },
      { start: 1, end: 2, caretStops: stops(1, [0, 10]) },
    ];
    const goal = caretGoalX(empty, { position: 1, lineIndex: 0 });
    const middle = verticalCaretMove(empty, { position: 1, lineIndex: 0 }, 'down', goal);
    expect(middle).toEqual({ position: 1, lineIndex: 1 });
    expect(lineEdge(empty, middle, 'start')).toEqual({ position: 1, lineIndex: 1 });
    expect(verticalCaretMove(empty, middle, 'down', goal)).toEqual({
      position: 2,
      lineIndex: 2,
    });
  });

  it('keys the goal column by shape and visual line', () => {
    const goal = { shapeId: 'first', position: 5, lineIndex: 1 };
    expect(sameCaretGoalKey(goal, { ...goal })).toBe(true);
    expect(sameCaretGoalKey(goal, { ...goal, shapeId: 'second' })).toBe(false);
    expect(sameCaretGoalKey(goal, { ...goal, lineIndex: 0 })).toBe(false);
  });

  it('jumps whole words and stops at the text edges', () => {
    expect(wordBoundary('Hello brave world', 0, 1)).toBe(5);
    expect(wordBoundary('Hello brave world', 5, 1)).toBe(6);
    expect(wordBoundary('Hello brave world', 8, -1)).toBe(6);
    expect(wordBoundary('Hello brave world', 0, -1)).toBe(0);
    expect(wordBoundary('Hello brave world', 17, 1)).toBe(17);
  });
});

function stops(start: number, xs: number[]): Array<{ position: number; x: number }> {
  return xs.map((x, index) => ({ position: start + index, x }));
}
