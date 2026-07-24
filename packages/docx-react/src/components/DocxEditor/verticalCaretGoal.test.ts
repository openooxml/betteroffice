import { describe, expect, test } from 'bun:test';
import { paragraphVerticalMove, VerticalCaretGoal } from './verticalCaretGoal';

describe('VerticalCaretGoal', () => {
  test('retains goal X across consecutive vertical moves', () => {
    const goal = new VerticalCaretGoal();
    expect(goal.current()).toBeUndefined();
    goal.retain(172.5);
    expect(goal.current()).toBe(172.5);
    goal.retain(172.5);
    expect(goal.current()).toBe(172.5);
  });

  test('resets after a non-vertical caret action', () => {
    const goal = new VerticalCaretGoal();
    goal.retain(84);
    goal.reset();
    expect(goal.current()).toBeUndefined();
  });
});

describe('paragraphVerticalMove', () => {
  test('keeps header and footer navigation available without a display-list facade', () => {
    const paragraphs = [
      { paraId: 'first', length: 8 },
      { paraId: 'second', length: 3 },
    ];
    expect(
      paragraphVerticalMove(
        paragraphs,
        { story: 'hf:rIdHeader', paraId: 'first', offset: 6 },
        'down'
      )
    ).toEqual({ story: 'hf:rIdHeader', paraId: 'second', offset: 3 });
    expect(
      paragraphVerticalMove(
        paragraphs,
        { story: 'hf:rIdFooter', paraId: 'second', offset: 2 },
        'up'
      )
    ).toEqual({ story: 'hf:rIdFooter', paraId: 'first', offset: 2 });
  });
});
