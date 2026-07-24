import { describe, expect, test } from 'bun:test';
import { VerticalCaretGoal } from './verticalCaretGoal';

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
