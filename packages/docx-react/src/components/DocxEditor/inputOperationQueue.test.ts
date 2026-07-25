import { describe, expect, test } from 'bun:test';
import { InputOperationQueue } from './inputOperationQueue';
import { VerticalCaretGoal } from './verticalCaretGoal';

describe('InputOperationQueue', () => {
  test('orders a horizontal goal reset after an in-flight vertical move', async () => {
    const failures: unknown[] = [];
    const queue = new InputOperationQueue((error) => failures.push(error));
    const goal = new VerticalCaretGoal();
    let release!: () => void;
    const blocked = new Promise<void>((resolve) => {
      release = resolve;
    });

    queue.enqueue(async () => {
      await blocked;
      goal.retain(92);
    });
    queue.enqueue(() => goal.reset());
    release();
    await queue.idle();

    expect(failures).toEqual([]);
    expect(goal.current()).toBeUndefined();
  });

  test('abandons an awaited vertical move after a pointer selection', async () => {
    const failures: unknown[] = [];
    const queue = new InputOperationQueue((error) => failures.push(error));
    let selection = 'keyboard';
    let start!: () => void;
    let release!: () => void;
    const started = new Promise<void>((resolve) => {
      start = resolve;
    });
    const blocked = new Promise<void>((resolve) => {
      release = resolve;
    });
    const interactionEpoch = queue.captureInteractionEpoch();

    queue.enqueue(async () => {
      start();
      await blocked;
      if (!queue.isInteractionEpochCurrent(interactionEpoch)) return;
      selection = 'vertical';
    });
    await started;
    selection = 'pointer';
    queue.advanceInteractionEpoch();
    release();
    await queue.idle();

    expect(failures).toEqual([]);
    expect(selection).toBe('pointer');
  });
});
