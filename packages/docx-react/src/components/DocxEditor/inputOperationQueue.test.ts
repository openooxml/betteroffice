import { describe, expect, test } from 'bun:test';
import { InputOperationQueue } from './inputOperationQueue';
import { VerticalCaretGoal } from './verticalCaretGoal';

describe('InputOperationQueue', () => {
  test('flush waits for accepted operations and rejects only for their failures', async () => {
    const errors: unknown[] = [];
    const queue = new InputOperationQueue((error) => errors.push(error));
    let release!: () => void;
    let applied = false;
    queue.enqueue(async () => {
      await new Promise<void>((resolve) => {
        release = resolve;
      });
      applied = true;
    });
    const flushed = queue.flush();
    await Promise.resolve();
    expect(applied).toBe(false);
    release();
    await flushed;
    expect(applied).toBe(true);
    const failure = new Error('input failed');
    queue.enqueue(() => {
      throw failure;
    });
    await expect(queue.flush()).rejects.toBe(failure);
    expect(errors).toEqual([failure]);
    queue.enqueue(() => {});
    await queue.flush();
  });

  test('a flush reports the first input failure since its checkpoint', async () => {
    const queue = new InputOperationQueue(() => {});
    const since = queue.failureCheckpoint();
    const first = new Error('first');
    const second = new Error('second');
    queue.enqueue(() => {
      throw first;
    });
    queue.enqueue(() => {
      throw second;
    });
    await expect(queue.flush(since)).rejects.toBe(first);
    const later = queue.failureCheckpoint();
    queue.enqueue(() => {
      throw second;
    });
    await expect(queue.flush(later)).rejects.toBe(second);
  });

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

  test('runs admitted operations in order and settles each with its own result', async () => {
    const queue = new InputOperationQueue(() => {});
    const order: string[] = [];
    let release!: () => void;
    queue.enqueue(
      () =>
        new Promise<void>((resolve) => {
          release = () => {
            order.push('input');
            resolve();
          };
        })
    );
    const command = queue.run(() => {
      order.push('command');
      return 'done';
    });
    queue.enqueue(() => {
      order.push('later input');
    });
    await Promise.resolve();
    expect(order).toEqual([]);
    release();
    expect(await command).toBe('done');
    await queue.idle();
    expect(order).toEqual(['input', 'command', 'later input']);
  });

  test('a failed command rejects alone; failed input fails only what was behind it', async () => {
    const errors: unknown[] = [];
    const queue = new InputOperationQueue((error) => errors.push(error));
    const refusal = new Error('command refused');
    await expect(
      queue.run(() => {
        throw refusal;
      })
    ).rejects.toBe(refusal);
    await queue.flush();
    const lost = new Error('input lost');
    queue.enqueue(() => {
      throw lost;
    });
    const behind = queue.run((inputLost) => inputLost);
    await queue.idle();
    expect(errors).toEqual([lost]);
    expect(await behind).toBe(true);
    expect(await queue.run((inputLost) => inputLost)).toBe(false);
  });

  test('reports when accepted work starts and stops waiting', async () => {
    const changes: boolean[] = [];
    const queue = new InputOperationQueue(
      () => {},
      (pending) => changes.push(pending)
    );
    expect(queue.hasPending()).toBe(false);
    queue.enqueue(() => {});
    void queue.run(() => {});
    expect(queue.hasPending()).toBe(true);
    await queue.idle();
    expect(queue.hasPending()).toBe(false);
    expect(changes).toEqual([true, false]);
  });
});
