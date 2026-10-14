import { describe, expect, test } from 'bun:test';
import { createInputCoordinator, PptxInputFailure, PptxInputStale } from './inputCoordinator';

function deferred() {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<void>((done, fail) => {
    resolve = done;
    reject = fail;
  });
  return { promise, resolve, reject };
}

describe('PPTX input coordinator', () => {
  test('applies work at once while nothing is pending', async () => {
    const coordinator = createInputCoordinator();
    const order: string[] = [];
    void coordinator.input(() => {
      order.push('key');
    });
    const result = coordinator.run(() => {
      order.push('command');
      return 7;
    });
    expect(order).toEqual(['key', 'command']);
    expect(coordinator.busy()).toBe(false);
    expect(await result).toBe(7);
  });

  test('orders commands and later input behind pending input', async () => {
    let changes = 0;
    const coordinator = createInputCoordinator(() => {
      changes += 1;
    });
    const order: string[] = [];
    const image = deferred();
    void coordinator.input(() => image.promise.then(() => void order.push('image')));
    const undo = coordinator.run(() => order.push('undo'));
    void coordinator.input(() => {
      order.push('typing');
    }, 'keyboard');
    expect(coordinator.busy()).toBe(true);
    expect(coordinator.keyboardQueued()).toBe(true);
    expect(order).toEqual([]);
    image.resolve();
    await undo;
    await Promise.resolve();
    expect(order).toEqual(['image', 'undo', 'typing']);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(coordinator.busy()).toBe(false);
    expect(changes).toBe(2);
  });

  test('fails commands admitted behind failed input, but not later ones', async () => {
    const coordinator = createInputCoordinator();
    const image = deferred();
    const imported = coordinator.input(() => image.promise);
    void imported.catch(() => {});
    const behind = coordinator.run(() => 'ran');
    const failure = new Error('decode failed');
    image.reject(failure);
    const error = await behind.catch((value: unknown) => value);
    expect(error).toBeInstanceOf(PptxInputFailure);
    expect((error as PptxInputFailure).cause).toBe(failure);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(await coordinator.run(() => 'later')).toBe('later');
  });

  test('refuses work queued for a replaced document', async () => {
    const coordinator = createInputCoordinator();
    const image = deferred();
    const order: string[] = [];
    void coordinator.input(() => image.promise);
    const typing = coordinator.input(() => {
      order.push('stale typing');
    }, 'keyboard');
    const stale = coordinator.run(() => order.push('stale command'));
    coordinator.reset();
    expect(coordinator.busy()).toBe(false);
    expect(coordinator.keyboardQueued()).toBe(false);
    expect(await coordinator.run(() => 'fresh')).toBe('fresh');
    image.resolve();
    expect(await typing.catch((error: unknown) => error)).toBeInstanceOf(PptxInputStale);
    expect(await stale.catch((error: unknown) => error)).toBeInstanceOf(PptxInputStale);
    expect(order).toEqual([]);
    expect(coordinator.busy()).toBe(false);
  });
});
