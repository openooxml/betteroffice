import { describe, expect, test } from 'bun:test';
import { DocumentLoadGeneration } from './documentLoadGeneration';

describe('DocumentLoadGeneration', () => {
  test('rejects stale completion without resolving the current load', async () => {
    const loads = new DocumentLoadGeneration();
    const generationA = loads.begin();
    let resolvedA = false;
    void loads.waitForCompletion(generationA).then(() => {
      resolvedA = true;
    });

    const generationB = loads.begin();
    let resolvedB = false;
    const pendingB = loads.waitForCompletion(generationB).then(() => {
      resolvedB = true;
    });
    await Promise.resolve();

    expect(resolvedA).toBe(true);
    expect(loads.complete(generationA)).toBe(false);
    await Promise.resolve();
    expect(resolvedB).toBe(false);

    expect(loads.complete(generationB)).toBe(true);
    await pendingB;
    expect(resolvedB).toBe(true);
  });

  test('invalidates pending work on unmount', async () => {
    const loads = new DocumentLoadGeneration();
    const generation = loads.begin();
    const pending = loads.waitForCompletion(generation);

    loads.invalidate();

    await pending;
    expect(loads.isCurrent(generation)).toBe(false);
    expect(loads.complete(generation)).toBe(false);
  });
});
