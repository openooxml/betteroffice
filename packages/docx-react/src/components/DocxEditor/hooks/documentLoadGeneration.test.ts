import { describe, expect, test } from 'bun:test';
import { DocumentLoadGeneration } from './documentLoadGeneration';

describe('DocumentLoadGeneration', () => {
  test('reports layout errors before and after load completion without completing the load', () => {
    const loads = new DocumentLoadGeneration();
    const generation = loads.begin();
    const errors: Error[] = [];
    const report = (error: Error) => errors.push(error);
    const before = new Error('preflight');
    const after = new Error('pagination');

    loads.reportError(generation, before, report);
    expect(loads.complete(generation)).toBe(true);
    loads.reportError(generation, after, report);
    expect(errors).toEqual([before, after]);

    loads.begin();
    loads.reportError(generation, new Error('stale'), report);
    expect(errors).toEqual([before, after]);
  });

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
