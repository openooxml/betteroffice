import { describe, expect, test } from 'bun:test';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { DisplayListQueryEpochGate } from './displayListQueryEpochGate';

const queries = (pageIndex: number) =>
  ({
    displayList: { pages: [{ pageIndex, width: 1, height: 1, primitives: [] }] },
  }) as unknown as DisplayListQueries;

describe('DisplayListQueryEpochGate', () => {
  test('returns no facade while display-list queries are unavailable', async () => {
    const gate = new DisplayListQueryEpochGate();
    expect(await gate.resolve()).toBeNull();

    gate.invalidate();
    const pending = gate.resolve();
    gate.clear();
    expect(await pending).toBeNull();
    expect(await gate.resolve()).toBeNull();
  });

  test('waits for a fresh facade after invalidation', async () => {
    const gate = new DisplayListQueryEpochGate();
    const stale = { queries: queries(1), frameEpoch: 4 };
    const fresh = { queries: queries(2), frameEpoch: 5 };
    gate.publish(stale);
    gate.invalidate();

    const pending = gate.resolve();
    gate.publish(fresh);

    expect(await pending).toBe(fresh);
  });

  test('does not release a queued move before its resident frame epoch', async () => {
    const gate = new DisplayListQueryEpochGate();
    gate.publish({ queries: queries(5), frameEpoch: 5 });
    const expected = { queries: queries(7), frameEpoch: 7 };
    const pending = gate.resolve(7);
    let settled = false;
    void pending.then(() => {
      settled = true;
    });

    gate.publish({ queries: queries(6), frameEpoch: 6 });
    await Promise.resolve();
    expect(settled).toBe(false);

    gate.publish(expected);
    expect(await pending).toBe(expected);
  });
});
