import { describe, expect, it } from 'bun:test';

import {
  awarenessPeers,
  expireAwarenessRecords,
  reduceAwarenessEntries,
  reduceTypingInference,
  type AwarenessRecord,
} from './awareness';

const remote = {
  clientId: 12,
  clock: 1,
  state: {
    user: { name: 'Bright Fox', color: '#137333' },
    cursor: {
      story: 'body',
      anchor: Uint8Array.of(1),
      head: Uint8Array.of(2),
    },
  },
};

describe('awareness peer state', () => {
  it('accepts only larger clocks while duplicate messages refresh expiry', () => {
    const first = reduceAwarenessEntries(new Map(), [remote], 4, 100);
    const stale = reduceAwarenessEntries(
      first.records,
      [
        {
          ...remote,
          clock: 0,
          state: {
            ...remote.state,
            user: { ...remote.state.user, name: 'Stale Name' },
          },
        },
      ],
      4,
      200
    );
    const duplicate = reduceAwarenessEntries(stale.records, [remote], 4, 300);

    expect(awarenessPeers(duplicate.records)).toMatchObject([
      {
        clientId: 12,
        clock: 1,
        user: { name: 'Bright Fox' },
        lastSeenAt: 300,
      },
    ]);
    expect(stale.peersChanged).toBe(false);
    expect(duplicate.peersChanged).toBe(false);
  });

  it('keeps leave tombstones from reviving stale states', () => {
    const present = reduceAwarenessEntries(new Map(), [remote], 4, 100);
    const left = reduceAwarenessEntries(
      present.records,
      [{ clientId: 12, clock: 2, state: null }],
      4,
      200
    );
    const stale = reduceAwarenessEntries(left.records, [remote], 4, 300);

    expect(awarenessPeers(left.records)).toEqual([]);
    expect(awarenessPeers(stale.records)).toEqual([]);
    expect(left.peersChanged).toBe(true);
  });

  it('only refreshes the name flag timestamp when the cursor moves', () => {
    const first = reduceAwarenessEntries(new Map(), [remote], 4, 100);
    const heartbeat = reduceAwarenessEntries(
      first.records,
      [{ ...remote, clock: 2 }],
      4,
      200
    );
    const moved = reduceAwarenessEntries(
      heartbeat.records,
      [
        {
          ...remote,
          clock: 3,
          state: {
            ...remote.state,
            cursor: {
              ...remote.state.cursor,
              head: Uint8Array.of(3),
            },
          },
        },
      ],
      4,
      300
    );

    expect(awarenessPeers(heartbeat.records)[0]?.cursorMovedAt).toBe(100);
    expect(awarenessPeers(moved.records)[0]?.cursorMovedAt).toBe(300);
  });

  it('uses typing inference until a newer awareness clock arrives', () => {
    const present = reduceAwarenessEntries(new Map(), [remote], 4, 100);
    const inferred = reduceTypingInference(
      present.records,
      {
        clientId: 12,
        story: 'body',
        paraId: 'p1',
        endOffset: 8,
      },
      4,
      200
    );
    const newer = reduceAwarenessEntries(
      inferred.records,
      [{ ...remote, clock: 2 }],
      4,
      300
    );

    expect(awarenessPeers(inferred.records)[0]?.inferredCursor).toEqual({
      clientId: 12,
      story: 'body',
      paraId: 'p1',
      endOffset: 8,
    });
    expect(awarenessPeers(newer.records)[0]?.inferredCursor).toBeNull();
  });

  it('expires peers after the configured age', () => {
    const record: AwarenessRecord = {
      clientId: 12,
      clock: 1,
      state: remote.state,
      lastSeenAt: 100,
      cursorMovedAt: 100,
      inferredCursor: null,
    };
    const current = new Map([[12, record]]);

    expect(expireAwarenessRecords(current, 44_999, 45_000).records.size).toBe(1);
    const expired = expireAwarenessRecords(current, 45_100, 45_000);
    expect(expired.records.size).toBe(0);
    expect(expired.peersChanged).toBe(true);
  });
});
