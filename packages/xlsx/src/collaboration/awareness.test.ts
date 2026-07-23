import { describe, expect, it } from 'bun:test';
import {
  applyAwarenessUpdates,
  awarenessPeers,
  colorForClientId,
  expireAwarenessPeers,
  normalizeCollaborationUser,
  resolveAwarenessCursor,
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
  type AwarenessPeerStore,
} from './awareness';
import type { AwarenessUpdate } from './types';

const peer: AwarenessUpdate = {
  clientId: 22,
  clock: 1,
  state: {
    user: { name: 'Quiet Otter', color: '#0B57D0' },
    cursor: {
      sheet: 'sheet:0',
      anchor: { row: 4, col: 5 },
      head: { row: 4, col: 5 },
    },
  },
};

describe('awareness peer state', () => {
  it('accepts larger clocks, refreshes equal clocks, and ignores older clocks', () => {
    const store: AwarenessPeerStore = new Map();
    expect(applyAwarenessUpdates(store, [peer], 11, 100)).toBe(true);
    expect(awarenessPeers(store)[0]).toMatchObject({
      clientId: 22,
      clock: 1,
      lastSeen: 100,
      cursorMovedAt: 100,
    });

    expect(
      applyAwarenessUpdates(store, [{ ...peer, state: null }], 11, 200)
    ).toBe(false);
    expect(awarenessPeers(store)[0].lastSeen).toBe(200);
    expect(
      applyAwarenessUpdates(store, [{ ...peer, clock: 0, state: null }], 11, 300)
    ).toBe(false);
    expect(awarenessPeers(store)[0].lastSeen).toBe(200);

    const moved: AwarenessUpdate = {
      ...peer,
      clock: 2,
      state: {
        ...peer.state!,
        cursor: {
          sheet: 'sheet:0',
          anchor: { row: 8, col: 9 },
          head: { row: 10, col: 12 },
        },
      },
    };
    expect(applyAwarenessUpdates(store, [moved], 11, 400)).toBe(true);
    expect(awarenessPeers(store)[0]).toMatchObject({
      clock: 2,
      cursorMovedAt: 400,
      cursor: moved.state?.cursor,
    });
  });

  it('clears explicit leaves, retains their clocks, and expires stale entries', () => {
    const store: AwarenessPeerStore = new Map();
    applyAwarenessUpdates(store, [peer], 11, 0);
    expect(
      applyAwarenessUpdates(store, [{ clientId: 22, clock: 2, state: null }], 11, 1_000)
    ).toBe(true);
    expect(awarenessPeers(store)).toEqual([]);
    expect(
      applyAwarenessUpdates(store, [{ ...peer, clock: 1 }], 11, 2_000)
    ).toBe(false);
    expect(awarenessPeers(store)).toEqual([]);
    expect(expireAwarenessPeers(store, 45_999, 45_000)).toBe(false);
    expect(store.size).toBe(1);
    expect(expireAwarenessPeers(store, 46_000, 45_000)).toBe(false);
    expect(store.size).toBe(0);
  });

  it('expires live peers at the timeout and ignores the local client', () => {
    const store: AwarenessPeerStore = new Map();
    expect(applyAwarenessUpdates(store, [{ ...peer, clientId: 11 }], 11, 0)).toBe(false);
    expect(store.size).toBe(0);
    applyAwarenessUpdates(store, [peer], 11, 0);
    expect(expireAwarenessPeers(store, 44_999, 45_000)).toBe(false);
    expect(expireAwarenessPeers(store, 45_000, 45_000)).toBe(true);
    expect(awarenessPeers(store)).toEqual([]);
  });
});

describe('awareness identity and cursor resolution', () => {
  it('derives fixed-palette colors and normalizes host identity', () => {
    expect(colorForClientId(8)).toBe(colorForClientId(0));
    expect(normalizeCollaborationUser({ name: '  Swift Fox  ' }, 8)).toEqual({
      name: 'Swift Fox',
      color: colorForClientId(8),
    });
    expect(
      normalizeCollaborationUser({ name: 'Swift Fox', color: '#abcdef' }, 8)
    ).toEqual({ name: 'Swift Fox', color: '#ABCDEF' });
    expect(
      normalizeCollaborationUser({ name: 'Swift Fox', color: '#abc' }, 8)
    ).toEqual({ name: 'Swift Fox', color: '#AABBCC' });
  });

  it('normalizes ranges and drops unknown sheets or out-of-bounds cells', () => {
    const cursor = {
      sheet: 'sheet:1',
      anchor: { row: 8, col: 9 },
      head: { row: 2, col: 3 },
    };
    expect(resolveAwarenessCursor(cursor, ['sheet:0', 'sheet:1'])).toEqual({
      sheetIndex: 1,
      range: { top: 2, left: 3, bottom: 8, right: 9 },
    });
    expect(resolveAwarenessCursor(cursor, ['sheet:0'])).toBeNull();
    expect(
      resolveAwarenessCursor(
        { ...cursor, head: { row: XLSX_MAX_ROWS, col: 0 } },
        ['sheet:1']
      )
    ).toBeNull();
    expect(
      resolveAwarenessCursor(
        { ...cursor, head: { row: 0, col: XLSX_MAX_COLUMNS } },
        ['sheet:1']
      )
    ).toBeNull();
  });
});
