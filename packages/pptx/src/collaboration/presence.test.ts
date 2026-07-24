import { describe, expect, it } from 'bun:test';
import {
  PRESENCE_CLOCK_RETENTION_MS,
  PRESENCE_EXPIRY_MS,
  PresencePeers,
  presenceColorForClientId,
  type AwarenessUpdateEntry,
} from './presence';

function entry(
  clientId: number,
  clock: number,
  cursor: { slideId: string; shapeId?: string } | null = null
): AwarenessUpdateEntry {
  return {
    clientId,
    clock,
    state: {
      clientId,
      clock,
      user: { name: `Peer ${clientId}`, color: '#3949AB' },
      cursor,
    },
  };
}

describe('PresencePeers', () => {
  it('accepts only increasing clocks and keeps cursor movement time across heartbeats', () => {
    const peers = new PresencePeers(1);
    expect(peers.apply([entry(2, 3, { slideId: 's1', shapeId: 'a' })], 100)).toBe(true);
    expect(peers.apply([entry(2, 2, { slideId: 's1', shapeId: 'b' })], 200)).toBe(false);
    expect(peers.peers[0].state.clock).toBe(3);
    expect(peers.apply([entry(2, 4, { slideId: 's1', shapeId: 'a' })], 300)).toBe(true);
    const state = entry(2, 4, { slideId: 's1', shapeId: 'a' }).state;
    if (!state) throw new Error('Expected presence state');
    expect(peers.peers).toEqual([
      {
        state,
        lastSeen: 300,
        cursorMovedAt: 100,
      },
    ]);
  });

  it('tracks cursor moves, explicit leave tombstones, and local echo suppression', () => {
    const peers = new PresencePeers(1);
    peers.apply([entry(1, 1), entry(2, 1, { slideId: 's1' })], 100);
    peers.apply([entry(2, 2, { slideId: 's2', shapeId: 'b' })], 200);
    expect(peers.peers[0].cursorMovedAt).toBe(200);

    expect(peers.apply([{ clientId: 2, clock: 3, state: null }], 300)).toBe(true);
    expect(peers.peers).toEqual([]);
    expect(peers.apply([entry(2, 2, { slideId: 's1' })], 400)).toBe(false);
    expect(peers.peers).toEqual([]);
  });

  it('expires silent peers after 45 seconds', () => {
    const peers = new PresencePeers(1);
    peers.apply([entry(2, 1), entry(3, 1)], 1_000);
    peers.apply([entry(3, 2)], 20_000);

    expect(peers.expire(1_000 + PRESENCE_EXPIRY_MS - 1)).toBe(false);
    expect(peers.expire(1_000 + PRESENCE_EXPIRY_MS)).toBe(true);
    expect(peers.peers.map((peer) => peer.state.clientId)).toEqual([3]);
  });

  it('reclaims expired clocks after the replay grace period', () => {
    const peers = new PresencePeers(1);
    peers.apply([entry(2, 5)], 100);

    expect(peers.expire(100 + PRESENCE_EXPIRY_MS)).toBe(true);
    expect(peers.peers).toEqual([]);
    expect(peers.trackedIdCount).toBe(1);
    expect(peers.nextExpiryAt).toBe(100 + PRESENCE_CLOCK_RETENTION_MS);
    expect(peers.apply([entry(2, 4)], 100 + PRESENCE_EXPIRY_MS)).toBe(false);
    expect(peers.expire(100 + PRESENCE_CLOCK_RETENTION_MS - 1)).toBe(false);
    expect(peers.trackedIdCount).toBe(1);
    expect(peers.expire(100 + PRESENCE_CLOCK_RETENTION_MS)).toBe(false);
    expect(peers.trackedIdCount).toBe(0);
    expect(peers.nextExpiryAt).toBeUndefined();
    expect(peers.apply([entry(2, 1)], 101 + PRESENCE_CLOCK_RETENTION_MS)).toBe(true);
    expect(peers.peers[0].state.clock).toBe(1);
  });

  it('caps tracked ids by evicting the least recently seen first', () => {
    const peers = new PresencePeers(1, { maxTrackedIds: 3 });
    peers.apply([entry(2, 1)], 100);
    peers.apply([entry(3, 1)], 200);
    peers.apply([entry(4, 1)], 300);
    peers.apply([entry(2, 2)], 400);
    peers.apply([entry(5, 1)], 500);

    expect(peers.trackedIdCount).toBe(3);
    expect(peers.peers.map((peer) => peer.state.clientId)).toEqual([2, 4, 5]);
    for (let clientId = 6; clientId < 100; clientId += 1) {
      peers.apply([entry(clientId, 1)], 500 + clientId);
    }
    expect(peers.trackedIdCount).toBe(3);
    expect(peers.peers.map((peer) => peer.state.clientId)).toEqual([97, 98, 99]);
  });

  it('derives one of eight deterministic colors', () => {
    expect(presenceColorForClientId(42)).toBe(presenceColorForClientId(42));
    expect(new Set(Array.from({ length: 8 }, (_, index) => presenceColorForClientId(index))).size)
      .toBe(8);
  });
});
