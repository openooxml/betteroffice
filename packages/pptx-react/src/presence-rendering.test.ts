import { describe, expect, it } from 'bun:test';
import type { PptxPresencePeer } from '@betteroffice/pptx';
import {
  groupPresenceBySlide,
  groupShapePresence,
  limitPresence,
} from './presence-rendering';

function peer(
  clientId: number,
  slideId: string,
  shapeId?: string,
  cursorMovedAt = clientId
): PptxPresencePeer {
  return {
    state: {
      clientId,
      clock: 1,
      user: { name: `Peer ${clientId}`, color: '#3949AB' },
      cursor: { slideId, ...(shapeId ? { shapeId } : {}) },
    },
    lastSeen: cursorMovedAt,
    cursorMovedAt,
  };
}

describe('presence rendering bounds', () => {
  it('caps toolbar peers and reports the overflow', () => {
    const limited = limitPresence(Array.from({ length: 20 }, (_, index) => index), 3);
    expect(limited).toEqual({ visible: [0, 1, 2], overflow: 17 });
  });

  it('caps each slide bucket and ignores unknown slides', () => {
    const grouped = groupPresenceBySlide(
      [
        peer(1, 'slide-a'),
        peer(2, 'slide-a'),
        peer(3, 'slide-a'),
        peer(4, 'slide-b'),
        peer(5, 'unknown'),
      ],
      new Set(['slide-a', 'slide-b']),
      2
    );

    expect(grouped.get('slide-a')?.visible.map((entry) => entry.state.clientId)).toEqual([
      1,
      2,
    ]);
    expect(grouped.get('slide-a')?.overflow).toBe(1);
    expect(grouped.get('slide-b')?.visible.map((entry) => entry.state.clientId)).toEqual([4]);
    expect(grouped.has('unknown')).toBe(false);
  });

  it('merges peers on one shape and bounds distinct outlines', () => {
    const grouped = groupShapePresence(
      [
        peer(1, 'slide-a', 'shape-a', 10),
        peer(2, 'slide-a', 'shape-a', 20),
        peer(3, 'slide-a', 'shape-b'),
        peer(4, 'slide-a', 'shape-c'),
        peer(5, 'slide-b', 'shape-a'),
      ],
      'slide-a',
      new Set(['shape-a', 'shape-b', 'shape-c']),
      2
    );

    expect(grouped.visible.map(({ shapeId, peer: entry, count }) => ({
      shapeId,
      clientId: entry.state.clientId,
      count,
    }))).toEqual([
      { shapeId: 'shape-a', clientId: 2, count: 2 },
      { shapeId: 'shape-b', clientId: 3, count: 1 },
    ]);
    expect(grouped.overflow).toBe(1);
  });
});
