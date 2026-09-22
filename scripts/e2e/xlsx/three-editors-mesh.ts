import { expect } from 'bun:test';

import { SHEET, exchange, fingerprint, replicas } from './context';
import type { XlsxScenario } from './context';

const ROW = 340;
const REGION = 'A341:F360';

export const threeEditorsMesh: XlsxScenario = {
  name: 'three-editors-mesh-with-undo',
  description: 'Three replicas edit in a full mesh, each undoes its own last change, and convergence is checked after every exchange round together with per-replica history depth.',
  participants: ['web:a', 'web:b', 'web:c'],
  run(ctx) {
    const peers = replicas(ctx, ['web:a', 'web:b', 'web:c']);
    const [a, b, c] = peers;
    const a1 = (row: number, col: number) => a.handle.cell(SHEET, row, col).a1;
    const converged = () => {
      const reference = fingerprint(a.handle, REGION);
      for (const peer of peers) expect(fingerprint(peer.handle, REGION)).toBe(reference);
      return reference;
    };

    peers.forEach((peer, index) => {
      expect(peer.timer.op('editCell:own', () => peer.handle.editCell(SHEET, ROW, index, `peer ${index}`)).applied).toBe(true);
      expect(peer.timer.op('editCells:block', () => peer.handle.editCells(SHEET, [
        { row: ROW + 1 + index, col: 0, input: String((index + 1) * 10) },
        { row: ROW + 1 + index, col: 1, input: `=${a1(ROW + 1 + index, 0)}*3` },
      ])).applied).toBe(true);
    });

    const delivered = exchange(peers);
    expect(delivered).toBeGreaterThan(0);
    const meshed = converged();
    expect(meshed).toContain('peer 0');
    expect(meshed).toContain('peer 2');

    for (const peer of peers) {
      expect(peer.timer.op('historyState', () => peer.handle.historyState()).undoDepth).toBe(2);
      expect(peer.timer.op('undo:own', () => peer.handle.undo()).applied).toBe(true);
    }
    exchange(peers, 'applyUpdate:afterUndo');
    const afterUndo = converged();
    expect(afterUndo).not.toBe(meshed);
    expect(afterUndo).toContain('peer 0');

    for (const peer of peers) {
      expect(peer.timer.op('historyState:afterUndo', () => peer.handle.historyState())).toMatchObject({ undoDepth: 1, redoDepth: 1 });
    }

    expect(a.timer.op('redo:own', () => a.handle.redo()).applied).toBe(true);
    exchange(peers, 'applyUpdate:afterRedo');
    converged();
    expect([...b.handle.encodeStateVector()].join()).toBe([...c.handle.encodeStateVector()].join());
  },
};
