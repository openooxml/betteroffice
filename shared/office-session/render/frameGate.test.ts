import { expect, test } from 'bun:test';
import { createFrameGate, MAX_RECTS_PER_SURFACE } from './frameGate';
import type { ConsumedMessage } from './types';

const rect = { x: 0, y: 0, width: 4, height: 4 };
function consumed(epoch: number, status: ConsumedMessage['status'] = 'presented'): ConsumedMessage {
  return { session: 'session', incarnation: 1, epoch, status };
}

test('only one frame is outstanding; damage and geometry merge while credit is out', () => {
  const gate = createFrameGate();
  expect(gate.take()).toBeUndefined();
  gate.markDamage('page:1', rect);
  const first = gate.take()!;
  expect(gate.canPost()).toBe(false);
  gate.markDamage('page:1', rect);
  gate.markDamage('page:1', { ...rect, x: 6 });
  gate.markDamage('page:1', { ...rect, x: 3 });
  gate.markGeometry('page:1', 3);
  gate.markGeometry('page:1', 2);
  expect(gate.take()).toBeUndefined();
  gate.onConsumed(consumed(first.epoch + 1));
  expect(gate.canPost()).toBe(false);
  gate.onConsumed(consumed(first.epoch));
  const next = gate.take()!;
  expect(next.epoch).toBe(first.epoch + 1);
  expect(next.surfaces.get('page:1')).toEqual({ rects: [{ ...rect, width: 10 }], full: false, geometryRevision: 3 });
  expect(first.surfaces.get('page:1')).toEqual({ rects: [], full: true });
  gate.onConsumed(consumed(first.epoch));
  expect(gate.take()).toBeUndefined();
  gate.onConsumed(consumed(next.epoch));
  expect(gate.canPost()).toBe(true);
  expect(gate.take()).toBeUndefined();
});

test('discarded and superseded frames force full replacements and preserve changed geometry', () => {
  for (const status of ['discarded', 'superseded'] as const) {
    const gate = createFrameGate();
    gate.markDamage('pane:1', rect); gate.markGeometry('pane:1', 5);
    const first = gate.take()!;
    gate.markDamage('pane:1', { ...rect, x: 20 });
    gate.markGeometry('pane:1', 6);
    gate.onConsumed(consumed(first.epoch, status));
    expect(gate.take()!.surfaces.get('pane:1')).toEqual({ rects: [], full: true, geometryRevision: 6 });
  }
});

test('resync bypasses outstanding credit and marks specific or all known surfaces full', () => {
  const gate = createFrameGate();
  gate.markDamage('pane:1', rect); gate.markDamage('pane:2', rect);
  const first = gate.take()!;
  gate.onResync('pane:1');
  gate.markDamage('pane:1', rect);
  expect(gate.take()).toBeUndefined();
  gate.onConsumed(consumed(first.epoch));
  const second = gate.take()!;
  expect([...second.surfaces]).toEqual([['pane:1', { rects: [], full: true }]]);
  gate.onResync();
  gate.onConsumed(consumed(second.epoch));
  expect([...gate.take()!.surfaces.values()]).toEqual([{ rects: [], full: true }, { rects: [], full: true }]);
});

test('geometry alone can be taken and nonoverlapping damage stays separate', () => {
  const gate = createFrameGate();
  gate.markGeometry('page:1', 1);
  const first = gate.take()!;
  expect(first.surfaces.get('page:1')).toEqual({ rects: [], full: true, geometryRevision: 1 });
  gate.onConsumed(consumed(first.epoch));
  gate.markDamage('page:1', rect); gate.markDamage('page:1', { ...rect, x: 10 });
  gate.markDamage('page:1', { ...rect, width: 0 });
  expect(gate.take()!.surfaces.get('page:1')?.rects.length).toBe(2);
});

test('the fifth disjoint rectangle collapses pending damage to one union', () => {
  expect(MAX_RECTS_PER_SURFACE).toBe(4);
  for (const count of [MAX_RECTS_PER_SURFACE, MAX_RECTS_PER_SURFACE + 1]) {
    const gate = createFrameGate();
    gate.markDamage('page:1', 'full'); gate.onConsumed(consumed(gate.take()!.epoch));
    for (let index = 0; index < count; index += 1) gate.markDamage('page:1', { ...rect, x: index * 10 });
    const rects = gate.take()!.surfaces.get('page:1')!.rects;
    if (count === MAX_RECTS_PER_SURFACE) expect(rects.length).toBe(count);
    else expect(rects).toEqual([{ ...rect, width: 44 }]);
  }
});

test('base follows only presented frames and the initial frame fully replaces every surface', () => {
  for (const status of ['discarded', 'superseded'] as const) {
    const gate = createFrameGate();
    gate.markDamage('page:A', rect); gate.markDamage('page:B', rect);
    const first = gate.take()!;
    expect(first.baseEpoch).toBeUndefined();
    expect([...first.surfaces.values()]).toEqual([{ rects: [], full: true }, { rects: [], full: true }]);
    gate.onConsumed(consumed(first.epoch));
    gate.markDamage('page:A', rect);
    const delta = gate.take()!;
    expect(delta.baseEpoch).toBe(first.epoch);
    expect(delta.surfaces.get('page:A')).toEqual({ rects: [rect], full: false });
    gate.onConsumed(consumed(delta.epoch, status));
    const retry = gate.take()!;
    expect(retry.baseEpoch).toBe(first.epoch);
    expect(retry.surfaces.get('page:A')).toEqual({ rects: [], full: true });
    gate.onConsumed(consumed(retry.epoch));
    gate.markDamage('page:B', rect);
    const next = gate.take()!;
    expect(next.baseEpoch).toBe(retry.epoch);
    expect(next.surfaces.get('page:B')).toEqual({ rects: [rect], full: false });
  }
});
